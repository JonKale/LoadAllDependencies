﻿using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Xml.Linq;
using System.IO;

using ComServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;
using Task = System.Threading.Tasks.Task;
using IAsyncServiceProvider = Microsoft.VisualStudio.Shell.IAsyncServiceProvider;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace LoadDependencies
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class LoadAllDependencies
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid( "a24ce2d6-6871-4350-9936-0fb86e1ffa7d" );

        /// <summary>
        /// Output window pane title and ID
        /// </summary>
        const string dependencyLoadFailuresPaneTitle = "Project Dependency Load Failures";
        public static readonly Guid dependencyLoadFailuresPaneId = new Guid( "ef47923b-e63a-4a95-a12c-d90e0126e5a3" );

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="LoadAllDependencies"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private LoadAllDependencies( AsyncPackage package, OleMenuCommandService commandService )
        {
            this.package = package ?? throw new ArgumentNullException( nameof( package ) );
            commandService = commandService ?? throw new ArgumentNullException( nameof( commandService ) );

            var menuCommandID = new CommandID( CommandSet, CommandId );
            var menuItem = new MenuCommand( this.Execute, menuCommandID );
            commandService.AddCommand( menuItem );
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static LoadAllDependencies Instance { get; private set; }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync( AsyncPackage package )
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync( package.DisposalToken );

            var commandService = await package.GetServiceAsync( typeof( IMenuCommandService ) ) as OleMenuCommandService;
            Instance = new LoadAllDependencies( package, commandService );
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute( object sender, EventArgs e )
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var errorMessage = String.Empty;
            var project = String.Empty;
            var errors = new[] { new { project, errorMessage } }.ToList();

            var (projectId, selectedProjectPath) = this.GetSelectedProject();
            if (projectId == Guid.Empty)
            {
                return;
            }

            var dependentProjectPaths = this.GetReferencedProjects( selectedProjectPath );
            var dependentProjects = dependentProjectPaths.Select( p => new { path = p, guid = this.GetProjectGuid( p ) } )
                                                         .Where( p => p.guid != Guid.Empty );

            var solution = (IVsSolution4)Package.GetGlobalService( typeof( SVsSolution ) );
            Assumes.Present( solution );

            // "If the project was not previously unloaded, then this method does nothing and returns S_FALSE." - MSDN
            // I'd love to know a better way of determining whether a project is loaded or not
            try
            {
                _ = ErrorHandler.ThrowOnFailure( solution.ReloadProject( ref projectId ) );
            }
            catch( Exception ex )
            {
                var message = ex.Message;
                while (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                    message += $"\r\n{ex.Message}";
                }

                errors.Add( new { project = selectedProjectPath, errorMessage = message } );
            }

            foreach ( var dependentProject in dependentProjects )
            {
                var guid = dependentProject.guid;
                try
                {
                    _ = ErrorHandler.ThrowOnFailure( solution.ReloadProject( ref guid ) );
                }
                catch (Exception ex)
                {
                    var message = ex.Message;
                    while (ex.InnerException != null)
                    {
                        ex = ex.InnerException;
                        message += $"\r\n{ex.Message}";
                    }

                    errors.Add( new { project = dependentProject.path, errorMessage = message } );
                }
            }

            errors = errors.Skip( 1 ).ToList();
            if ( errors.Count > 0 )
            {
                var outputWindow = (IVsOutputWindow)Package.GetGlobalService( typeof( SVsOutputWindow ) );
                Assumes.Present( solution );

                var guid = dependencyLoadFailuresPaneId;
                _ = ErrorHandler.ThrowOnFailure( 
                    outputWindow.CreatePane( ref guid, dependencyLoadFailuresPaneTitle, 1, 1 ) );
                _ = ErrorHandler.ThrowOnFailure( 
                    outputWindow.GetPane( ref guid, out var dependencyLoadFailurePane ) );

                try
                {
                    foreach ( var error in errors )
                    {
                        _ = ErrorHandler.ThrowOnFailure(
                            dependencyLoadFailurePane.OutputString(
                                $"Failed to load {error.project}\r\n{error.errorMessage}\r\n" ) );
                    }

                    _ = ErrorHandler.ThrowOnFailure( dependencyLoadFailurePane.Activate() );
                }
                finally
                {
                    _ = Marshal.ReleaseComObject( dependencyLoadFailurePane );
                }

                var message = 
                    errors[0].project == selectedProjectPath ? 
                        errors.Count == 1 ? 
                            "The selected project failed to load" :
                            "This project and one of more of its dependencies failed to load" : 
                        "One or more of this project's dependencies failed to load";

                _ = VsShellUtilities.ShowMessageBox(
                    this.package,
                    $"{message}. See the Output window for details.",
                    "Error loading project or dependencies",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST );
            }
        }

        private (Guid, string) GetSelectedProject()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var monitorSelection = (IVsMonitorSelection)Package.GetGlobalService( typeof( SVsShellMonitorSelection ) );
            Assumes.Present( monitorSelection );

            _ = ErrorHandler.ThrowOnFailure(
                monitorSelection.GetCurrentSelection(
                    out var hierarchyPtr, out var projectItemId, out var multiItemSelectPtr, out var selectionContainerPtr ) );
            try
            {
                if (Marshal.GetTypedObjectForIUnknown( hierarchyPtr, typeof( IVsHierarchy ) ) is IVsHierarchy hierarchy)
                {
                    var solution = (IVsSolution)Package.GetGlobalService( typeof( SVsSolution ) );
                    Assumes.Present( solution );

                    _ = ErrorHandler.ThrowOnFailure( solution.GetSolutionInfo( out var directory, out _, out _ ) );
                    _ = ErrorHandler.ThrowOnFailure( solution.GetGuidOfProject( hierarchy, out var projectGuid ) );
                    _ = ErrorHandler.ThrowOnFailure( solution.GetUniqueNameOfProject( hierarchy, out var uniqueName ) );
                    var projectPath = Path.GetFullPath( Path.Combine( directory, uniqueName ) );

                    return ( projectGuid, projectPath );
                }
            }
            finally
            {
                _ = Marshal.Release( hierarchyPtr );
                _ = Marshal.Release( selectionContainerPtr );
                if (multiItemSelectPtr != null)
                {
                    _ = Marshal.ReleaseComObject( multiItemSelectPtr );
                }
            }

            return ( Guid.Empty, String.Empty );
        }

        IEnumerable<string> GetReferencedProjects( string projectPath, HashSet<string> accumulator = null )
        {
            accumulator = accumulator ?? new HashSet<string>();

            /// Note that this assumes that project files keep their references in /Project/ItemGroup/ProjectReference 
            /// nodes. This function will need extending to handle other reference xpaths, but I've not seen any yet 
            /// (currently works for .csproj, .vcxproj, .sqlproj, fsproj and .vbproj)
            var project = XDocument.Load( projectPath );
            var projectReferences = project.Root
                                           .Elements().Where( e => e.Name.LocalName == "ItemGroup" )
                                           .Elements().Where( e => e.Name.LocalName == "ProjectReference" );
            foreach (var path in projectReferences.Select(
                r => Path.GetFullPath(
                        Path.Combine( Path.GetDirectoryName( projectPath ), r.Attribute( "Include" ).Value ) ) ))
            {
                if (accumulator.Add( path ))
                {
                    this.GetReferencedProjects( path, accumulator );
                }
            }

            return accumulator;
        }

        private Guid GetProjectGuid( string projectPath )
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var solution = (IVsSolution)Package.GetGlobalService( typeof( SVsSolution ) );
            Assumes.Present( solution );

            _ = ErrorHandler.ThrowOnFailure( solution.GetProjectOfUniqueName( projectPath, out var hierarchy ) );
            if (hierarchy != null)
            {
                _ = ErrorHandler.ThrowOnFailure(
                    hierarchy.GetGuidProperty(
                        VSConstants.VSITEMID_ROOT, (int)__VSHPROPID.VSHPROPID_ProjectIDGuid, out var projectGuid ) );

                if (projectGuid != null)
                {
                    return projectGuid;
                }
            }

            return Guid.Empty;
        }
    }
}
