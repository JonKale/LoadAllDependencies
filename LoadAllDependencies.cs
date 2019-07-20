using System;
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
        /// Gets the service provider from the owner package.
        /// </summary>
        // private IAsyncServiceProvider ServiceProvider => this.package;

        private DTE2 GetDTE2() => Package.GetGlobalService( typeof( DTE ) ) as DTE2;

        private SVsSolution Solution
        {
            get
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                return (SVsSolution)
                    new ServiceProvider( this.GetDTE2() as ComServiceProvider ).GetService( typeof( SVsSolution ) );
            }
        }

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
            var selectedProjectPath = this.GetSourceFilePath();
            var dependentProjectPaths = this.GetReferencedProjects( selectedProjectPath );
            var dependentProjectGuids = dependentProjectPaths.Select( p => this.GetProjectGuid( p ) ).ToArray();
            var solution = (IVsSolution4)this.Solution;
            for(var i = 0; i < dependentProjectGuids.Length; ++i)
            {
                ErrorHandler.ThrowOnFailure(solution.ReloadProject(ref dependentProjectGuids[i]));
            }
        }

        private string GetSourceFilePath()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var selectedItems = (Array)this.GetDTE2().ToolWindows.SolutionExplorer.SelectedItems;
            if (selectedItems == null)
            {
                return String.Empty;
            }

            foreach (UIHierarchyItem selItem in selectedItems)
            {
                if (selItem.Object is Project project)
                {
                    return project.FileName;
                }
            }

            return String.Empty;
        }

        IEnumerable<string> GetReferencedProjects( string projectPath, HashSet<string> accumulator = null )
        {
            accumulator = accumulator ?? new HashSet<string>();
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

            var solution = (IVsSolution)this.Solution;
            Assumes.Present( solution );
            _ = solution.GetProjectOfUniqueName( projectPath, out var hierarchy );
            if (hierarchy != null)
            {
                ErrorHandler.ThrowOnFailure(
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
