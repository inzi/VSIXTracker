using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms.Design;
using Task = System.Threading.Tasks.Task;

namespace VSIXTracker
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ToggleTracking
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;
        public const int CommandId2 = 0x0101;

        private OleMenuCommand menuItem;
        private OleMenuCommand menuItem2;


        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("2b6b9b93-c7fa-4486-9f7c-f6fdd6299e74");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ToggleTracking"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ToggleTracking(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            menuItem = new OleMenuCommand(this.Execute, menuCommandID);
            menuItem.BeforeQueryStatus += MenuCommand_BeforeQueryStatus;
            commandService.AddCommand(menuItem);

            var menuCommandID2 = new CommandID(CommandSet, CommandId2);
            menuItem2 = new OleMenuCommand(this.Execute, menuCommandID2);
            menuItem2.BeforeQueryStatus += MenuCommand_BeforeQueryStatus;
            commandService.AddCommand(menuItem2);

        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ToggleTracking Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ToggleTracking's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ToggleTracking(package, commandService);
        }


        private bool IsTrackingActive()
        {
            var dte = (DTE)Package.GetGlobalService(typeof(SDTE));

            // Get the Properties collection of the DTE object.
            Properties properties = dte.Properties["Environment", "ProjectsAndSolution"];

            // Find the "Track Active Item in Solution Explorer" property.
            Property trackActiveItemPropertyOption = properties.Item("TrackFileSelectionInExplorer");

            if (trackActiveItemPropertyOption != null)
            {
                return (bool)trackActiveItemPropertyOption.Value;
            }
            return false;
        }
        private void ToggleTrackingActive()
        {
            var dte = (DTE)Package.GetGlobalService(typeof(SDTE));
            Properties properties = dte.Properties["Environment", "ProjectsAndSolution"];
            Property trackActiveItemPropertyOption = properties.Item("TrackFileSelectionInExplorer");
            if (trackActiveItemPropertyOption != null)
            {
                trackActiveItemPropertyOption.Value = !(bool)trackActiveItemPropertyOption.Value;
            }
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {

            var dte = (DTE)Package.GetGlobalService(typeof(SDTE));

            var myCommand = sender as OleMenuCommand;
            if (myCommand == null)
            {
                return;
            }

            ToggleTrackingActive();
           
        }

        private void MenuCommand_BeforeQueryStatus(object sender, EventArgs e)
        {
            var command = sender as OleMenuCommand;
            if (command.CommandID.ID == CommandId)
            {
                command.Visible = !IsTrackingActive();

            }
            else if (command.CommandID.ID == CommandId2)
            {
                command.Visible = IsTrackingActive();

            }
          
        }
    }
}
