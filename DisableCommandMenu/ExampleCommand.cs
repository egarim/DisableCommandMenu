//-----------------------------------------------------------------------
// <copyright file="C:\Users\Joche\source\repos\DisableCommandMenu\DisableCommandMenu\ExampleCommand.cs" company="BitFwks">
//     Author:  
//     Copyright (c) BitFwks. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace DisableCommandMenu
{
    /// <summary>
    /// Command handler
    /// </summary>
    sealed class ExampleCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("66ac74b3-74ab-4fce-82f8-87f7a393f76e");
        OleMenuCommand menuItem;

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExampleCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        ExampleCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            global::System.ComponentModel.Design.CommandID menuCommandID = new CommandID(CommandSet, CommandId);
            menuItem = new OleMenuCommand(Execute, menuCommandID);
            menuItem.BeforeQueryStatus += MenuItem_BeforeQueryStatus;
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string message = string.Format(CultureInfo.CurrentCulture,
                                           "Inside {0}.MenuItemCallback()",
                                           GetType().FullName);
            string title = nameof(ExampleCommand);

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(package,
                                            message,
                                            title,
                                            OLEMSGICON.OLEMSGICON_INFO,
                                            OLEMSGBUTTON.OLEMSGBUTTON_OK,
                                            OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand myCommand = sender as OleMenuCommand;
            if(null != myCommand)
            {
                myCommand.Text = "New Text";
            }
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ExampleCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ExampleCommand(package, commandService);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ExampleCommand Instance
        {
            get;
            private set;
        }
    }
}
