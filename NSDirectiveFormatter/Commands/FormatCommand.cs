//------------------------------------------------------------------------------
// <copyright file="FormatCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------
using System;
using System.ComponentModel.Design;
using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;

namespace UsingDirectiveFormatter.Commands
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class FormatCommand
	{
		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 0x0100;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("2bb0aad7-e323-43dc-883c-cab65d5684c7");

		public static FormatCommand Instance { get; private set; }
		/// <summary>
		/// Initializes the singleton instance of the command.
		/// </summary>
		public static void Initialize(FormatCommandPackage package)
		{
			Instance = new FormatCommand(package);
		}

		/// <summary>
		/// The DTE
		/// </summary>
		private DTE2 Dte { get; set; }

		/// <summary>
		/// The document
		/// </summary>
		private Document Document { get; set; }

		private readonly FormatCommandPackage package;

		/// <summary>
		/// Initializes a new instance of the <see cref="FormatCommand"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		private FormatCommand(FormatCommandPackage package)
		{
			this.package = package;

			var serviceProvider = this.package as IServiceProvider;
			Assumes.Present(serviceProvider);

			if(serviceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService) {
				var menuCommandID = new CommandID(CommandSet, CommandId);
				var menuItem = new OleMenuCommand(MenuItemCallback, menuCommandID);
				menuItem.BeforeQueryStatus += MenuItem_BeforeQueryStatus;
				commandService.AddCommand(menuItem);
			}

			Dte = serviceProvider.GetService(typeof(SDTE)) as DTE2;
			Assumes.Present(Dte);
		}

		/// <summary>
		/// Handles the BeforeQueryStatus event of the MenuItem control.
		/// </summary>
		/// <param name="sender">The source of the event.</param>
		/// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
		private void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var command = (OleMenuCommand)sender;

			command.Visible = false;
			command.Enabled = false;

			Document = Dte.ActiveDocument;

			if(Document == null
				|| !Document.IsCSharpCode()
				) {
				return;
			}

			command.Visible = true;
			command.Enabled = true;
		}

		/// <summary>
		/// This function is the callback used to execute the command when the menu item is clicked.
		/// See the constructor to see how the menu item is associated with this function using
		/// OleMenuCommandService service and MenuCommand class.
		/// </summary>
		/// <param name="sender">Event sender.</param>
		/// <param name="e">Event args.</param>
		private void MenuItemCallback(object sender, EventArgs e)
		{
			Execute();
		}

		public void Execute()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var options = package.GetOptions();

			Document
				.ToIWpfTextView(Dte)
				.TextBuffer
				.Format(options);
		}
	}
}
