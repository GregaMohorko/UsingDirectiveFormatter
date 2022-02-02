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
	internal sealed class FormatCommand
	{
		public static FormatCommand Instance { get; private set; }
		public static void Initialize(FormatCommandPackage package)
		{
			Instance = new FormatCommand(package);
		}

		private DTE2 Dte { get; set; }
		private Document Document { get; set; }

		private readonly FormatCommandPackage package;

		/// <summary>
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		private FormatCommand(FormatCommandPackage package)
		{
			this.package = package;

			var serviceProvider = this.package as IServiceProvider;
			Assumes.Present(serviceProvider);

			if(serviceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService) {
				var menuCommandID = new CommandID(PackageGuids.guidFormatCommandPackageCmdSet, PackageIds.FormatCommandId);
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
		private void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var command = (OleMenuCommand)sender;

			Document = Dte.ActiveDocument;

			if(Document == null
				|| !Document.IsCSharpCode()
				) {
				command.Visible = false;
				command.Enabled = false;
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
		private void MenuItemCallback(object sender, EventArgs e)
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
