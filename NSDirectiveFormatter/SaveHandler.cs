using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Commanding;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Editor.Commanding;
using Microsoft.VisualStudio.Text.Editor.Commanding.Commands;
using Microsoft.VisualStudio.Utilities;
using UsingDirectiveFormatter.Commands;

namespace UsingDirectiveFormatter
{
	// follows implementation from https://github.com/madskristensen/CodeCleanupOnSave/blob/master/src/SaveHandler.cs

	[Export(typeof(ICommandHandler))]
	[Name(nameof(SaveHandler))]
	[ContentType("csharp")]
	[TextViewRole(PredefinedTextViewRoles.PrimaryDocument)]
	public class SaveHandler : ICommandHandler<SaveCommandArgs>
	{
		public string DisplayName => nameof(SaveHandler);
		public CommandState GetCommandState(SaveCommandArgs args) => CommandState.Available;

		private readonly IEditorCommandHandlerServiceFactory commandService;

		[ImportingConstructor]
		public SaveHandler(IEditorCommandHandlerServiceFactory commandService)
		{
			this.commandService = commandService;
		}

		public bool ExecuteCommand(SaveCommandArgs args, CommandExecutionContext executionContext)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			DTE2 dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
			Assumes.Present(dte);

			// Check if file is part of a project first.
			if(args.SubjectBuffer.Properties.TryGetProperty(typeof(ITextDocument), out ITextDocument textDoc)) {
				var filePath = textDoc.FilePath;

				ProjectItem item = dte.Solution?.FindProjectItem(filePath);

				if(string.IsNullOrEmpty(item?.ContainingProject?.FullName)) {
					return true;
				}
			}

			var options = FormatCommandPackage.Instance.GetOptions();

			// Then check if it's been enabled in the options.
			if(!options.FormatOnSave) {
				return true;
			}

			try {
				args.SubjectBuffer.Format(options);
			}catch(Exception ex) {
				System.Diagnostics.Debug.WriteLine(ex);
			}

			return true;
		}
	}
}
