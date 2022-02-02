using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using UsingDirectiveFormatter.Options;

namespace UsingDirectiveFormatter.Commands
{
	[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
	[InstalledProductRegistration(Vsix.Name, Vsix.Description, Vsix.Version, IconResourceID = 400)] // Info on this package for Help/About
	[ProvideOptionPage(typeof(FormatOptionGrid), Vsix.Name, "Options", 0, 0, true)]
	[ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
	[Guid(PackageGuids.guidFormatCommandPackageString)]
	[ProvideMenuResource("Menus.ctmenu", 1)]
	public sealed class FormatCommandPackage : AsyncPackage
	{
		public static FormatCommandPackage Instance { get;private set; }

		protected override async System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
		{
			Instance = this;

			FormatCommand.Initialize(this);

			await base.InitializeAsync(cancellationToken, progress);
		}

		internal FormatOptionGrid GetOptions()
		{
			return (FormatOptionGrid)GetDialogPage(typeof(FormatOptionGrid));
		}
	}
}
