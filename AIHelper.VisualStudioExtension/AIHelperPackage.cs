using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;

namespace AIHelper.VisualStudioExtension;

[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
[InstalledProductRegistration("Сессии AIHelper", "Показывает сессии AIHelper внутри Visual Studio.", "1.0")]
[ProvideMenuResource("Menus.ctmenu", 1)]
[ProvideToolWindow(typeof(AIHelperSessionsWindow))]
[Guid(Guids.PackageGuidString)]
public sealed class AIHelperPackage : AsyncPackage
{
    protected override async System.Threading.Tasks.Task InitializeAsync(
        CancellationToken cancellationToken,
        IProgress<ServiceProgressData> progress)
    {
        await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
        await AIHelperSessionsCommand.InitializeAsync(this);
    }
}
