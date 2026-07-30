using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;

namespace AIHelper.VisualStudioExtension;

internal sealed class AIHelperSessionsCommand
{
    public const int CommandId = 0x0100;

    public static readonly Guid CommandSet = new(Guids.CommandSetGuidString);

    private readonly AsyncPackage _package;

    private AIHelperSessionsCommand(AsyncPackage package, OleMenuCommandService commandService)
    {
        _package = package ?? throw new ArgumentNullException(nameof(package));
        commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

        var menuCommandId = new CommandID(CommandSet, CommandId);
        var menuItem = new MenuCommand(Execute, menuCommandId);
        commandService.AddCommand(menuItem);
    }

    public static async System.Threading.Tasks.Task InitializeAsync(AsyncPackage package)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
        var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;

        if (commandService is not null)
        {
            _ = new AIHelperSessionsCommand(package, commandService);
        }
    }

    private void Execute(object sender, EventArgs e)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        _package.JoinableTaskFactory.RunAsync(async () =>
        {
            var window = await _package.ShowToolWindowAsync(
                typeof(AIHelperSessionsWindow),
                0,
                create: true,
                cancellationToken: _package.DisposalToken);

            if (window?.Frame is null)
            {
                throw new NotSupportedException("Cannot create AIHelper Sessions tool window.");
            }
        }).FileAndForget("AIHelperSessions/OpenToolWindow");
    }
}
