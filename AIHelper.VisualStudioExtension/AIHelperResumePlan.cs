namespace AIHelper.VisualStudioExtension;

public sealed class AIHelperResumePlan
{
    private AIHelperResumePlan(
        string toolName,
        string workingDirectory,
        string target,
        string executable,
        string[] arguments,
        string description,
        bool isDeepLink,
        string clipboardPrompt)
    {
        ToolName = toolName;
        WorkingDirectory = workingDirectory;
        Target = target;
        Executable = executable;
        Arguments = arguments;
        Description = description;
        IsDeepLink = isDeepLink;
        ClipboardPrompt = clipboardPrompt;
    }

    public string ToolName { get; }

    public string WorkingDirectory { get; }

    public string Target { get; }

    public string Executable { get; }

    public string[] Arguments { get; }

    public string Description { get; }

    public bool IsDeepLink { get; }

    public string ClipboardPrompt { get; }

    public static AIHelperResumePlan Command(
        string toolName,
        string workingDirectory,
        string executable,
        string[] arguments,
        string commandPreview,
        string description)
    {
        return new AIHelperResumePlan(
            toolName,
            workingDirectory,
            commandPreview,
            executable,
            arguments,
            description,
            false,
            string.Empty);
    }

    public static AIHelperResumePlan CommandWithClipboardPrompt(
        string toolName,
        string workingDirectory,
        string executable,
        string[] arguments,
        string commandPreview,
        string clipboardPrompt,
        string description)
    {
        return new AIHelperResumePlan(
            toolName,
            workingDirectory,
            commandPreview,
            executable,
            arguments,
            description,
            false,
            clipboardPrompt);
    }

    public static AIHelperResumePlan DeepLink(
        string toolName,
        string target,
        string description)
    {
        return new AIHelperResumePlan(
            toolName,
            string.Empty,
            target,
            string.Empty,
            [],
            description,
            true,
            string.Empty);
    }
}
