using Microsoft.VisualStudio.Shell;

namespace AIHelper.VisualStudioExtension;

public sealed class AIHelperSessionsWindow : ToolWindowPane
{
    public AIHelperSessionsWindow() : base(null)
    {
        Caption = "AIHelper Sessions";
        Content = new AIHelperSessionsControl();
    }
}
