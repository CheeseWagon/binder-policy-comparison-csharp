using System.ComponentModel;

namespace AmWins.Architecture.Anthropic.Tasks;

#region "Enumeration: eClaudeImageType"
public enum eClaudeImageType
{
    None = 0,
    [Description("image/jpeg")]
    Jpeg,
    [Description("image/png")]
    Png,
    [Description("image/gif")]
    Gif,
    [Description("image/bmp")]
    Bmp,
    [Description("image/tiff")]
    Tiff
}
#endregion
#region "Enumeration: eMessageRole"
public enum eMessageRole
{
    User = 1,
    Assistant = 2,
    System = 3
}
#endregion
#region "Enumeration: ePromptSectionType"
public enum ePromptSectionType
{
    Text = 1,
    Images = 2,
    System = 3
}
#endregion

