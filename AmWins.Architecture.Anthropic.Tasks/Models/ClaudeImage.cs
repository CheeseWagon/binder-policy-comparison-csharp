namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class ClaudeImage
{
    #region "Properties"
    public string? Key { get; set; }
    public eClaudeImageType? MediaType { get; set; }
    public string? Data { get; set; }
    #endregion
    #region "Read-Only Properties"
    public string? Type => "image";
    public string? Extension => MediaType switch
    {
        eClaudeImageType.Jpeg => "jpg",
        eClaudeImageType.Png => "png",
        eClaudeImageType.Gif => "gif",
        eClaudeImageType.Bmp => "bmp",
        eClaudeImageType.Tiff => "tiff",
        _ => null
    };
    public string? ContentType => MediaType switch
    {
        eClaudeImageType.Jpeg => "image/jpeg",
        eClaudeImageType.Png => "image/png",
        eClaudeImageType.Gif => "image/gif",
        eClaudeImageType.Bmp => "image/bmp",
        eClaudeImageType.Tiff => "image/tiff",
        _ => null
    };
    #endregion
}
