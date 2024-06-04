namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class DocumentComparisonRequest
{
    #region "Properties"
    public bool? SubmitRequestToClaude { get; set; }
    public List<ClaudeDocument>? Documents { get; set; }
    #endregion
}
