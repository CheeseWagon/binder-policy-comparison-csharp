namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class ComparisonTopic
{
    #region "Properties"
    public string? Type { get; set; }
    public string? Text { get; set; }
    public string? Key { get; set; }
    public int? BinderTotal { get; set; }
    public int? PolicyTotal { get; set; }
    public int? DiscrepancyCount { get; set; }
    public int? DiscrepancyMatchedActualCount { get; set; }
    public object? Value { get; set; }
    public dynamic? Result { get; set; }

    public List<ClaudeDocument>? Documents { get; set; }
    #endregion
}
