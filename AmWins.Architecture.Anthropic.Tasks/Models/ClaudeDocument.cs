namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class ClaudeDocument
{
    #region "Properties"
    public string? NamedInsured { get; set; }
    public string? Term { get; set; }
    public string? Lob { get; set; }
    public string? DocumentNumber { get; set; }
    public string? PolicyNumber { get; set; }
    public string? DocumentType { get; set; }
    public string? ExtractedText { get; set; }
    public string? FileName { get; set; }
    public int? SubmissionFileMarketId { get; set; }

    public int? SubmissionFileId { get; set; }

    public List<ClaudeImage>? PageImages { get; set; }
    public List<ComparisonTopic>? Topics { get; set; }
    public List<ComparisonTopic>? NotApplicableTopics { get; set; }
    public List<Discrepancy>? Discrepancies { get; set; }

    public dynamic? Values { get; set; }
    #endregion
}
