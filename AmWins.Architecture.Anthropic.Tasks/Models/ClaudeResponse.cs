namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class ClaudeResponse
{
    #region "Constants"
    private const string QuestionAnswersPattern =
        @".*?<response>.*?(?<questionsJson>.*?)</response>";
    #endregion
    #region "Properties"
    public DateTime? DateOfExecution { get; set; }
    public string? PdfFile { get; set; }

    public List<ClaudePrompt>? PromptsExecuted { get; set; }
    public List<ClaudeDocument>? DocumentsProcessed { get; set; }
    public dynamic? Result { get; set; }
    #endregion
    #region "Read-Only Properties"
    public TimeSpan? Duration => new TimeSpan(PromptsExecuted?.Sum(p => p.Duration?.Ticks ?? 0) ?? 0);
    public decimal? PromptCost => PromptsExecuted?.Sum(p => p.Cost ?? 0);
    public decimal? ResponseCost => PromptsExecuted?.Sum(p => p.ResponseCost ?? 0);
    public decimal? TotalCost => PromptCost + ResponseCost;
    #endregion
}
