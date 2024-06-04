using ns = Newtonsoft.Json;
using sts = System.Text.Json.Serialization;

namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class PromptQuestion
{
    #region "Properties"
    [ns.JsonProperty("questionNumber")]
    [sts.JsonPropertyName("questionNumber")]
    public int? QuestionNumber { get; set; }

    [ns.JsonProperty("questionKey")]
    [sts.JsonPropertyName("questionKey")]
    public string? Key { get; set; }

    [ns.JsonProperty("questionText")]
    [sts.JsonPropertyName("questionText")]
    public string? Description { get; set; }

    public string? Type { get; set; }

    public string? Answer { get; set; }
    #endregion
}
