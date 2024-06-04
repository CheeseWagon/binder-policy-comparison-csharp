using System.Text;
using System.Text.RegularExpressions;

using Newtonsoft.Json;

using ns = Newtonsoft.Json;
using sts = System.Text.Json.Serialization;

namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class ClaudePrompt : KeyedItem
{
    #region "Constants"
    private const string QuestionAnswersPattern =
        @".*?<response>.*?(?<questionsJson>.*?)</response>";
    #endregion
    #region "Member Variables"
    private string? _promptText;
    private TimeSpan? _duration;
    #endregion
    #region "Properties"
    public string? Name { get; set; }
    public string? Description { get; set; }
    public string? System { get; set; }
    public bool? IsEnabled { get; set; }
    public string? Status { get; set; }
    public string? TopicsFile { get; set; }
    public TimeSpan? Duration 
    { 
        get
        {
            if (_duration.HasValue) return _duration;
            if (ModelsUsed == null || !ModelsUsed.Any()) return null;

            return new TimeSpan(ModelsUsed.Sum(m => m.Duration?.Ticks ?? 0));
        } 
        set => _duration = value; 
    }

    public decimal? Cost { get; set; }
    public decimal? ResponseCost { get; set; }

    [ns.JsonIgnore()]
    [sts.JsonIgnore()]
    public string? PromptText
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(_promptText)) return _promptText;
            if (Sections == null || !Sections.Any()) return null;

            var builder = new StringBuilder(Constants.HUMAN_PROMPT);
            foreach (var section in Sections)
            {
                builder.Append(Constants.LINE_FEED_TEMPLATE);
                var sectionText = section.Value;
                if (string.IsNullOrWhiteSpace(sectionText)) continue;

                if (!string.IsNullOrWhiteSpace(QuestionsFormat))
                    sectionText = sectionText.Replace("{{QuestionsFormat}}", QuestionsFormat);
                if (!string.IsNullOrWhiteSpace(JsonFormat))
                    sectionText = sectionText.Replace("{{JsonFormat}}", JsonFormat);
                if (!string.IsNullOrWhiteSpace(ExtractedText))
                    sectionText = sectionText.Replace("{{ExtractedText}}", ExtractedText);
                if (!string.IsNullOrWhiteSpace(section.QuestionsText))
                    sectionText = section.QuestionsText;
                builder.Append(sectionText);
            }
            builder.Append(Constants.ASSISTANT_PROMPT);

            return builder.ToString().Replace("{{lineFeed}}", Constants.LINE_FEED_TEMPLATE);
        }
        set => _promptText = value;
    }

    public dynamic? Result { get; set; }
    public dynamic? ResponseFromAI { get; set; }

    public List<ClaudeModel>? ModelsUsed { get; set; }
    public List<PromptSection>? Sections { get; set; }
    #endregion
    #region "Obsolete Properties"
    public string? ExtractedText { get; set; }
    public string? JsonFormat { get; set; }
    public string? QuestionsFormat { get; set; }
    public List<PromptQuestion>? Questions
    {
        get
        {
            try
            {
                if (Result == null) return null;

                var completionResult = Result["completion"];
                if (completionResult == null) return null;

                object test = completionResult;
                string completionString = JsonConvert.SerializeObject(test);
                var isValidMatch = Regex.Match(completionString, QuestionAnswersPattern);
                if (!isValidMatch.Success) return null;

                var questionsJson = isValidMatch.Groups["questionsJson"].Value;
                if (string.IsNullOrWhiteSpace(questionsJson)) return null;

                questionsJson = questionsJson.Replace("\r\n", string.Empty);
                questionsJson = questionsJson.Replace("\\n", string.Empty);
                questionsJson = questionsJson.Replace("\\\"", "\"");
                if (string.IsNullOrWhiteSpace(questionsJson)) return null;

                var resultQuestions = JsonConvert.DeserializeObject<dynamic>(questionsJson);
                var questionsArray = resultQuestions?["questions"];
                var questionsString = JsonConvert.SerializeObject(questionsArray);
                List<PromptQuestion> actualQuestions = JsonConvert.DeserializeObject<List<PromptQuestion>>(questionsString);

                return actualQuestions == null || (!actualQuestions?.Any() ?? true) ? null : actualQuestions;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
    #endregion
}
