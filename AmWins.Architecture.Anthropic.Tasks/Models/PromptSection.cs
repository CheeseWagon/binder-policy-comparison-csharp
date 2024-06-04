using System.Text;

namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class PromptSection
{
    #region "Properties"
    public List<string>? Keys { get; set; }
    public ePromptSectionType? Type { get; set; }
    public eMessageRole Role { get; set; }
    public int? RequestedDocuments { get; set; }

    public string? Key { get; set; }
    public string? Value { get; set; }
    public bool? AppendResponseType { get; set; }
    public List<PromptQuestion>? Questions { get; set; }
    public List<string>? Text { get; set; }
    #endregion
    #region "Read-Only Properties"
    public string? RoleText
    {
        get
        {
            switch (Role)
            {
                case eMessageRole.User:
                    return "user";
                case eMessageRole.Assistant:
                    return "assistant";
                case eMessageRole.System:
                    return "system";
                default:
                    return null;
            }
        }
    }
    public string? QuestionsText
    {
        get
        {
            if (Questions == null || !Questions.Any()) return null;

            var builder = new StringBuilder();

            var templateValue = Value ?? string.Empty;

            builder.Append("{{lineFeed}}<questions>");
            foreach (var question in Questions)
            {
                builder.Append(Constants.LINE_FEED_TEMPLATE);
                var questionDescription = question.Description;

                if (AppendResponseType.GetValueOrDefault())
                {
                    if (string.Compare(question.Type, "value", StringComparison.InvariantCultureIgnoreCase) == 0
                        && templateValue.IndexOf("(please specify)") == 0)
                        templateValue += "(please specify)";
                    if (string.Compare(question.Type, "boolean", StringComparison.InvariantCultureIgnoreCase) == 0
                        && templateValue.IndexOf("(please respond with Yes or No)") == 0)
                        templateValue += "(please respond with Yes or No)";
                }
                builder.Append($"<question> {questionDescription} </question>");
            }
            builder.Append("{{lineFeed}}</questions>");

            var questionsText = builder.ToString();
            if (!string.IsNullOrWhiteSpace(questionsText))
            {
                templateValue = templateValue.Replace("{{QuestionsText}}", builder.ToString());
                builder = new StringBuilder();
            }

            templateValue = templateValue.Replace("{{lineFeed}}", Constants.LINE_FEED_TEMPLATE);
            if (!string.IsNullOrWhiteSpace(templateValue))
                builder.Append(templateValue);

            return builder.ToString();
        }
    }
    #endregion
}
