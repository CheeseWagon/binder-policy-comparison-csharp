using System.Text.RegularExpressions;

namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class ClaudeModel : KeyedItem
{
    #region "Constants"
    const string costAmountPattern = @"\$(?<decimalAmount>.*?)\sper\s(?<tokenAmount>.*?)$";
    #endregion
    #region "Properties"
    public decimal? Rate { get; set; }
    public string? Model { get; set; }
    public string? Url { get; set; }
    public string? InputCost { get; set; }
    public string? OutputCost { get; set; }
    public TimeSpan? Duration { get; set; }
    public string? Status { get; set; }

    public dynamic? RequestMessage { get; set; }
    public dynamic? Result { get; set; }
    public dynamic? ResponseFromAI { get; set; }
    public decimal? Cost { get; set; }
    public decimal? ResponseCost { get; set; }
    #endregion
    #region "Implementation"
    protected Match? IsValidMatch(string? cost)
    {
        if (string.IsNullOrWhiteSpace(cost))
            return null;

        return Regex.Match(cost, costAmountPattern, RegexOptions.Singleline);
    }
    protected int? GetPerTokenAmount()
    {
        if (string.IsNullOrWhiteSpace(InputCost))
            return null;

        var isValidMatch = Regex.Match(InputCost, costAmountPattern, RegexOptions.Singleline);
        if (!(isValidMatch?.Success ?? false)) return null;

        var tokenAmount = isValidMatch.Groups["tokenAmount"].Value;
        switch (tokenAmount)
        {
            default:
                return 1000000;
        }
    }
    #endregion
    #region "Read-Only Properties"
    public decimal? InputRate
    {
        get
        {
            var isValidMatch = IsValidMatch(InputCost);
            if (!(isValidMatch?.Success ?? false)) return null;

            return decimal.TryParse(isValidMatch.Groups["decimalAmount"].Value, out var decimalValue) ? decimalValue : null;
        }
    }
    public decimal? OutputRate
    {
        get
        {
            var isValidMatch = IsValidMatch(OutputCost);
            if (isValidMatch == null || !isValidMatch.Success) return null;

            return decimal.TryParse(isValidMatch.Groups["decimalAmount"].Value, out var decimalValue) ? decimalValue : null;
        }
    }
    public int? PerTokenAmount => GetPerTokenAmount();
    #endregion
}
