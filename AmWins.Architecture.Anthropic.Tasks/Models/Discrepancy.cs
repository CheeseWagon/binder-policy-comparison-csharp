namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class Discrepancy
{
    #region "Properties"
    public bool? BinderMatchesActual { get; set; }
    public bool? PolicyMatchesActual { get; set; }
    public bool? SpecifiedInActual { get; set; }
    public bool? ValuesAreEqual { get; set; }
    public bool? VariedByPunctuationOnly { get; set; }
    public bool? VariedByCapitalizationOnly { get; set; }
    public bool? ValuesAreSynonyms { get; set; }
    public bool? ValuesAreNull { get; set; }
    public bool? NumericEquivalentsAreEqual { get; set; }
    public bool? DateEquivalentsAreEqual { get; set; }
    public bool? ValuesAreSubstring { get; set; }

    public ComparisonTopic? Binder { get; set; }
    public ComparisonTopic? Policy { get; set; }
    public ComparisonTopic? Other { get; set; }
    #endregion
    #region "Read-Only Properties"
    public bool IsCorrect => !VariedByPunctuationOnly.GetValueOrDefault() && !VariedByCapitalizationOnly.GetValueOrDefault() &&
                             !ValuesAreSynonyms.GetValueOrDefault() && !ValuesAreNull.GetValueOrDefault() && !ValuesAreSubstring.GetValueOrDefault() &&
                             !NumericEquivalentsAreEqual.GetValueOrDefault() && !DateEquivalentsAreEqual.GetValueOrDefault() &&
                             !ValuesAreEqual.GetValueOrDefault();

    public List<string>? Messages
    {
        get
        {
            if (IsCorrect) return null;

            var messages = new List<string>();
            if (ValuesAreEqual.GetValueOrDefault()) messages.Add("Values are equivalent.");
            if (BinderMatchesActual.GetValueOrDefault()) messages.Add("Binder matches actual");
            if (PolicyMatchesActual.GetValueOrDefault()) messages.Add("Policy matches actual");
            if (SpecifiedInActual.GetValueOrDefault()) messages.Add("Specified in actual");
            if (VariedByPunctuationOnly.GetValueOrDefault()) messages.Add("Varied by punctuation only");
            if (VariedByCapitalizationOnly.GetValueOrDefault()) messages.Add("Varied by capitalization only");
            if (ValuesAreSynonyms.GetValueOrDefault()) messages.Add("Values are synonyms");
            if (ValuesAreNull.GetValueOrDefault()) messages.Add("Values are null");
            if (NumericEquivalentsAreEqual.GetValueOrDefault()) messages.Add("Numeric equivalents are equal");
            if (DateEquivalentsAreEqual.GetValueOrDefault()) messages.Add("Date equivalents are equal");
            if (ValuesAreSubstring.GetValueOrDefault()) messages.Add("One value is contained within the other.");
            return messages.Any() ? messages : null;
        }
    }

    public string? Key => Binder?.Key ?? Policy?.Key ?? Other?.Key;
    public object? BinderValue => Binder?.Value;
    public object? PolicyValue => Policy?.Value;
    public object? OtherValue => Other?.Value;
    public bool MatchesActual => BinderMatchesActual.GetValueOrDefault() || PolicyMatchesActual.GetValueOrDefault();
    #endregion
}
