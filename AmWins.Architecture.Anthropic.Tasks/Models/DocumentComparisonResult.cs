namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class DocumentComparisonResult
{
    #region "Properties"
    public string? TopicsFile { get; set; }
    public string? Key { get; set; }
    public string? KeyName { get; set; }
    public List<ClaudeDocument>? Documents { get; set; }
    public List<ComparisonTopic>? Topics { get; set; }
    #endregion
    #region "Read-Only Properties"
    public List<Discrepancy>? Discrepancies => Documents?.FirstOrDefault(document => document.DocumentType == "discrepancies")?.Discrepancies;
    public List<ClaudeDocument> Binders => Documents?.Where(document => document.DocumentType == "binder").ToList() ?? new List<ClaudeDocument>();
    public List<ClaudeDocument> Policies => Documents?.Where(document => document.DocumentType == "policy").ToList() ?? new List<ClaudeDocument>();
    public List<ClaudeDocument> Checklists => Documents?.Where(document => document.DocumentType == "checklist").ToList() ?? new List<ClaudeDocument>();

    public int? TotalDiscrepanciesMatchedActual => TotalBinderDiscrepancyValuesMatchActual + TotalPolicyDiscrepancyValuesMatchActual;
    public int? TotalBinderDiscrepancyValuesMatchActual => Discrepancies?.Count(discrepancy => discrepancy.BinderMatchesActual.GetValueOrDefault());
    public int? TotalPolicyDiscrepancyValuesMatchActual => Discrepancies?.Count(discrepancy => discrepancy.PolicyMatchesActual.GetValueOrDefault());
    public int? TotalDiscrepanciesVariedByPuncutationOnly => Discrepancies?.Count(discrepancy => discrepancy.VariedByPunctuationOnly.GetValueOrDefault());
    public int? TotalDiscrepanciesVariedByCapitalizationOnly => Discrepancies?.Count(discrepancy => discrepancy.VariedByCapitalizationOnly.GetValueOrDefault());
    public int? TotalDiscrepanciesBothAreNulls => Discrepancies?.Count(discrepancy => discrepancy.ValuesAreNull.GetValueOrDefault());
    public int? TotalDiscrepanciesNumericEquivalentsAreEqual => Discrepancies?.Count(discrepancy => discrepancy.NumericEquivalentsAreEqual.GetValueOrDefault());
    public int? TotalDiscrepanciesDateRangesAreEqual => Discrepancies?.Count(discrepancy => discrepancy.DateEquivalentsAreEqual.GetValueOrDefault());
    #endregion
}
