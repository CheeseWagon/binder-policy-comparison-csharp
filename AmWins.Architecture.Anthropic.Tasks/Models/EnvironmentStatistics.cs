
namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class EnvironmentStatistics
{
    #region "Properties"
    public string? Id { get; set; }
    public string? Name { get; set; }
    public string? FileName { get; set; }

    public List<DocumentComparisonResult>? ComparisonResults { get; set; }
    #endregion
    #region "Read-Only Properties"
    public int? TotalExpectedValues => TotalExpectedValuesFoundInBinders + TotalExpectedValuesFoundInPolicies;
    public int? TotalExpectedValuesFoundInBinders
    {
        get
        {
            var totalExpectedValues = 0;
            var comparisonResults = ComparisonResults ?? new();
            foreach (var comparisonResult in comparisonResults)
            {
                var binders = comparisonResult.Binders ?? new();
                foreach (var binder in binders)
                {
                    var binderTopics = binder.Topics ?? new();
                    foreach (var binderTopic in binderTopics)
                    {
                        var binderResultsString = binderTopic.ToResultsString();
                        if (!string.IsNullOrWhiteSpace(binderResultsString) && binderResultsString != "Nothing")
                            totalExpectedValues += 1;
                    }
                }
            }
            return totalExpectedValues;
        }
    }
    public int? TotalExpectedValuesFoundInPolicies
    {
        get
        {
            var totalExpectedValues = 0;
            var comparisonResults = ComparisonResults ?? new();
            foreach (var comparisonResult in comparisonResults)
            {
                var policies = comparisonResult.Policies ?? new();
                foreach (var policy in policies)
                {
                    var policyTopics = policy.Topics ?? new();
                    foreach (var policyTopic in policyTopics)
                    {
                        var policyResultsString = policyTopic.ToResultsString();
                        if (!string.IsNullOrWhiteSpace(policyResultsString) && policyResultsString != "Nothing")
                            totalExpectedValues += 1;
                    }
                }
            }
            return totalExpectedValues;
        }
    }
    public int? TotalDiscrepancies
    {
        get
        {
            var totalDiscrepancies = 0;
            var comparisonResults = ComparisonResults ?? new();
            foreach (var comparisonResult in comparisonResults)
            {
                var comparisonResultDiscrepancies = comparisonResult.Discrepancies ?? new();
                totalDiscrepancies += comparisonResultDiscrepancies.Count;
            }
            return totalDiscrepancies;
        }
    }

    public List<ComparisonTopic> Topics
    {
        get
        {
            var topics = new List<ComparisonTopic>();
            var comparisonResults = ComparisonResults ?? new();
            foreach (var comparisonResult in comparisonResults)
            {
                var comparisonResultTopics = comparisonResult.Topics ?? new();
                foreach (var topic in comparisonResultTopics)
                {
                    var existingTopic = topics.FirstOrDefault(t => t.Key == topic.Key);
                    if (existingTopic != null)
                    {
                        existingTopic.BinderTotal += topic.BinderTotal.GetValueOrDefault();
                        existingTopic.PolicyTotal += topic.PolicyTotal.GetValueOrDefault();
                        existingTopic.DiscrepancyCount += topic.DiscrepancyCount.GetValueOrDefault();
                        existingTopic.DiscrepancyMatchedActualCount += topic.DiscrepancyMatchedActualCount.GetValueOrDefault();
                        if (topic.Documents != null && topic.Documents.Any())
                        {
                            existingTopic.Documents ??= new();
                            existingTopic.Documents.AddRange(topic.Documents);
                        }
                    }
                    else
                    {
                        var newTopic = new ComparisonTopic
                        {
                            BinderTotal = topic.BinderTotal.GetValueOrDefault(),
                            PolicyTotal = topic.PolicyTotal.GetValueOrDefault(),
                            DiscrepancyCount = topic.DiscrepancyCount.GetValueOrDefault(),
                            DiscrepancyMatchedActualCount = topic.DiscrepancyMatchedActualCount.GetValueOrDefault(),
                            Key = topic.Key,
                            Text = topic.Text,
                            Type = topic.Type
                        };
                        if (topic.Documents != null && topic.Documents.Any())
                        {
                            newTopic.Documents ??= new();
                            newTopic.Documents.AddRange(topic.Documents);
                        }
                        topics.Add(newTopic);
                    }
                }
            }

            return topics;
        }
    }

    public List<ClaudeDocument> Binders => ComparisonResults?.SelectMany(result => result.Binders).ToList() ?? new List<ClaudeDocument>();
    public List<ClaudeDocument> Policies => ComparisonResults?.SelectMany(result => result.Policies).ToList() ?? new List<ClaudeDocument>();
    public List<ClaudeDocument> Checklists => ComparisonResults?.SelectMany(result => result.Checklists).ToList() ?? new List<ClaudeDocument>();

    public bool? HasChecklists => Checklists?.Any();
    public int? TotalBindersProcessed => ComparisonResults?.Sum(result => result?.Documents?.FirstOrDefault(document => document.DocumentType == "binder") != null ? 1 : 0);
    public int? TotalPoliciesProcessed => ComparisonResults?.Sum(result => result?.Documents?.FirstOrDefault(document => document.DocumentType == "policy") != null ? 1 : 0);
    public int? TotalDiscrepanciesFound => ComparisonResults?.Sum(result => result.Documents?.FirstOrDefault(document => document.DocumentType == "discrepancies")?.Discrepancies?.Count ?? 0);
    public int? TotalDiscrepanciesMatchedActual => ComparisonResults?.Sum(result => result.TotalDiscrepanciesMatchedActual);
    public int? TotalBinderDiscrepancyValuesMatchActual => ComparisonResults?.Sum(result => result.TotalBinderDiscrepancyValuesMatchActual);
    public int? TotalPolicyDiscrepancyValuesMatchActual => ComparisonResults?.Sum(result => result.TotalPolicyDiscrepancyValuesMatchActual);


    public int? TotalDiscrepanciesVariedByPuncutationOnly => ComparisonResults?.Sum(comparisonResult => comparisonResult.TotalDiscrepanciesVariedByPuncutationOnly);
    public int? TotalDiscrepanciesVariedByCapitalizationOnly => ComparisonResults?.Sum(comparisonResult => comparisonResult.TotalDiscrepanciesVariedByCapitalizationOnly);
    public int? TotalDiscrepanciesBothAreNulls => ComparisonResults?.Sum(comparisonResult => comparisonResult.TotalDiscrepanciesBothAreNulls);
    public int? TotalDiscrepanciesNumericEquivalentsAreEqual => ComparisonResults?.Sum(comparisonResult => comparisonResult.TotalDiscrepanciesNumericEquivalentsAreEqual);
    public int? TotalDiscrepanciesDateRangesAreEqual => ComparisonResults?.Sum(comparisonResult => comparisonResult.TotalDiscrepanciesDateRangesAreEqual);
    #endregion 
}
