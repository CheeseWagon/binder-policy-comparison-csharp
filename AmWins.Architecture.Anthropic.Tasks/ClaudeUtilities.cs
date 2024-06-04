using System.Diagnostics;
using System.Dynamic;
using System.Text;

using Newtonsoft.Json.Serialization;
using Newtonsoft.Json;

using Microsoft.Extensions.Configuration;

using AmWins.Architecture.Anthropic.Tasks.Models;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Text.Json;
using OfficeOpenXml;
using AmWins.Architecture.Anthropic.Tasks;
using System.Reflection;

namespace AmWins.Architecture.Anthropic.Tasks;

public static class ClaudeUtilities
{

    #region "Implementation"
    static void TokenizeObject(this JObject obj, List<string> tokens)
    {
        foreach (var property in obj.Properties())
        {
            tokens.Add(property.Name);

            if (property.Value.Type == JTokenType.Object)
            {
                TokenizeObject((JObject)property.Value, tokens);
            }
            else if (property.Value.Type == JTokenType.Array)
            {
                ((JArray)property.Value).TokenizeArray(tokens);
            }
            else if (property.Value.Type == JTokenType.String)
            {
                string value = (string)property.Value;

                // Check if the value is an image
                if (IsImageContent(value))
                {
                    // Tokenize the image content
                    tokens.AddRange(TokenizeImageContent(value));
                }
                else
                {
                    // Tokenize the string value
                    tokens.AddRange(TokenizeString(value));
                }
            }
            else
            {
                tokens.Add(property.Value.ToString());
            }
        }
    }
    static void TokenizeArray(this JArray array, List<string> tokens)
    {
        foreach (var item in array)
        {
            if (item.Type == JTokenType.Object)
            {
                TokenizeObject((JObject)item, tokens);
            }
            else if (item.Type == JTokenType.Array)
            {
                TokenizeArray((JArray)item, tokens);
            }
            else if (item.Type == JTokenType.String)
            {
                string value = (string)item;

                // Check if the value is an image
                if (IsImageContent(value))
                {
                    // Tokenize the image content
                    tokens.AddRange(TokenizeImageContent(value));
                }
                else
                {
                    // Tokenize the string value
                    tokens.AddRange(TokenizeString(value));
                }
            }
            else
            {
                tokens.Add(item.ToString());
            }
        }
    }
    static bool IsImageContent(string value)
    {
        // Check if the value represents an image (e.g., base64-encoded image data)
        // Modify this method based on your specific image content format
        return value.StartsWith("data:image");
    }
    static List<string> TokenizeImageContent(string imageContent)
    {
        // Tokenize the image content based on your specific requirements
        // This example assumes the image content is a base64-encoded string
        return new List<string> { "<image>" };
    }
    static List<string> TokenizeString(string value)
    {
        // Tokenize the string value based on your specific requirements
        // This example splits the string into individual words using regular expressions
        return Regex.Split(value, @"\W+")
            .Where(token => !string.IsNullOrWhiteSpace(token))
            .ToList();
    }

    #endregion
    #region "Extensions"
    public static EnvironmentStatistics? Analyze(this IConfiguration? configuration, List<DocumentComparisonResult>? comparisonResults, string? clientName = null)
    {
        if (comparisonResults == null || !comparisonResults.Any()) return null;

        var environmentStatistics = new EnvironmentStatistics
        {
            Name = clientName ?? "Unknown",
            ComparisonResults = new()
        };

        var topicsSummary = new List<ComparisonTopic>();
        foreach (var comparisonResult in comparisonResults)
        {
            #region "Actual Checklist"
            var checklistDocument = comparisonResult.Documents?.FirstOrDefault(result => result.DocumentType == "checklist");
            var checklistDiscrepancies = checklistDocument?.Discrepancies ?? new();
            var notAvailableTopics = checklistDocument?.NotApplicableTopics ?? new();
            var comparisonTopics = comparisonResult.Topics ?? new();
            #endregion
            #region "Discrepancies"
            // Determine what discrepancies where found by the Comparison result
            var discrepancies = comparisonResult.Discrepancies ?? new();
            foreach (var discrepancy in discrepancies)
            {
                var discrepancyKey = discrepancy.Key;
                var binderValue = discrepancy.Binder?.ToResultsString();
                var policyValue = discrepancy.Policy?.ToResultsString();

                var comparisonTopic = comparisonTopics.FirstOrDefault(topic => string.Compare(topic.Key, discrepancyKey, StringComparison.InvariantCultureIgnoreCase) == 0);
                if (!string.IsNullOrWhiteSpace(binderValue) || !string.IsNullOrWhiteSpace(policyValue))
                {
                    if (comparisonTopic != null)
                    {
                        comparisonTopic.DiscrepancyCount ??= 0;
                        comparisonTopic.DiscrepancyCount++;
                    }
                }

                // Determine if the discrepancy was due to punctuation only
                if (!string.IsNullOrWhiteSpace(binderValue) || !string.IsNullOrWhiteSpace(policyValue))
                    discrepancy.ValuesAreEqual = string.CompareOrdinal(binderValue, policyValue) == 0;

                if (!discrepancy.ValuesAreEqual.GetValueOrDefault())
                    discrepancy.VariedByPunctuationOnly = binderValue?.StripPunctuation() == policyValue?.StripPunctuation();

                // Determine if the discrepancy was due to capitalization only
                if (!discrepancy.ValuesAreEqual.GetValueOrDefault() &&
                    !discrepancy.VariedByPunctuationOnly.GetValueOrDefault())
                    discrepancy.VariedByCapitalizationOnly = binderValue?.StripPunctuation().Unformat() == policyValue?.StripPunctuation().Unformat();

                // Determine if both values are null
                discrepancy.ValuesAreNull = (binderValue == null && policyValue == null) || (policyValue == "Nothing" && binderValue == "Nothing");

                // Values Contain a Substring (One value contains the other)
                if (binderValue != null && policyValue != null && !discrepancy.ValuesAreEqual.GetValueOrDefault() &&
                    !discrepancy.ValuesAreNull.GetValueOrDefault() && !discrepancy.VariedByPunctuationOnly.GetValueOrDefault() &&
                    !discrepancy.VariedByCapitalizationOnly.GetValueOrDefault())
                    discrepancy.ValuesAreSubstring = binderValue.Contains(policyValue) || policyValue.Contains(binderValue);

                //var binderNumericValue = binderValue?.ExtractNumbers();
                //var policyNumericValue = policyValue?.ExtractNumbers();
                //if (binderNumericValue != null && policyNumericValue != null)
                //    discrepancy.NumericEquivalentsAreEqual = binderNumericValue == policyNumericValue;

                var binderValueDates = binderValue?.ExtractDates();
                var policyValueDates = policyValue?.ExtractDates();
                if (binderValueDates != null && policyValueDates != null)
                    discrepancy.DateEquivalentsAreEqual = binderValueDates.DateRangesAreEqual(policyValueDates);

                var checkListDiscrepancy = checklistDiscrepancies.FirstOrDefault(cd => cd.Key == discrepancy.Key);
                if (checkListDiscrepancy != null && (!string.IsNullOrWhiteSpace(binderValue) || !string.IsNullOrWhiteSpace(policyValue)))
                {
                    var binderMatchValue = binderValue.StripPunctuation().Unformat();
                    var policyMatchValue = policyValue.StripPunctuation().Unformat();
                    var checklistBinderValue = checkListDiscrepancy.Binder?.ToResultsString()?.StripPunctuation().Unformat();
                    var checklistPolicyValue = checkListDiscrepancy.Policy?.ToResultsString()?.StripPunctuation().Unformat();
                    if (!string.IsNullOrWhiteSpace(checklistBinderValue) || !string.IsNullOrWhiteSpace(checklistPolicyValue))
                    {
                        discrepancy.BinderMatchesActual = binderMatchValue == checklistBinderValue;
                        discrepancy.PolicyMatchesActual = policyMatchValue == checklistPolicyValue;

                        if (discrepancy.BinderMatchesActual.GetValueOrDefault() || discrepancy.PolicyMatchesActual.GetValueOrDefault())
                        {
                            if (comparisonTopic != null)
                            {
                                comparisonTopic.DiscrepancyMatchedActualCount ??= 0;
                                comparisonTopic.DiscrepancyMatchedActualCount++;
                            }
                        }
                        discrepancy.SpecifiedInActual = true;
                    }
                }
            }
            #endregion
            #region "Binder Values Found"
            // Total the number of times that the current binder found a value expected from the prompt
            var binderDocument = comparisonResult.Documents?.FirstOrDefault(result => result.DocumentType == "binder");
            if (binderDocument != null)
            {
                var binderTopics = binderDocument.Topics ?? new();
                foreach (var comparisonTopic in comparisonTopics)
                {
                    var binderTopic = binderTopics.FirstOrDefault(topic => string.Compare(topic.Key, comparisonTopic.Key, StringComparison.InvariantCultureIgnoreCase) == 0);
                    var binderValue = binderTopic?.ToResultsString();
                    if (!string.IsNullOrWhiteSpace(binderValue) && binderValue != "Nothing")
                    {
                        comparisonTopic.BinderTotal ??= 0;
                        comparisonTopic.BinderTotal++;

                        comparisonTopic.Documents ??= new();
                        comparisonTopic.Documents.Add(new ClaudeDocument { FileName = binderDocument.FileName, DocumentType = binderDocument.DocumentType });
                    }
                }
            }
            #endregion
            #region "Policy Values Found"
            // Total the number of times that the current policy found a value expected from the prompt
            var policyDocument = comparisonResult.Documents?.FirstOrDefault(result => result.DocumentType == "policy");
            if (policyDocument != null)
            {
                var policyTopics = policyDocument.Topics ?? new();
                foreach (var comparisonTopic in comparisonTopics)
                {
                    var policyTopic = policyTopics.FirstOrDefault(topic => string.Compare(topic.Key, comparisonTopic.Key, StringComparison.InvariantCultureIgnoreCase) == 0);
                    var policyValue = policyTopic?.ToResultsString();
                    if (!string.IsNullOrWhiteSpace(policyValue) && policyValue != "Nothing")
                    {
                        comparisonTopic.PolicyTotal ??= 0;
                        comparisonTopic.PolicyTotal++;

                        comparisonTopic.Documents ??= new();
                        comparisonTopic.Documents.Add(new ClaudeDocument { FileName = policyDocument.FileName, DocumentType = policyDocument.DocumentType });
                    }
                }
            }
            #endregion
        }

        environmentStatistics.ComparisonResults = comparisonResults;

        return environmentStatistics;
    }
    public static async Task<ClaudeResponse?> CompareDocuments(this IConfiguration configuration, DocumentComparisonRequest? request, HttpClient? client = null)
    {
        var promptsToUse = configuration?.Prompts();
        if (promptsToUse == null || !promptsToUse.Any() ||
            request?.Documents == null || request?.Documents.Count < 2) return null;

        client ??= configuration?.CreateClient();
        var documentsToCompare = request?.Documents ?? new();
        var modelsToUse = configuration?.Models() ?? new();
        foreach (var promptToUse in promptsToUse)
        {
            var stopWatch = Stopwatch.StartNew();

            var requestBody = configuration?.ToMessage(documentsToCompare, promptToUse);
            if (!request?.SubmitRequestToClaude.GetValueOrDefault() ?? true ||
                requestBody == null) continue;

            promptToUse.ModelsUsed ??= new();
            foreach (var modelToUse in modelsToUse)
            {
                var modelStopWatch = Stopwatch.StartNew();
                var modelResponse = new ClaudeModel { Key = modelToUse.Key, Model = modelToUse.Model, InputCost = modelToUse.InputCost, OutputCost = modelToUse.OutputCost };

                if (requestBody != null)
                    requestBody.model = modelResponse.Model;

                var requestMessageJson = JsonConvert.SerializeObject(requestBody, new JsonSerializerSettings
                {
                    ContractResolver = new CamelCasePropertyNamesContractResolver(),
                    NullValueHandling = NullValueHandling.Ignore
                }).Replace("#n", @"\n").Replace("#t", @"\t");

                var requestResult = new { result = JsonConvert.DeserializeObject<dynamic>(requestMessageJson) };
                var requestResultJson = JsonConvert.SerializeObject(requestResult, new JsonSerializerSettings
                {
                    ContractResolver = new CamelCasePropertyNamesContractResolver(),
                    NullValueHandling = NullValueHandling.Ignore
                });

                modelResponse.RequestMessage = requestResultJson;
                var requestMessage = new HttpRequestMessage(HttpMethod.Post, modelToUse.Url);
                requestMessage.Content = new StringContent(requestMessageJson, Encoding.UTF8, "application/json");
                var httpResponseMessage = await client.SendAsync(requestMessage);
                if (httpResponseMessage == null)
                    modelResponse.Status = "No response returned from Claude.";
                else
                {
                    modelResponse.Status = "Response returned from Claude AI.";

                    var responseBody = await httpResponseMessage.Content.ReadAsStringAsync();

                    var responseResult = new { result = JsonConvert.DeserializeObject<dynamic>(responseBody) };
                    var responseResultJson = JsonConvert.SerializeObject(responseResult, new JsonSerializerSettings
                    {
                        ContractResolver = new CamelCasePropertyNamesContractResolver(),
                        NullValueHandling = NullValueHandling.Ignore
                    });

                    modelResponse.Result = JsonConvert.DeserializeObject<dynamic>(responseBody);

                    modelResponse.Cost = requestResultJson.CalculateMessageCost(modelToUse) * modelToUse.InputRate;
                    modelResponse.ResponseCost = responseResultJson.CalculateMessageCost(modelResponse) * modelToUse.OutputRate;
                }

                modelStopWatch.Stop();
                var totalModelTime = TimeSpan.FromMilliseconds(stopWatch.ElapsedMilliseconds);
                modelResponse.Duration = totalModelTime;

                promptToUse.ModelsUsed.Add(modelResponse);
            }

            stopWatch.Stop();
            var totalTime = TimeSpan.FromMilliseconds(stopWatch.ElapsedMilliseconds);
            promptToUse.Duration = totalTime;
        }

        return new ClaudeResponse { DateOfExecution = DateTime.Now, PromptsExecuted = promptsToUse };
    }
    public static List<ClaudePrompt>? Prompts(this IConfiguration configuration)
    {
        var promptsToUseSection = configuration.GetSection("promptsToUse");
        if (promptsToUseSection == null) return null;

        var promptsToUse = promptsToUseSection.Get<List<string>>() ?? new();
        if (!promptsToUse.Any()) return null;

        var promptsToReturn = new List<ClaudePrompt>();
        foreach (var promptKey in promptsToUse)
        {
            var prompt = configuration.AvailablePrompts()?.FirstOrDefault(p => string.Compare(p.Key, promptKey, StringComparison.InvariantCultureIgnoreCase) == 0);
            if (prompt != null) promptsToReturn.Add(prompt);
        }

        return promptsToReturn;
    }
    public static List<ClaudePrompt>? AvailablePrompts(this IConfiguration configuration) 
    {
        var availablePromptsSection = configuration.GetSection("prompts");
        if (availablePromptsSection == null) return null;

        var availablePrompts = availablePromptsSection.Get<List<ClaudePrompt>>() ?? new();
        if (!availablePrompts.Any()) return null;

        return availablePrompts;
    }
    public static HttpClient CreateClient(this IConfiguration configuration, IHttpClientFactory? factory = null)
    {
        var client = factory?.CreateClient() ?? new();
        if (configuration != null)
        {
            client.DefaultRequestHeaders.Add("x-api-key", configuration["claudeApiKey"] ?? string.Empty);
            client.DefaultRequestHeaders.Add("anthropic-version", configuration["claudeVersion"] ?? "2023-06-01");
        }

        client.Timeout = TimeSpan.FromMinutes(30);

        return client;
    }
    public static List<ClaudeModel>? Models(this IConfiguration configuration)
    {
        var modelsToProcessSection = configuration.GetSection("modelsToProcess");
        if (modelsToProcessSection == null) return null;

        var modelsToProcess = modelsToProcessSection.Get<List<string>>() ?? new();
        if (!modelsToProcess.Any()) return null;

        var modelsToReturn = new List<ClaudeModel>();
        foreach (var modelKey in modelsToProcess)
        {
            var model = configuration?.AvailableModels()?.FirstOrDefault(m => string.Compare(m.Key, modelKey, StringComparison.InvariantCultureIgnoreCase) == 0);
            if (model != null) modelsToReturn.Add(model);
        }

        return modelsToReturn;
    }
    public static List<ClaudeModel>? AvailableModels(this IConfiguration? configuration, string? clientKeyName = null)
    {
        if (configuration == null) return new();
        return configuration?.GetSection($"availableModels")?.Get<List<ClaudeModel>>() ?? new();
    }
    public static dynamic? ToMessage(this IConfiguration? configuration, List<ClaudeDocument> documents, ClaudePrompt promptToUse)
    {
        if (documents == null || documents.Count == 0 ||
            promptToUse?.Sections == null || !promptToUse.Sections.Any()) return null;

        // create an array of type dynamic
        var messages = new List<dynamic>();

        var messageToReturn = new ExpandoObject() as IDictionary<string, object>;

        var maxTokensToSample = 4096;
        var maxTokensValue = configuration?["maxTokensToSample"];
        if (!string.IsNullOrWhiteSpace(maxTokensValue) && int.TryParse(maxTokensValue, out var maxTokens))
            maxTokensToSample = maxTokens; 

        messageToReturn.Add("max_tokens", maxTokensToSample);

        if (!string.IsNullOrWhiteSpace(promptToUse.System))
        {
            var systemMessage = promptToUse.System.Replace("{{lineFeed}}", Constants.LINE_FEED_TEMPLATE).Replace("{{tab}}", Constants.TAB_TEMPLATE);
            if (!string.IsNullOrWhiteSpace(promptToUse?.TopicsFile))
                systemMessage = systemMessage.Replace("{{TOPICS}}", promptToUse.GetTopics() ?? string.Empty);
            messageToReturn.Add("system", systemMessage);
        }

        var sections = promptToUse?.Sections ?? new();
        foreach (var section in sections)
        {
            var message = new ExpandoObject() as IDictionary<string, object>;
            var sectionType = section.Type ?? ePromptSectionType.Text;
            switch (sectionType)
            {
                case ePromptSectionType.Text:
                    messages.Add(documents.GetTextMessage(section));
                    break;
                case ePromptSectionType.Images:
                    {
                        var imagesMessage = documents.GetImagesMessage(section);
                        if (imagesMessage != null)
                            messages.Add(imagesMessage);
                    }
                    break;
            }
        }

        messageToReturn.Add("messages", messages);

        return messageToReturn;
    }
    public static dynamic? GetTextMessage(this List<ClaudeDocument> documents, PromptSection promptSection)
    {
        var messageToReturn = new ExpandoObject() as IDictionary<string, object>;

        if (!string.IsNullOrWhiteSpace(promptSection.RoleText))
            messageToReturn.Add("role", promptSection.RoleText);

        string? messageContent = null;

        var requestDocuments = promptSection.RequestedDocuments;
        if (requestDocuments.HasValue)
        {
            var requestedDocument = documents?[requestDocuments.Value - 1];
            if (!string.IsNullOrWhiteSpace(promptSection?.Value))
                messageContent = promptSection.Value.Replace(promptSection.Key.ToPlaceholder() ?? string.Empty, requestedDocument?.ExtractedText) ?? string.Empty;
        }

        if (string.IsNullOrWhiteSpace(messageContent))
            messageContent = promptSection?.Value;

        if (messageContent is not null)
        {
            List<dynamic> messages = new List<dynamic>();

            var textMessage = new ExpandoObject() as IDictionary<string, object>;

            textMessage.Add("type", "text");
            textMessage.Add("text", messageContent.Replace("key", promptSection?.Key).Replace("{{lineFeed}}", Constants.LINE_FEED_TEMPLATE).Replace("{{tab}}", Constants.TAB_TEMPLATE).Replace("{{quote}}", Constants.DOUBLE_QUOTE_TEMPLATE));

            messageToReturn.Add("content", new[] { textMessage });
        }

        return !messageToReturn.Any() ? null : messageToReturn;
    }
    public static dynamic? GetImagesMessage(this List<ClaudeDocument> documents, PromptSection promptSection)
    {
        if (documents == null || !documents.Any() || !promptSection.RequestedDocuments.HasValue) return null;

        var requestDocumentsIndex = promptSection.RequestedDocuments.Value - 1;

        var requestedDocument = documents?[requestDocumentsIndex];
        if (requestedDocument?.PageImages == null || !requestedDocument.PageImages.Any()) return null;

        var messageToReturn = new ExpandoObject() as IDictionary<string, object>;

        if (!string.IsNullOrWhiteSpace(promptSection.RoleText))
            messageToReturn.Add("role", promptSection.RoleText);

        var messages = new List<dynamic>();

        var pageImages = requestedDocument.PageImages;
        foreach (var pageImage in pageImages)
        {
            var message = new ExpandoObject() as IDictionary<string, object>;

            var imageBytes = Convert.FromBase64String(pageImage.Data ?? string.Empty);
            var imageData = Encoding.UTF8.GetString(imageBytes);
            message.Add("type", "image");
            message.Add("source", new { type = "base64", media_type = pageImage.ContentType, data = imageData });
            messages.Add(message);
        }

        var textMessages = promptSection.Text ?? new();
        foreach (var textMessage in textMessages)
        {
            var message = new ExpandoObject() as IDictionary<string, object>;
            message.Add("type", "text");
            message.Add("content", textMessage.Replace("key", promptSection.Key ?? "nokey").Replace("{{lineFeed}}", Constants.LINE_FEED_TEMPLATE).Replace("{{tab}}", Constants.TAB_TEMPLATE).Replace("{{quote}}", Constants.DOUBLE_QUOTE_TEMPLATE));
            messages.Add(message);
        }

        messageToReturn.Add("content", messages);

        return !messages.Any() ? null : messageToReturn;
    }
    public static string? ToResultsString(this ComparisonTopic? results)
    {
        if (results?.Result == null) return "Nothing";

        dynamic? dynamicResults = results.Result;
        if (dynamicResults is JArray jarray)
        {
            var topicsList = jarray?.Children();
            if (topicsList == null) return "Nothing";

            var topicValues = new List<string>();
            foreach (var result in topicsList)
                topicValues.Add(result.ToString());

            return string.Join("\n", topicValues);
        }
        else if (dynamicResults is JObject resultsObject)
        {
            var topicValues = new List<string>();
            var properties = resultsObject.Properties();
            foreach (var property in properties)
            {
                var topicKey = string.Empty;
                string? topicValue = null;
                var propertyName = property.Name;
                if (propertyName == "name" || propertyName == "key")
                {
                    topicKey = property.Value.ToString().FormatTopicKey();
                    var propertyValue = properties.FirstOrDefault(prop => prop.Name == "value");
                    if (propertyValue != null)
                    {
                        try
                        {
                            topicValue = propertyValue.Value.ToString();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }

                    }
                }
                else if (propertyName != "value")
                {
                    try
                    {
                        topicKey = propertyName.FormatTopicKey();
                        topicValue = property.Value.ToString();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }


                }

                if (!string.IsNullOrWhiteSpace(topicKey) && !string.IsNullOrWhiteSpace(topicValue?.ToString()))
                    topicValues.Add($"{topicKey}: {topicValue}");
            }

            return string.Join("\n", topicValues);
        }
        else if (dynamicResults is JProperty property)
        {
            var topicKey = string.Empty;
            object? topicValue = null;
            var propertyName = property.Name;
            if (propertyName != "value")
            {
                topicKey = propertyName.FormatTopicKey();
                topicValue = property.Value;
            }

            if (!string.IsNullOrWhiteSpace(topicKey))
                return topicValue.ToSafeString();
        }

        try
        {
            return dynamicResults?.ToString();
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }

        return "Nothing";
    }
    public static string StripPunctuation(this string item)
    {
        var builder = new StringBuilder();
        var punctuation = new[] { '.', '#', '$', '%', '^', '*', '@', '!', ';', '?', '/', ',', ']', '[', '{', '}', '+', '-', '_', '=', '(', ')', '&', '\'', ' ' };
        foreach (var character in item)
        {
            if (!punctuation.Contains(character))
            {
                builder.Append(character);
            }
        }
        return builder.ToString();
    }
    public static string Unformat(this string inString)
    {

        var builder = new StringBuilder();
        var newString = inString;

        // Replace an ampersand with the word "And"
        if (newString.IndexOf('&') != -1)
            newString = newString.Replace("&", "AND");

        var aryString = newString.ToCharArray();

        foreach (var c in aryString)
        {
            if (char.IsNumber(c))
                builder.Append(c);
            else
            {
                if (char.IsLower(c))
                    builder.Append(char.ToUpper(c));
                else
                {
                    if (char.IsUpper(c))
                        builder.Append(c);
                }
            }
        }

        return builder.ToString();
    }
    public static List<DateTime>? ExtractDates(this string inString)
    {
        var dates = new List<DateTime>();
        var newString = inString;

        const string mmddyyyySlashPattern = @"^(0[1-9]|1[0-2])\/(0[1-9]|1\d|2\d|3[01])\/(19|20)\d{2}$";
        const string yyyymmddSlashPattern = @"^(19|20)\d{2}\/(0[1-9]|1[0-2])\/(0[1-9]|1\d|2\d|3[01])$";
        const string ddshortmonthyyyyPattern = @"(0[1-9]|[12]\d|3[01]) (Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) (19|20)\d{2}$";
        const string ddlongmonthyyyyPattern = @"(0[1-9]|[12]\d|3[01]) (January|Feburary|March|April|May|June|July|August|September|October|November|December) (19|20)\d{2}$";
        const string shortmonthddyyyyPattern = @"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) (0[1-9]|[12]\d|3[01]), (19|20)\d{2}$";
        const string longmonthddyyyyPattern = @"(January|Feburary|March|April|May|June|July|August|September|October|November|December) (0[1-9]|[12]\d|3[01]), (19|20)\d{2}$";
        const string mmddyyyyDashPattern = @"^(0[1-9]|[12]\d|3[01])-(0[1-9]|1[0-2])-(19|20)\d{2}$";
        const string yyyymmddDashPattern = @"^(19|20)\d{2}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$";

        var patterns = new[] { mmddyyyySlashPattern, yyyymmddSlashPattern, ddshortmonthyyyyPattern, ddlongmonthyyyyPattern, shortmonthddyyyyPattern, longmonthddyyyyPattern, mmddyyyyDashPattern, yyyymmddDashPattern };
        foreach (var pattern in patterns)
        {
            var dateMatches = Regex.Matches(newString, pattern);
            foreach (Match match in dateMatches)
            {
                if (DateTime.TryParse(match.Value, out var result) && !dates.Contains(result))
                    dates.Add(result);
            }
        }

        return dates.Any() ? dates : null;
    }
    public static bool? DateRangesAreEqual(this List<DateTime> dateRanges, List<DateTime> otherDateRanges)
    {
        if (dateRanges == null || otherDateRanges == null || dateRanges.Count != otherDateRanges.Count) return false;

        foreach (var date in dateRanges)
        {
            if (!otherDateRanges.Contains(date))
                return false;
        }

        return true;
    }
    public static void ToJsonFile<T>(this T itemToSave, string fileName) where T : class
    {
        if (itemToSave == default(T) || string.IsNullOrWhiteSpace(fileName))
            return;

        var exportDirectory = Path.GetDirectoryName(fileName);
        if (exportDirectory == null)
            return;

        if (!Directory.Exists(exportDirectory))
            Directory.CreateDirectory(exportDirectory);

        File.WriteAllText(fileName,
            JsonConvert.SerializeObject(itemToSave,
                new JsonSerializerSettings
                {
                    ContractResolver = new CamelCasePropertyNamesContractResolver(),
                    NullValueHandling = NullValueHandling.Ignore
                }));
    }
    public static async Task<T?> LoadJsonFromFile<T>(this string fileName) where T : class
    {
        var json = await File.ReadAllTextAsync(fileName);
        var item = JsonConvert.DeserializeObject<T>(json);
        return item;
    }
    public static dynamic? ExtractActualResponse(this string? responseJson)
    {
        if (string.IsNullOrWhiteSpace(responseJson)) return null;

        dynamic? response = null;
        try
        {
            // Parse the JSON string
            var jsonDocument = JsonDocument.Parse(responseJson);

            var hasResultElement = jsonDocument.RootElement.TryGetProperty("result", out var resultElement);
            if (!hasResultElement) return null;

            // Get the "content" array from the JSON
            var hasContentElement = resultElement.TryGetProperty("content", out var contentElement);
            if (!hasContentElement) return response = JsonConvert.DeserializeObject<dynamic>(JsonConvert.SerializeObject(resultElement));

            var contentArray = contentElement.EnumerateArray();
            if (!contentArray.Any()) return response = JsonConvert.DeserializeObject<dynamic>(JsonConvert.SerializeObject(contentElement)); ;

            // Find the first object in the "content" array that has a "text" property
            var textObject = contentArray.FirstOrDefault(obj => obj.TryGetProperty("text", out _));

            // If no object with a "text" property is found, return an empty string
            if (textObject.ValueKind == JsonValueKind.Undefined)
                return null;

            // Get the value of the "text" property
            var textValue = textObject.GetProperty("text").GetString();
            if (string.IsNullOrWhiteSpace(textValue)) return null;

            try
            {
                var textJson = "{" + textValue;
                if (!textJson.EndsWith("}"))
                    textJson += "}";

                response = JsonConvert.DeserializeObject<dynamic>(textJson);
            }
            catch (Exception)
            {
                try
                {
                    response = JsonConvert.DeserializeObject<dynamic>(textValue);
                }
                catch (Exception)
                {
                    response = null;
                }
            }
        }
        catch (Exception)
        {
            response = null;
        }

        return response;
    }
    public static string? FormatTopicKey(this string? topicKey) => topicKey?.Replace(" ", string.Empty).ToCamelCase();
    public static string ToCamelCase(this string item)
    {
        var words = item.Trim().Split(' ');
        for (var index = 0; index < words.Length; index++)
            if (words[index].Length > 0)
                words[index] = words[index][..1].ToLower() + words[index][1..];
        var returnString = string.Join(" ", words);
        return returnString;
    }
    public static bool Among(this object target, params object[] items)
    {
        return items.Contains(target);
    }
    public static async Task<DocumentComparisonResult?> GetResultsFromFile(this string? resultsFileLocation)
    {
        if (string.IsNullOrWhiteSpace(resultsFileLocation) || !File.Exists(resultsFileLocation)) return null;

        var comparisonResult = new DocumentComparisonResult();
        var results = new List<ClaudeDocument>();

        var binderDocument = new ClaudeDocument { DocumentType = "binder", Topics = new() };
        var policyDocument = new ClaudeDocument { DocumentType = "policy", Topics = new() };

        var actualResult = await resultsFileLocation.LoadJsonFromFile<dynamic>();
        if (actualResult == null) return null;

        dynamic topicsDocuments = actualResult["both"] ?? actualResult["common"] ?? actualResult["cyber"];
        if (topicsDocuments == null) return null;

        if (actualResult["both"] != null)
            comparisonResult.TopicsFile = "commonAndCyberDeclarations.json";
        else if (actualResult["common"] != null)
            comparisonResult.TopicsFile = "commonDeclarations.json";
        else if (actualResult["cyber"] != null)
            comparisonResult.TopicsFile = "cyberDeclarations.json";

        #region "Binder Topics"
        var binderResults = topicsDocuments["firstdocument"];
        if (binderResults != null && binderResults is JArray)
        {
            var topicsList = binderResults?.Children();
            foreach (JObject result in topicsList)
            {
                var properties = result.Properties();
                foreach (var property in properties)
                {
                    var topicKey = string.Empty;
                    object? topicValue = null;
                    var propertyName = property.Name;
                    if (propertyName == "name" || propertyName == "key")
                    {
                        topicKey = property.Value.ToString().FormatTopicKey();
                        var propertyValue = properties.FirstOrDefault(prop => prop.Name == "value");
                        if (propertyValue != null)
                            topicValue = propertyValue.Value;
                    }
                    else if (propertyName != "value")
                    {
                        topicKey = propertyName.FormatTopicKey();
                        topicValue = property.Value;
                    }

                    if (!string.IsNullOrWhiteSpace(topicKey))
                        binderDocument.Topics.Add(new ComparisonTopic { Key = topicKey, Result = topicValue });
                }
            }
        }
        else if (binderResults is JObject resultsObject)
        {
            var properties = resultsObject.Properties();
            foreach (var property in properties)
            {
                var topicKey = string.Empty;
                object? topicValue = null;
                var propertyName = property.Name;
                if (propertyName == "name" || propertyName == "key")
                {
                    topicKey = property.Value.ToString().FormatTopicKey();
                    var propertyValue = properties.FirstOrDefault(prop => prop.Name == "value");
                    if (propertyValue != null)
                        topicValue = propertyValue.Value;
                }
                else if (propertyName != "value")
                {
                    topicKey = propertyName.FormatTopicKey();
                    topicValue = property.Value;
                }

                if (!string.IsNullOrWhiteSpace(topicKey))
                    binderDocument.Topics.Add(new ComparisonTopic { Key = topicKey, Result = topicValue });
            }
        }
        if (binderDocument.Topics != null && binderDocument.Topics.Any())
            results.Add(binderDocument);
        #endregion
        #region "Policy Topics"
        var policyResults = topicsDocuments["seconddocument"];
        if (policyResults != null && policyResults is JArray)
        {
            var topicsList = policyResults?.Children();
            foreach (JObject result in topicsList)
            {
                var properties = result.Properties();
                foreach (var property in properties)
                {
                    var topicKey = string.Empty;
                    object? topicValue = null;
                    var propertyName = property.Name;
                    if (propertyName == "name" || propertyName == "key")
                    {
                        topicKey = property.Value.ToString().FormatTopicKey();
                        var propertyValue = properties.FirstOrDefault(prop => prop.Name == "value");
                        if (propertyValue != null)
                            topicValue = propertyValue.Value;
                    }
                    else if (propertyName != "value")
                    {
                        topicKey = propertyName.FormatTopicKey();
                        topicValue = property.Value;
                    }

                    if (!string.IsNullOrWhiteSpace(topicKey))
                        policyDocument.Topics.Add(new ComparisonTopic { Key = topicKey, Result = topicValue });
                }
            }
        }
        else if (policyResults is JObject resultsObject)
        {
            var properties = resultsObject.Properties();
            foreach (var property in properties)
            {
                var topicKey = string.Empty;
                object? topicValue = null;
                var propertyName = property.Name;
                if (propertyName == "name" || propertyName == "key")
                {
                    topicKey = property.Value.ToString().FormatTopicKey();
                    var propertyValue = properties.FirstOrDefault(prop => prop.Name == "value");
                    if (propertyValue != null)
                        topicValue = propertyValue.Value;
                }
                else if (propertyName != "value")
                {
                    topicKey = propertyName.FormatTopicKey();
                    topicValue = property.Value;
                }

                if (!string.IsNullOrWhiteSpace(topicKey))
                    policyDocument.Topics.Add(new ComparisonTopic { Key = topicKey, Result = topicValue });
            }
        }
        if (policyDocument.Topics != null && policyDocument.Topics.Any())
            results.Add(policyDocument);
        #endregion
        #region "Discrepancies"
        var aiDiscrepanciesDocument = new ClaudeDocument { DocumentType = "discrepancies", Topics = new() };
        var discrepancyResults = topicsDocuments["discrepancies"];
        if (discrepancyResults != null && discrepancyResults is JArray)
        {
            var topicsList = discrepancyResults?.Children();
            foreach (JObject result in topicsList)
            {
                var properties = result.Properties();
                foreach (var property in properties)
                {
                    var topicKey = string.Empty;
                    object? topicValue = null;
                    var propertyName = property.Name;
                    if (propertyName == "name" || propertyName == "key" || propertyName == "topic" || propertyName == "field")
                    {
                        topicKey = property.Value.ToString().FormatTopicKey();
                        var propertyValues = new List<object?>();
                        var propertyValue = properties.FirstOrDefault(prop => prop.Name == "firstdocumentValue");
                        if (propertyValue == null)
                            propertyValue = properties.FirstOrDefault(prop => prop.Name == "firstdocument");
                        if (propertyValue == null)
                            propertyValue = properties.FirstOrDefault(prop => prop.Name == "firstValue");
                        propertyValues.Add(propertyValue?.Value);

                        propertyValue = properties.FirstOrDefault(prop => prop.Name == "seconddocumentValue");
                        if (propertyValue == null)
                            propertyValue = properties.FirstOrDefault(prop => prop.Name == "seconddocument");
                        if (propertyValue == null)
                            propertyValue = properties.FirstOrDefault(prop => prop.Name == "secondValue");

                        propertyValues.Add(propertyValue?.Value);
                        topicValue = propertyValues;
                    }
                    else if (!propertyName.Among("firstdocumentValue", "firstdocument", "firstValue", "seconddocumentValue", "seconddocument", "secondValue"))
                    {
                        topicKey = propertyName.FormatTopicKey();
                        topicValue = property.Value;
                    }

                    if (!string.IsNullOrWhiteSpace(topicKey))
                        aiDiscrepanciesDocument.Topics.Add(new ComparisonTopic { Key = topicKey, Value = topicValue });
                }
            }
        }
        else if (discrepancyResults is JObject resultsObject)
        {
            var properties = resultsObject.Properties();
            foreach (var property in properties)
            {
                var topicKey = string.Empty;
                object? topicValue = null;
                var propertyName = property.Name;
                if (propertyName == "name" || propertyName == "key" || propertyName == "topic" || propertyName == "field")
                {
                    topicKey = property.Value.ToString().FormatTopicKey();
                    var propertyValues = new List<object?>();
                    var propertyValue = properties.FirstOrDefault(prop => prop.Name == "firstdocumentValue");
                    if (propertyValue == null)
                        propertyValue = properties.FirstOrDefault(prop => prop.Name == "firstdocument");
                    if (propertyValue == null)
                        propertyValue = properties.FirstOrDefault(prop => prop.Name == "firstValue");
                    propertyValues.Add(propertyValue?.Value);

                    propertyValue = properties.FirstOrDefault(prop => prop.Name == "seconddocumentValue");
                    if (propertyValue == null)
                        propertyValue = properties.FirstOrDefault(prop => prop.Name == "seconddocument");
                    if (propertyValue == null)
                        propertyValue = properties.FirstOrDefault(prop => prop.Name == "secondValue");
                    propertyValues.Add(propertyValue?.Value);
                    topicValue = propertyValues;
                }
                else if (!propertyName.Among("firstdocumentValue", "firstdocument", "firstValue", "seconddocumentValue", "seconddocument", "secondValue"))
                {
                    topicKey = propertyName.FormatTopicKey();
                    topicValue = property.Value;
                }

                if (!string.IsNullOrWhiteSpace(topicKey))
                    aiDiscrepanciesDocument.Topics.Add(new ComparisonTopic { Key = topicKey, Result = topicValue });
            }

            var discrepancyTopics = aiDiscrepanciesDocument.Topics ?? new();
            var allDiscrepancies = new List<Discrepancy>();
            foreach (var discrepancyTopic in discrepancyTopics)
            {
                var topicKey = discrepancyTopic.Key;
                var topicValue = discrepancyTopic.Result;
                if (topicValue is not JObject discrepancyDetails) continue;

                var topicDiscrepancy = new Discrepancy();
                var discrepancyProperties = discrepancyDetails.Properties();

                var binderValue = discrepancyProperties.FirstOrDefault(p => p.Name == "firstdocument");
                if (binderValue != null)
                    topicDiscrepancy.Binder = new ComparisonTopic { Key = topicKey, Result = binderValue.Value };

                var policyValue = discrepancyProperties.FirstOrDefault(p => p.Name == "seconddocument");
                if (policyValue != null)
                    topicDiscrepancy.Policy = new ComparisonTopic { Key = topicKey, Result = policyValue.Value };

                allDiscrepancies.Add(topicDiscrepancy);
            }
            aiDiscrepanciesDocument.Discrepancies = allDiscrepancies.Any() ? allDiscrepancies : null;
        }
        if (aiDiscrepanciesDocument.Discrepancies != null && aiDiscrepanciesDocument.Discrepancies.Any())
            results.Add(aiDiscrepanciesDocument);
        #endregion

        comparisonResult.Documents = results;

        return comparisonResult;
    }
    public static async Task? ToExcelFile(this IConfiguration? configuration, EnvironmentStatistics environmentStatistics, string? templateFile = null)
    {
        if (string.IsNullOrWhiteSpace(templateFile) || !File.Exists(templateFile) ||
            environmentStatistics?.ComparisonResults == null || !environmentStatistics.ComparisonResults.Any())
            return;


        // Open the Excel file
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        using var package = new ExcelPackage(new FileInfo(templateFile));

        var hasChecklists = environmentStatistics.HasChecklists;


        #region "Create the Calculations Worksheet"
        var worksheetNames = environmentStatistics.ComparisonResults.Select(result => result.Key).Distinct().ToList();
        var calculationsWorksheet = package.Workbook.Worksheets.Add("Calculations");
        var row = 1;
        foreach (var worksheetName in worksheetNames)
            calculationsWorksheet.Cells[$"A{row++}"].Value = worksheetName;

        var worksheetNamesRange = $"A1:A{row - 1}";
        calculationsWorksheet.Names.Add("WorksheetNames", calculationsWorksheet.Cells[worksheetNamesRange]);
        calculationsWorksheet.Hidden = eWorkSheetHidden.Hidden;
        #endregion
        #region "Populate the Summary Worksheet"
        var summaryWorkbook = package.Workbook;
        var summaryWorksheet = package.Workbook.Worksheets[0];
        summaryWorksheet.Cells["C1"].Value = configuration?["clientName"] ?? string.Empty;

        summaryWorksheet.Cells["C6"].Value = environmentStatistics?.TotalExpectedValues;
        summaryWorksheet.Cells["D6"].Formula = "TotalCorrectValues";
        summaryWorksheet.Cells["E6"].Formula = "TotalMissed";
        summaryWorksheet.Cells["F6"].Formula = "IFERROR(D6/(C6 + E6), 0)";

        summaryWorksheet.Cells["C7"].Value = environmentStatistics?.TotalExpectedValuesFoundInBinders;
        summaryWorksheet.Cells["D7"].Formula = "TotalCorrectValuesInBinders";
        summaryWorksheet.Cells["E7"].Formula = "TotalMissedInBinders";
        summaryWorksheet.Cells["F7"].Formula = "IFERROR(D7/(C7 + E7), 0)";

        summaryWorksheet.Cells["C8"].Value = environmentStatistics?.TotalExpectedValuesFoundInPolicies;
        summaryWorksheet.Cells["D8"].Formula = "TotalCorrectValuesInPolicies";
        summaryWorksheet.Cells["E8"].Formula = "TotalMissedInPolicies";
        summaryWorksheet.Cells["F8"].Formula = "IFERROR(D8/(C8 + E8), 0)";

        summaryWorksheet.Cells["C9"].Value = environmentStatistics?.TotalDiscrepancies;
        summaryWorksheet.Cells["D9"].Formula = "TotalCorrectDiscrepancies";
        summaryWorksheet.Cells["F9"].Formula = "IFERROR(D9/C9, 0)";

        summaryWorksheet.Cells["C10"].Value = environmentStatistics?.TotalDiscrepancies;
        summaryWorksheet.Cells["D10"].Formula = "TotalIncorrectDiscrepancies";
        summaryWorksheet.Cells["F10"].Formula = "IFERROR(D10/C10, 0)";


        var summaryRowIds = new Dictionary<string, int>
        {
            {"namedInsured", 17 },
            {"mailingAddress", 19 },
            {"policyNumber", 21 },
            {"term", 23 },
            {"entityType", 25 },
            {"market", 27 },
            {"locationSchedule", 29 },
            {"premium", 31 },
            {"mep", 33 },
            {"commission", 35 },
            {"terrorism", 37 },
            {"claims", 39 },
            {"mainCoverages", 47 },
            {"retentionsDeductibles", 49 },
            {"retroDate", 51 },
            {"litigationDate", 53 },
            {"continuityDate", 55 },
            {"additionalInterest", 57 },
            {"additionalCoverageExtensions", 59 }
        };

        var statisticsTopics = environmentStatistics?.Topics ?? new();
        if (statisticsTopics != null)
        {
            row = 1;
            foreach (var statisticTopic in statisticsTopics)
            {
                var statisticsKey = statisticTopic.Key ?? string.Empty;
                if (!summaryRowIds.ContainsKey(statisticsKey)) continue;

                var summaryRow = summaryRowIds[statisticsKey];
                summaryWorksheet.Cells[$"C{summaryRow}"].Value = statisticTopic.BinderTotal;
                summaryWorksheet.Cells[$"D{summaryRow}"].Formula = $"{statisticsKey}CorrectInBinder";
                summaryWorksheet.Cells[$"E{summaryRow}"].Formula = $"{statisticsKey}MissedInBinder";
                summaryWorksheet.Cells[$"F{summaryRow}"].Value = statisticTopic.PolicyTotal;
                summaryWorksheet.Cells[$"G{summaryRow}"].Formula = $"{statisticsKey}CorrectInPolicy";
                summaryWorksheet.Cells[$"H{summaryRow}"].Formula = $"{statisticsKey}MissedInPolicy";
                summaryWorksheet.Cells[$"I{summaryRow}"].Formula = $"IFERROR((D{summaryRow}+G{summaryRow}+E{summaryRow})/(C{summaryRow}+F{summaryRow}+H{summaryRow}),0)";
                summaryWorksheet.Cells[$"J{summaryRow}"].Formula = $"IFERROR(D{summaryRow}/(C{summaryRow}+E{summaryRow}),0)";
                summaryWorksheet.Cells[$"K{summaryRow}"].Formula = $"IFERROR(G{summaryRow}/(F{summaryRow}+H{summaryRow}),0)";
                summaryWorksheet.Cells[$"M{summaryRow}"].Value = statisticTopic.DiscrepancyCount;
                summaryWorksheet.Cells[$"N{summaryRow}"].Formula = $"{statisticsKey}CorrectDiscrepancies";
                summaryWorksheet.Cells[$"O{summaryRow}"].Formula = $"{statisticsKey}IncorrectDiscrepancies";
                if (hasChecklists.GetValueOrDefault())
                {
                    summaryWorksheet.Cells[$"P{summaryRow}"].Formula = $"{statisticsKey}DiscrepanciesInActual";
                    summaryWorksheet.Cells[$"Q{summaryRow}"].Formula = $"{statisticsKey}TotalDiscrepancyValuesMatchedActual";
                    summaryWorksheet.Cells[$"R{summaryRow}"].Formula = $"IFERROR(Q{summaryRow}/P{summaryRow},0)";
                    summaryWorksheet.Cells[$"S{summaryRow}"].Formula = $"{statisticsKey}TotalBinderValuesMatchedActual";
                    summaryWorksheet.Cells[$"T{summaryRow}"].Formula = $"{statisticsKey}TotalPolicyValuesMatchedActual";

                    calculationsWorksheet.Cells[$"H{row}"].Formula = $"SUMPRODUCT(COUNTIFS(INDIRECT(\"'\" & {worksheetNamesRange} & \"'!O42:O61\"), TRUE, INDIRECT(\"'\" & {worksheetNamesRange} & \"'!G42:G61\"), \"{statisticsKey}\"))";
                    summaryWorkbook.Names.Add($"{statisticsKey}DiscrepanciesInActual", calculationsWorksheet.Cells[$"H{row}"]);
                    calculationsWorksheet.Cells[$"I{row}"].Formula = $"SUMPRODUCT(COUNTIFS(INDIRECT(\"'\" & {worksheetNamesRange} & \"'!F42:F61\"),  TRUE, INDIRECT(\"'\" & {worksheetNamesRange} & \"'!G42:G61\"), \"{statisticsKey}\"))";
                    summaryWorkbook.Names.Add($"{statisticsKey}TotalBinderValuesMatchedActual", calculationsWorksheet.Cells[$"I{row}"]);
                    calculationsWorksheet.Cells[$"J{row}"].Formula = $"SUMPRODUCT(COUNTIFS(INDIRECT(\"'\" & {worksheetNamesRange} & \"'!N42:N61\"),  TRUE, INDIRECT(\"'\" & {worksheetNamesRange} & \"'!G42:G61\"), \"{statisticsKey}\"))";
                    summaryWorkbook.Names.Add($"{statisticsKey}TotalPolicyValuesMatchedActual", calculationsWorksheet.Cells[$"J{row}"]);
                    calculationsWorksheet.Cells[$"K{row}"].Formula = $"SUMPRODUCT(COUNTIFS(INDIRECT(\"'\" & {worksheetNamesRange} & \"'!S42:S61\"),  TRUE, INDIRECT(\"'\" & {worksheetNamesRange} & \"'!G42:G61\"), \"{statisticsKey}\"))";
                    summaryWorkbook.Names.Add($"{statisticsKey}TotalDiscrepancyValuesMatchedActual", calculationsWorksheet.Cells[$"K{row}"]);

                }

                // Create the associated named formaulas that are referenced in the summary worksheet
                calculationsWorksheet.Cells[$"B{row}"].Formula = $"SUMPRODUCT(COUNTIFS(INDIRECT(\"'\" & {worksheetNamesRange} & \"'!D11:D22\"),  TRUE, INDIRECT(\"'\" & {worksheetNamesRange} & \"'!G11:G22\"), \"{statisticsKey}\"))";
                summaryWorkbook.Names.Add($"{statisticsKey}CorrectInBinder", calculationsWorksheet.Cells[$"B{row}"]);
                calculationsWorksheet.Cells[$"C{row}"].Formula = $"SUMPRODUCT(COUNTIFS(INDIRECT(\"'\" & {worksheetNamesRange} & \"'!L11:L22\"),  TRUE, INDIRECT(\"'\" & {worksheetNamesRange} & \"'!G11:G22\"), \"{statisticsKey}\"))";
                summaryWorkbook.Names.Add($"{statisticsKey}CorrectInPolicy", calculationsWorksheet.Cells[$"C{row}"]);
                calculationsWorksheet.Cells[$"D{row}"].Formula = $"SUMPRODUCT(COUNTIFS(INDIRECT(\"'\" & {worksheetNamesRange} & \"'!E11:E22\"),  TRUE, INDIRECT(\"'\" & {worksheetNamesRange} & \"'!G11:G22\"), \"{statisticsKey}\"))";
                summaryWorkbook.Names.Add($"{statisticsKey}MissedInBinder", calculationsWorksheet.Cells[$"D{row}"]);
                calculationsWorksheet.Cells[$"E{row}"].Formula = $"SUMPRODUCT(COUNTIFS(INDIRECT(\"'\" & {worksheetNamesRange} & \"'!M11:M22\"),  TRUE, INDIRECT(\"'\" & {worksheetNamesRange} & \"'!G11:G22\"), \"{statisticsKey}\"))";
                summaryWorkbook.Names.Add($"{statisticsKey}MissedInPolicy", calculationsWorksheet.Cells[$"E{row}"]);
                calculationsWorksheet.Cells[$"F{row}"].Formula = $"SUMPRODUCT(COUNTIFS(INDIRECT(\"'\" & {worksheetNamesRange} & \"'!L42:L61\"),  TRUE, INDIRECT(\"'\" & {worksheetNamesRange} & \"'!G42:G61\"), \"{statisticsKey}\"))";
                summaryWorkbook.Names.Add($"{statisticsKey}CorrectDiscrepancies", calculationsWorksheet.Cells[$"F{row}"]);
                calculationsWorksheet.Cells[$"G{row}"].Formula = $"SUMPRODUCT(COUNTIFS(INDIRECT(\"'\" & {worksheetNamesRange} & \"'!L42:L61\"),  FALSE, INDIRECT(\"'\" & {worksheetNamesRange} & \"'!G42:G61\"), \"{statisticsKey}\"))";
                summaryWorkbook.Names.Add($"{statisticsKey}IncorrectDiscrepancies", calculationsWorksheet.Cells[$"G{row}"]);

                row++;
            }

            calculationsWorksheet.Cells[$"B{row}"].Formula = $"=SUM(B1:B{row - 1})";
            summaryWorkbook.Names.Add($"TotalCorrectValuesInBinders", calculationsWorksheet.Cells[$"B{row}"]);
            calculationsWorksheet.Cells[$"C{row}"].Formula = $"=SUM(C1:C{row - 1})";
            summaryWorkbook.Names.Add($"TotalCorrectValuesInPolicies", calculationsWorksheet.Cells[$"C{row}"]);
            calculationsWorksheet.Cells[$"D{row}"].Formula = $"=SUM(D1:D{row - 1})";
            summaryWorkbook.Names.Add($"TotalMissedInBinders", calculationsWorksheet.Cells[$"D{row}"]);
            calculationsWorksheet.Cells[$"E{row}"].Formula = $"=SUM(E1:E{row - 1})";
            summaryWorkbook.Names.Add($"TotalMissedInPolicies", calculationsWorksheet.Cells[$"E{row}"]);
            calculationsWorksheet.Cells[$"F{row}"].Formula = $"=SUM(F1:F{row - 1})";
            summaryWorkbook.Names.Add($"TotalCorrectDiscrepancies", calculationsWorksheet.Cells[$"F{row}"]);
            calculationsWorksheet.Cells[$"G{row}"].Formula = $"=SUM(G1:G{row - 1})";
            summaryWorkbook.Names.Add($"TotalIncorrectDiscrepancies", calculationsWorksheet.Cells[$"G{row}"]);

            calculationsWorksheet.Cells[$"C{row + 1}"].Formula = $"=SUM(TotalCorrectValuesInBinders, TotalCorrectValuesInPolicies)";
            summaryWorkbook.Names.Add($"TotalCorrectValues", calculationsWorksheet.Cells[$"C{row + 1}"]);

            calculationsWorksheet.Cells[$"E{row + 1}"].Formula = $"=SUM(TotalMissedInBinders, TotalMissedInPolicies)";
            summaryWorkbook.Names.Add($"TotalMissed", calculationsWorksheet.Cells[$"E{row + 1}"]);

            calculationsWorksheet.Cells[$"G{row + 1}"].Formula = $"=SUM(TotalCorrectDiscrepancies, TotalIncorrectDiscrepancies)";
            summaryWorkbook.Names.Add($"TotalCalculatedDiscrepancies", calculationsWorksheet.Cells[$"G{row + 1}"]);
        }
        #endregion
        #region "Create Results Worksheet for each Comparison Result"
        var comparisonTemplateWorksheet = package.Workbook.Worksheets[1];

        var topicRowIds = new Dictionary<string, int>
        {
            {"namedInsured", 11 },
            {"mailingAddress", 12 },
            {"policyNumber", 13 },
            {"term", 14 },
            {"entityType", 15 },
            {"market", 16 },
            {"locationSchedule", 17 },
            {"premium", 18 },
            {"mep", 19 },
            {"commission", 20 },
            {"terrorism", 21 },
            {"claims", 22 },
            {"mainCoverages", 29 },
            {"retentionsDeductibles", 30 },
            {"retroDate", 31 },
            {"litigationDate", 32 },
            {"continuityDate", 33 },
            {"additionalInterest", 34 },
            {"additionalCoverageExtensions", 35 }
        };
        var discrepanciesRowIds = new Dictionary<string, int>
        {
            {"namedInsured", 42 },
            {"mailingAddress", 43 },
            {"policyNumber", 44 },
            {"term", 45 },
            {"entityType", 46 },
            {"market", 47 },
            {"locationSchedule", 48 },
            {"premium", 49 },
            {"mep", 50 },
            {"commission", 51 },
            {"terrorism", 52 },
            {"claims", 53 },
            {"mainCoverages", 55 },
            {"retentionsDeductibles", 56 },
            {"retroDate", 57 },
            {"litigationDate", 58 },
            {"continuityDate", 59 },
            {"additionalInterest", 60 },
            {"additionalCoverageExtensions", 61 }
        };

        var comparisonResults = environmentStatistics?.ComparisonResults ?? new();
        foreach (var comparisonResult in comparisonResults)
        {
            #region "Copy the Comparison Template Worksheet"
            var comparisonResultWorksheet = package.Workbook.Worksheets[comparisonResult.Key] ??
                package.Workbook.Worksheets.Add(comparisonResult.Key, comparisonTemplateWorksheet);
            #endregion
            #region "Overall Comparison Result Details"
            comparisonResultWorksheet.Name = comparisonResult.Key;
            comparisonResultWorksheet.Cells["B1"].Value = configuration?["keyName"] ?? "Comparison:";
            comparisonResultWorksheet.Cells["C1"].Value = comparisonResult?.Key ?? "Not Specified";
            #endregion
            #region "Binder Values Found"
            var binderDocument = comparisonResult?.Documents?.FirstOrDefault(document => document.DocumentType == "binder");
            if (binderDocument != null)
            {
                var binderTopics = binderDocument.Topics ?? new();
                foreach (var binderTopic in binderTopics)
                {
                    if (!topicRowIds.ContainsKey(binderTopic.Key ?? string.Empty)) continue;

                    var binderResultString = binderTopic.ToResultsString();
                    if (!string.IsNullOrWhiteSpace(binderResultString) && binderResultString != "Nothing")
                        comparisonResultWorksheet.Cells[$"C{topicRowIds[binderTopic.Key ?? string.Empty]}"].Value = binderResultString;
                }
            }
            #endregion
            #region "Policy Values Found"
            var policyDocument = comparisonResult?.Documents?.FirstOrDefault(document => document.DocumentType == "policy");
            if (policyDocument != null)
            {
                var policyTopics = policyDocument.Topics ?? new();
                foreach (var policyTopic in policyTopics)
                {
                    if (!topicRowIds.ContainsKey(policyTopic.Key ?? string.Empty)) continue;

                    var policyResultString = policyTopic.ToResultsString();
                    if (!string.IsNullOrWhiteSpace(policyResultString) && policyResultString != "Nothing")
                        comparisonResultWorksheet.Cells[$"J{topicRowIds[policyTopic.Key ?? string.Empty]}"].Value = policyResultString;
                }
            }
            #endregion
            #region "Discrepancies"
            var checklistDocument = comparisonResult?.Documents?.FirstOrDefault(document => document.DocumentType == "checklist");
            var checklistDiscrepancies = checklistDocument?.Discrepancies ?? new();

            var comparisonDiscrepancies = comparisonResult?.Discrepancies ?? new();
            foreach (var discrepancy in comparisonDiscrepancies)
            {
                if (!discrepanciesRowIds.ContainsKey(discrepancy.Key ?? string.Empty)) continue;
                var discrepancyKey = discrepancy.Key;
                try
                {
                    comparisonResultWorksheet.Cells[$"L{discrepanciesRowIds[discrepancyKey ?? string.Empty]}"].Value = discrepancy?.IsCorrect;
                    if (!discrepancy?.IsCorrect ?? false)
                    {
                        var messages = discrepancy?.Messages;
                        if (messages != null)
                        {
                            var commentMessage = string.Join(Environment.NewLine, messages.ToArray());

                            var discrepancyCell = comparisonResultWorksheet.Cells[$"B{discrepanciesRowIds[discrepancyKey ?? string.Empty]}"];
                            if (discrepancyCell.Comment != null)
                                comparisonResultWorksheet.Comments.Remove(discrepancyCell.Comment);
                            discrepancyCell.AddComment(commentMessage);
                        }
                    }

                    var binderValue = discrepancy?.Binder?.ToResultsString();
                    if (!string.IsNullOrWhiteSpace(binderValue) && binderValue != "Nothing")
                        comparisonResultWorksheet.Cells[$"C{discrepanciesRowIds[discrepancyKey ?? string.Empty]}"].Value = binderValue;
                    var policyValue = discrepancy?.Policy?.ToResultsString();
                    if (!string.IsNullOrWhiteSpace(policyValue) && policyValue != "Nothing")
                        comparisonResultWorksheet.Cells[$"I{discrepanciesRowIds[discrepancyKey ?? string.Empty]}"].Value = policyValue;

                    comparisonResultWorksheet.Cells[$"F{discrepanciesRowIds[discrepancyKey ?? string.Empty]}"].Value = discrepancy?.BinderMatchesActual;
                    comparisonResultWorksheet.Cells[$"N{discrepanciesRowIds[discrepancyKey ?? string.Empty]}"].Value = discrepancy?.PolicyMatchesActual;
                    comparisonResultWorksheet.Cells[$"S{discrepanciesRowIds[discrepancyKey ?? string.Empty]}"].Value = discrepancy?.MatchesActual;

                    var actualDiscrepancy = checklistDiscrepancies.FirstOrDefault(cd => cd.Key == discrepancyKey);
                    if (actualDiscrepancy != null)
                    {
                        comparisonResultWorksheet.Cells[$"O{discrepanciesRowIds[discrepancyKey ?? string.Empty]}"].Value = true;
                        comparisonResultWorksheet.Cells[$"P{discrepanciesRowIds[discrepancyKey ?? string.Empty]}"].Value = actualDiscrepancy.Binder?.ToResultsString();
                        comparisonResultWorksheet.Cells[$"Q{discrepanciesRowIds[discrepancyKey ?? string.Empty]}"].Value = actualDiscrepancy.Policy?.ToResultsString();
                        comparisonResultWorksheet.Cells[$"R{discrepanciesRowIds[discrepancyKey ?? string.Empty]}"].Value = actualDiscrepancy.Other?.ToResultsString();
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Error processing topic: {discrepancyKey} from results file {comparisonResult.Key} Received Exception {e.Message}");
                }
            }
            #endregion
        }

        comparisonTemplateWorksheet.Hidden = eWorkSheetHidden.Hidden;
        #endregion
        #region "Assign Named Ranges"
        #endregion
        #region "Save Excel File"
        var excelFileName = environmentStatistics?.FileName ?? string.Empty;
        if (string.IsNullOrWhiteSpace(excelFileName))
            return;

        var file = new FileInfo(excelFileName);
        await package.SaveAsAsync(file);
        #endregion

    }
    public static decimal? CalculateCost(this string? textToCalculate, decimal? tokenRate)
    {
        if (string.IsNullOrWhiteSpace(textToCalculate)) return null;

        var tokens = Regex.Matches(textToCalculate, @"\W+");
        var tokenCount = tokens.Count;
        if (tokenCount == 0) return null;

        return Convert.ToDecimal(tokenCount) / 1000000 * tokenRate;
    }
    public static decimal? CalculateMessageCost(this string? messageJson, ClaudeModel? modelUsed)
    {
        if (string.IsNullOrWhiteSpace(messageJson) || modelUsed == null) return null;

        // Parse the JSON string
        var jsonObject = JObject.Parse(messageJson);

        List<string> tokens = new List<string>();

        // Tokenize the JSON properties recursively
        TokenizeObject(jsonObject, tokens);

        var tokenCount = tokens.Count;
        if (tokenCount == 0) return null;

        return Convert.ToDecimal(tokenCount) / modelUsed.PerTokenAmount;
    }
    public static string? ToPlaceholder(this string? key) => $"#o#o{key?.ToUpper()}#e#e".Replace("#o", "{").Replace("#e", "}");
    public static string ToSafeString(this object? item)
    {
        if (item == null)
            return string.Empty;
        var itemString = item.ToString();
        return string.IsNullOrWhiteSpace(itemString) ? string.Empty : itemString.Trim();
    }
    public static object? GetValue(this ClaudeDocument document, string? key) => document?.Topics?.FirstOrDefault(topic => topic.Key == key)?.Value;
    public static string? GetTopics(this ClaudePrompt prompt)
    {
        if (prompt == null || string.IsNullOrWhiteSpace(prompt.TopicsFile)) return null;

        var assembly = Assembly.GetExecutingAssembly();
        var stream =
            assembly.GetManifestResourceStream(
                $"AmWins.Architecture.Anthropic.Tasks.Resources.{prompt.TopicsFile}");
        if (stream == null)
            return null;

        using var sr = new StreamReader(stream);
        return sr.ReadToEnd();
    }

    #endregion
}
