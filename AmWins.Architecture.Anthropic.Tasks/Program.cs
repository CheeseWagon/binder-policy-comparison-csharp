#region "Usings"
using System.Dynamic;
using System.Text.RegularExpressions;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.AzureAppConfiguration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;

using AmWins.Architecture.Anthropic.Tasks;
using AmWins.Architecture.Anthropic.Tasks.Models;
using OfficeOpenXml;
using System;
#endregion
#region "Constants"
const int TOTAL_MENU_ITEMS = 3;
#endregion
#region "Member Variables"
Dictionary<string, string> processingLog = new();
ILogger Log;
IConfiguration Configuration;
IServiceProvider Services;
#endregion
#region "Main"
SetUpLogging();
SetUpConfiguration();
SetUpDependencyInjection();

int userInput;
do
{
    userInput = DisplayMenu();
    PerformAction(userInput);

} while (userInput != TOTAL_MENU_ITEMS);
#endregion
#region "Main Menu"
int DisplayMenu()
{
    Console.Clear();
    Console.WriteLine(@"Anthropic Claude AI PoC Tasks");
    Console.WriteLine();

    Console.WriteLine(@"01. Compare Binder/Policy Documents.");
    Console.WriteLine(@"02. Generate Statistics Spreadsheet.");

    Console.WriteLine();
    Console.WriteLine(@"03. Exit");

    var result = Console.ReadLine();
    return int.TryParse(result, out var returnValue) && (returnValue >= 1 && returnValue <= TOTAL_MENU_ITEMS) ? returnValue : 0;
}
void PerformAction(int userInput)
{
    Console.Clear();
    switch (userInput)
    {
        case 0:
            Console.WriteLine(@"Invalid Response.  Please select a value from the displayed menu.");
            PressReturn();
            break;
        case 1:
            ComparePoliciesAndBinders().Wait();
            break;
        case 2:
            AnalyzeComparisonResults().Wait();
            break;

    }
    #region "Obsolete Switch Statement"
    //switch (userInput)
    //{
    //    case 0:
    //        Console.WriteLine(@"Invalid Response.  Please select a value from the displayed menu.");
    //        PressReturn();
    //        break;
    //    case 1:
    //        CompareExistingChecklistsToAIComparisonResults().Wait();
    //        break;
    //    case 2:
    //        ExportDocumentComparisonsToExcelSpreadsheets().Wait();
    //        break;
    //    case 3:
    //        ParseResultsOfProcessedEnvironments().Wait();
    //        break;
    //    case 4:
    //        ComparePolicyBindersToSavedPolicies().Wait();
    //        break;
    //    case 5:
    //        RetryInitialSubmissions().Wait();
    //        break;
    //    case 6:
    //        ExportTargetedQuestionResultsToExcel().Wait();
    //        break;
    //    case 7:
    //        AnalyzeAiTargetedQuestionResults().Wait();
    //        break;
    //    case 8:
    //        AnalyzeWithTargetedQuestions().Wait();
    //        break;
    //    case 9:
    //        AnalyzeApplicationsWithImprovedPrompts().Wait();
    //        break;
    //    case 10:
    //        ExportResultsWithValuesToExcel().Wait();
    //        break;
    //    case 11:
    //        AnalyzeAiResults().Wait();
    //        break;
    //    case 12:
    //        CompareSubmissionFileToPdf().Wait();
    //        break;
    //    case 13:
    //        ExportQuestionsTemplate().Wait();
    //        break;
    //    case 14:
    //        ImportProfessionalLinesData().Wait();
    //        break;
    //    case 15:
    //        ImportAmLinkSubmissions().Wait();
    //        break;
    //    case 16:
    //        DetermineWhichQuestionsHaveAnswers().Wait();
    //        break;
    //    case 17:
    //        GeneratePromptsBasedOnSubmissionAnswers().Wait();
    //        break;
    //    case 18:
    //        ExportPromptTemplates();
    //        break;
    //    case 19:
    //        ExportTargetedQuestionResultsToExcelSampling().Wait();
    //        break;
    //    case 20:
    //        DocumentIntelligenceExtraction().Wait();
    //        break;
    //    case 21:
    //        CopyApplicationsToAssociatedMarkets().Wait();
    //        break;
    //    case 22:
    //        ExportApplicationSampleBySpecifiedAmount().Wait();
    //        break;
    //    case 23:
    //        AnalyzeChecklists().Wait();
    //        break;
    //    case 24:
    //        SaveCyberResultsToExcel().Wait();
    //        break;
    //}
    #endregion

}
#endregion
#region "Implementation"
void SetUpLogging()
{
    using var loggerFactory = LoggerFactory.Create(builder =>
    {
        builder
            .AddFilter("Microsoft", LogLevel.Warning)
            .AddFilter("System", LogLevel.Warning)
            .AddFilter("InitializeAmWinsConnectDatabases.Program", LogLevel.Debug)
            .AddConsole();
    });
    Log = loggerFactory.CreateLogger<Program>();
}
void SetUpConfiguration()
{
    var builder = new ConfigurationBuilder()
        .SetBasePath(Path.Combine(Directory.GetCurrentDirectory()))
        .AddJsonFile("appsettings.json");

    Configuration = builder.Build();
}
void SetUpDependencyInjection()
{
    var services = new ServiceCollection();

    // Configuration
    services.AddSingleton(_ => Configuration);
    services.AddHttpClient();

    Services = services.BuildServiceProvider();
}
void PressReturn(string? message = null)
{
    message ??= @"Press RETURN to continue...";
    Console.WriteLine(message);
    Console.ReadLine();
}
void PrintLog()
{
    Console.Clear();

    if (processingLog == null)
        return;

    foreach (var (key, value) in processingLog)
    {
        Console.WriteLine(value);
    }
}
string? GetPolicyDocumentWithMatchSubmissionFileMarketId(string? bindersDirectory, string binderFile)
{
    const string binderFilePattern = @"(?<submissionFileMarketId>.*?)\s-\s(?<documentNumber>.*?)\s-\sPOL\s-\sBinder.[pdf|PDF]";

    var binderFileName = Path.GetFileName(binderFile);
    var isValidBinderFile = Regex.Match(binderFileName, binderFilePattern, RegexOptions.Singleline);
    if (!isValidBinderFile.Success) return null;

    if (string.IsNullOrWhiteSpace(bindersDirectory)) return null;

    const string policyFilePattern = @"(?<submissionFileMarketId>.*?)\s-\s(?<documentNumber>.*?)\s-\sPOL\s-\sPolicy.[pdf|PDF]";

    var policyFiles = Directory.GetFiles(bindersDirectory, "*POL - Policy.pdf").ToList();
    foreach (var policyFile in policyFiles)
    {
        var policyFileName = Path.GetFileName(policyFile);
        string submissionFileMarketId = isValidBinderFile.Groups["submissionFileMarketId"].Value;
        var isValidPolicyFile = Regex.Match(policyFileName, policyFilePattern, RegexOptions.Singleline);
        if (!isValidPolicyFile.Success) continue;

        if (string.Compare(submissionFileMarketId, isValidPolicyFile.Groups["submissionFileMarketId"].Value, StringComparison.InvariantCultureIgnoreCase) == 0)
            return policyFile;
    }

    return null;
}
#endregion
#region "Actions"
async Task ComparePoliciesAndBinders()
{
    var clientName = Configuration["clientName"];
    var bindersDirectory = Configuration["bindersDirectory"];
    if (string.IsNullOrWhiteSpace(bindersDirectory))
        return;

    const string resultsWithChecklistPattern = @"(?<submissionFileId>.*?)\s-\s(?<documentNumber>.*?)\s-\sPOL\s-\sResults.[json|JSON|Json]";

    OnDisplayMessage(null, new MessageEventArgs { LogKey = clientName, Message = $"Retrieving Binder Documents for {clientName}" });
    var binderFiles = Directory.GetFiles(bindersDirectory, "* Binder.pdf").ToList();

    var index = 0;
    var totalFiles = binderFiles.Count;

    var templateFile = $"{bindersDirectory}Results Template.xlsx";
    foreach (var binderFile in binderFiles)
    {
        var resultsDocuments = new List<ClaudeDocument>();

        var binderFileName = Path.GetFileName(binderFile);
        OnDisplayMessage(null, new MessageEventArgs { LogKey = clientName, Message = $"Processing {++index} of {totalFiles} Document Comparisions. Current Binder: {binderFileName} for {clientName ?? string.Empty}" });

        #region "Get Related Documents"
        var binderDocument = new ClaudeDocument { FileName = binderFile, DocumentType = "binder", Topics = new() };
        var binderExtractedTextFile = binderFile.Replace(".pdf", " - Extracted Text.txt");
        if (File.Exists(binderExtractedTextFile))
            binderDocument.ExtractedText = await File.ReadAllTextAsync(binderExtractedTextFile);
        else
        {
            Console.WriteLine($"Binder Extracted Text File Not Found: {binderExtractedTextFile}");
            continue;
        }

        string? policyFileName = binderFile.Replace("Binder.pdf", "Policy.pdf");
        var policyFile = policyFileName;
        if (File.Exists(policyFile))
            policyFileName = Path.GetFileName(policyFile);
        else
        {
            policyFile = GetPolicyDocumentWithMatchSubmissionFileMarketId(bindersDirectory, binderFile);
            if (File.Exists(policyFile))
                policyFileName = Path.GetFileName(policyFile);
        }
        var policyDocument = new ClaudeDocument { FileName = policyFile, DocumentType = "policy", Topics = new() };
        var policyExtractedTextFile = policyFile?.Replace(".pdf", " - Extracted Text.txt");
        if (File.Exists(policyExtractedTextFile))
            policyDocument.ExtractedText = await File.ReadAllTextAsync(policyExtractedTextFile);
        else
        {
            Console.WriteLine($"Policy Extracted Text File Not Found: {policyExtractedTextFile}");
            continue;
        }
        #endregion
        #region "Perform Document Comparison"
        var comparsonRequest = new DocumentComparisonRequest { Documents = new() { binderDocument, policyDocument }, SubmitRequestToClaude = true };
        var documentComparison = await Configuration.CompareDocuments(comparsonRequest);
        if (documentComparison == null) continue;

        var comparisonFile = binderFile.Replace("Binder.pdf", "Comparison Response.json");
        var promptsExecuted = documentComparison.PromptsExecuted ?? new();
        foreach (var promptExecuted in promptsExecuted)
        {
            var modelsUsed = promptExecuted.ModelsUsed;
            if (modelsUsed == null || !modelsUsed.Any()) continue;

            foreach (var modelUsed in modelsUsed)
            {
                var modelUsedFileName = modelUsed.Model ?? "Unknown";
                var modelUsedFile = comparisonFile.Replace("Comparison Response.json", $"{promptExecuted.Name} - {modelUsedFileName} - Request.json");
                modelUsed.ToJsonFile(modelUsedFile);
                modelUsed.RequestMessage = null;
            }
        }
        documentComparison.ToJsonFile(comparisonFile);
        #endregion
        #region "Parse Comparison Results"
        var comparisonResultFileName = Path.GetFileName(comparisonFile);

        var comparisonResult = await comparisonFile.LoadJsonFromFile<ClaudeResponse>();
        if (comparisonResult?.PromptsExecuted == null || !comparisonResult.PromptsExecuted.Any()) continue;

        dynamic? results = new ExpandoObject();
        var actualResultsFileName = comparisonFile.Replace("Comparison Response.json", "Results.json");
        foreach (var promptExecuted in promptsExecuted)
        {
            var promptResult = promptExecuted.ModelsUsed?[0]?.Result;
            if (promptResult == null) continue;

            string resultsBody = @"{" + ((char)34).ToString() + "result" + ((char)34).ToString() + ": " + JsonConvert.SerializeObject(promptResult) + "}";
            try
            {
                var actualResponse = resultsBody.ExtractActualResponse();
                if (actualResponse == null) continue;

                if (promptExecuted.TopicsFile == "commonDeclarations.json")
                    results.Common = actualResponse;
                else if (promptExecuted.TopicsFile == "cyberDeclarations.json")
                    results.Cyber = actualResponse;
                else if (promptExecuted.TopicsFile == "commonAndCyberDeclarations.json")
                    results.Both = actualResponse;

                var exportDirectory = Path.GetDirectoryName(actualResultsFileName);
                if (exportDirectory == null) continue;

                if (!Directory.Exists(exportDirectory))
                    Directory.CreateDirectory(exportDirectory);

                File.WriteAllText(actualResultsFileName,
                    JsonConvert.SerializeObject(results,
                        new JsonSerializerSettings
                        {
                            ContractResolver = new CamelCasePropertyNamesContractResolver(),
                            NullValueHandling = NullValueHandling.Ignore
                        }));
            }
            catch (Exception e)
            {
                Console.WriteLine($"Failed to extract Response From: {comparisonFile} - Prompt: {promptExecuted.Name}. Received Message: {e.Message}.");
            }
        }
        #endregion
        #region "Process Results"
        if (!File.Exists(actualResultsFileName)) continue;

        var actualResult = await actualResultsFileName.LoadJsonFromFile<dynamic>();
        if (actualResult == null) continue;

        dynamic topicsDocuments = actualResult["both"] ?? actualResult["common"] ?? actualResult["cyber"];
        if (topicsDocuments == null) continue;

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
            resultsDocuments.Add(binderDocument);
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
            resultsDocuments.Add(policyDocument);
        #endregion
        #region "Discrepancies"
        var aiDiscrepanciesDocument = new ClaudeDocument { FileName = binderFile.Replace("Binder.pdf", "Checklist Results.xlsx"), DocumentType = "discrepancies", Topics = new() };
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
            resultsDocuments.Add(aiDiscrepanciesDocument);
        #endregion
        #region "Existing Checklist"
        const string binderWithChecklistPattern = @"(?<submissionFileId>.*?)\s-\s(?<documentNumber>.*?)\s-\sPOL\s-\sBinder.[pdf|PDF|Pdf]";
        var isValidMatch = Regex.Match(binderFileName, binderWithChecklistPattern, RegexOptions.Singleline);
        if (isValidMatch.Success)
        {
            var submissionFileId = Convert.ToInt32(isValidMatch.Groups["submissionFileId"].Value);
            var documentNumber = isValidMatch.Groups["documentNumber"].Value;
            var checkListFileName = $"{bindersDirectory}{submissionFileId} - Checklist - {documentNumber}.xlsx";
            if (File.Exists(checkListFileName))
            {
                var checkListDocument = GetMarkedDiscrpanciesFromExcelFile(checkListFileName);
                if (checkListDocument != null)
                {
                    checkListDocument.SubmissionFileId = submissionFileId;
                    checkListDocument.DocumentNumber = documentNumber;
                    resultsDocuments.Add(checkListDocument);
                }
            }
        }
        #endregion

        // TODO: Compile statistics on the discrepancies found
        // TODO: Save the results to an Excel file
        await SaveChecklistResultsExcelFile(templateFile, resultsDocuments);

        #endregion
    }

    OnDisplayMessage(null, new MessageEventArgs { LogKey = clientName, Message = $"Retrieving Results Documents for {clientName}" });
    var resultsFiles = Directory.GetFiles(bindersDirectory, "* Results.json").ToList();
    if (resultsFiles == null || !resultsFiles.Any()) return;

    var fileKeyName = Configuration["keyName"];
    var fileKeyPattern = Configuration["keyPattern"] ?? string.Empty;
    var environmentResults = new List<DocumentComparisonResult>();
    foreach (var resultsFile in resultsFiles)
    {
        var resultsFileName = Path.GetFileName(resultsFile);

        var comparisonResult = await resultsFile.GetResultsFromFile() ?? new();
        switch (comparisonResult.TopicsFile)
        {
            case "commonAndCyberDeclarations.json":
                {
                    var commonTopicsObject = JsonConvert.DeserializeObject<dynamic>(Constants.CommonDeclarations ?? string.Empty);
                    comparisonResult.Topics = JsonConvert.DeserializeObject<List<ComparisonTopic>>(JsonConvert.SerializeObject(commonTopicsObject?.topics));
                    var cyberTopicsObject = JsonConvert.DeserializeObject<dynamic>(Constants.CyberDeclarations ?? string.Empty);
                    comparisonResult.Topics.AddRange(JsonConvert.DeserializeObject<List<ComparisonTopic>>(JsonConvert.SerializeObject(cyberTopicsObject?.topics)));
                }
                break;
            case "commonDeclarations.json":
                {
                    var commonTopicsObject = JsonConvert.DeserializeObject<dynamic>(Constants.CommonDeclarations ?? string.Empty);
                    comparisonResult.Topics = JsonConvert.DeserializeObject<List<ComparisonTopic>>(JsonConvert.SerializeObject(commonTopicsObject?.topics));
                }
                break;
            case "cyberDeclarations.json":
                {
                    var cyberTopicsObject = JsonConvert.DeserializeObject<dynamic>(Constants.CyberDeclarations ?? string.Empty);
                    comparisonResult.Topics = JsonConvert.DeserializeObject<List<ComparisonTopic>>(JsonConvert.SerializeObject(cyberTopicsObject?.topics));
                }
                break;
        }


        var resultsKeyPattern = fileKeyPattern.Replace("Binder", "Results").Replace("[pdf|PDF|Pdf]", "json");
        if (!string.IsNullOrWhiteSpace(resultsKeyPattern) && !string.IsNullOrWhiteSpace(resultsFileName))
        {
            var isValidMatch = Regex.Match(resultsFileName, resultsKeyPattern, RegexOptions.Singleline);
            if (isValidMatch.Success)
            {
                comparisonResult.Key = isValidMatch.Groups["key"].Value;
                comparisonResult.KeyName = fileKeyName ?? "Id:";
            }
        }

        // Get the Binder File Name and Extracted Text
        var binderFile = resultsFile.Replace("Results.json", "Binder.pdf");
        var binderDocument = comparisonResult.Documents?.FirstOrDefault(doc => doc.DocumentType == "binder") ??
                                new ClaudeDocument { FileName = binderFile, DocumentType = "binder", Topics = new() };
        binderDocument.FileName ??= binderFile;
        var binderExtractedTextFile = binderFile.Replace(".pdf", " - Extracted Text.txt");
        if (File.Exists(binderExtractedTextFile))
            binderDocument.ExtractedText = await File.ReadAllTextAsync(binderExtractedTextFile);

        // Get the Policy File Name and Extracted Text
        var policyFile = resultsFile.Replace("Results.json", "Policy.pdf");
        var policyDocument = comparisonResult.Documents?.FirstOrDefault(doc => doc.DocumentType == "policy") ??
                                new ClaudeDocument { FileName = policyFile, DocumentType = "policy", Topics = new() };
        policyDocument.FileName ??= policyFile;
        var policyExtractedTextFile = policyFile.Replace(".pdf", " - Extracted Text.txt");
        if (File.Exists(policyExtractedTextFile))
            policyDocument.ExtractedText = await File.ReadAllTextAsync(policyExtractedTextFile);

        #region "Existing Checklist"
        var hasChecklistMatch = Regex.Match(resultsFileName, resultsWithChecklistPattern, RegexOptions.Singleline);
        if (hasChecklistMatch.Success)
        {
            var submissionFileId = Convert.ToInt32(hasChecklistMatch.Groups["submissionFileId"].Value);
            var documentNumber = hasChecklistMatch.Groups["documentNumber"].Value;
            var checkListFileName = $"{bindersDirectory}{submissionFileId} - Checklist - {documentNumber}.xlsx";
            if (File.Exists(checkListFileName))
            {
                var checkListDocument = GetMarkedDiscrpanciesFromExcelFile(checkListFileName);
                if (checkListDocument != null)
                {
                    checkListDocument.SubmissionFileId = submissionFileId;
                    checkListDocument.DocumentNumber = documentNumber;
                    comparisonResult.Documents?.Add(checkListDocument);
                }
            }
        }
        #endregion

        environmentResults.Add(comparisonResult);
    }

    // Compile totals and accuracy of the document comparisons
    var environmentStatistics = Configuration?.Analyze(environmentResults);
    if (environmentStatistics == null) return;

    environmentStatistics.FileName = $"{bindersDirectory}{clientName} - Statistics.xlsx";
    await Configuration?.ToExcelFile(environmentStatistics, templateFile);

    Console.WriteLine();
    PressReturn();
}
async Task AnalyzeComparisonResults()
{
    var clientName = Configuration["clientName"];
    var bindersDirectory = Configuration["bindersDirectory"];
    if (string.IsNullOrWhiteSpace(bindersDirectory))
        return;
    const string resultsWithChecklistPattern = @"(?<submissionFileId>.*?)\s-\s(?<documentNumber>.*?)\s-\sPOL\s-\sResults.[json|JSON|Json]";

    OnDisplayMessage(null, new MessageEventArgs { LogKey = clientName, Message = $"Retrieving Results Documents for {clientName}" });
    var resultsFiles = Directory.GetFiles(bindersDirectory, "* Results.json").ToList();
    if (resultsFiles == null || !resultsFiles.Any()) return;

    var templateFile = $"{bindersDirectory}Results Template.xlsx";
    var fileKeyName = Configuration["keyName"];
    var fileKeyPattern = Configuration["keyPattern"] ?? string.Empty;
    var environmentResults = new List<DocumentComparisonResult>();
    foreach (var resultsFile in resultsFiles)
    {
        var resultsFileName = Path.GetFileName(resultsFile);

        var comparisonResult = await resultsFile.GetResultsFromFile() ?? new();
        switch (comparisonResult.TopicsFile)
        {
            case "commonAndCyberDeclarations.json":
                {
                    var commonTopicsObject = JsonConvert.DeserializeObject<dynamic>(Constants.CommonDeclarations ?? string.Empty);
                    comparisonResult.Topics = JsonConvert.DeserializeObject<List<ComparisonTopic>>(JsonConvert.SerializeObject(commonTopicsObject?.topics));
                    var cyberTopicsObject = JsonConvert.DeserializeObject<dynamic>(Constants.CyberDeclarations ?? string.Empty);
                    comparisonResult.Topics.AddRange(JsonConvert.DeserializeObject<List<ComparisonTopic>>(JsonConvert.SerializeObject(cyberTopicsObject?.topics)));
                }
                break;
            case "commonDeclarations.json":
                {
                    var commonTopicsObject = JsonConvert.DeserializeObject<dynamic>(Constants.CommonDeclarations ?? string.Empty);
                    comparisonResult.Topics = JsonConvert.DeserializeObject<List<ComparisonTopic>>(JsonConvert.SerializeObject(commonTopicsObject?.topics));
                }
                break;
            case "cyberDeclarations.json":
                {
                    var cyberTopicsObject = JsonConvert.DeserializeObject<dynamic>(Constants.CyberDeclarations ?? string.Empty);
                    comparisonResult.Topics = JsonConvert.DeserializeObject<List<ComparisonTopic>>(JsonConvert.SerializeObject(cyberTopicsObject?.topics));
                }
                break;
        }


        var resultsKeyPattern = fileKeyPattern.Replace("Binder", "Results").Replace("[pdf|PDF|Pdf]", "json");
        if (!string.IsNullOrWhiteSpace(resultsKeyPattern) && !string.IsNullOrWhiteSpace(resultsFileName))
        {
            var isValidMatch = Regex.Match(resultsFileName, resultsKeyPattern, RegexOptions.Singleline);
            if (isValidMatch.Success)
            {
                comparisonResult.Key = isValidMatch.Groups["key"].Value;
                comparisonResult.KeyName = fileKeyName ?? "Id:";
            }
        }

        // Get the Binder File Name and Extracted Text
        var binderFile = resultsFile.Replace("Results.json", "Binder.pdf");
        var binderDocument = comparisonResult.Documents?.FirstOrDefault(doc => doc.DocumentType == "binder") ??
                                new ClaudeDocument { FileName = binderFile, DocumentType = "binder", Topics = new() };
        binderDocument.FileName ??= binderFile;
        var binderExtractedTextFile = binderFile.Replace(".pdf", " - Extracted Text.txt");
        if (File.Exists(binderExtractedTextFile))
            binderDocument.ExtractedText = await File.ReadAllTextAsync(binderExtractedTextFile);

        // Get the Policy File Name and Extracted Text
        var policyFile = resultsFile.Replace("Results.json", "Policy.pdf");
        var policyDocument = comparisonResult.Documents?.FirstOrDefault(doc => doc.DocumentType == "policy") ??
                                new ClaudeDocument { FileName = policyFile, DocumentType = "policy", Topics = new() };
        policyDocument.FileName ??= policyFile;
        var policyExtractedTextFile = policyFile.Replace(".pdf", " - Extracted Text.txt");
        if (File.Exists(policyExtractedTextFile))
            policyDocument.ExtractedText = await File.ReadAllTextAsync(policyExtractedTextFile);

        #region "Existing Checklist"
        var hasChecklistMatch = Regex.Match(resultsFileName, resultsWithChecklistPattern, RegexOptions.Singleline);
        if (hasChecklistMatch.Success)
        {
            var submissionFileId = Convert.ToInt32(hasChecklistMatch.Groups["submissionFileId"].Value);
            var documentNumber = hasChecklistMatch.Groups["documentNumber"].Value;
            var checkListFileName = $"{bindersDirectory}{submissionFileId} - Checklist - {documentNumber}.xlsx";
            if (File.Exists(checkListFileName))
            {
                var checkListDocument = GetMarkedDiscrpanciesFromExcelFile(checkListFileName);
                if (checkListDocument != null)
                {
                    checkListDocument.SubmissionFileId = submissionFileId;
                    checkListDocument.DocumentNumber = documentNumber;
                    comparisonResult.Documents?.Add(checkListDocument);
                }
            }
        }
        #endregion

        environmentResults.Add(comparisonResult);
    }

    // Compile totals and accuracy of the document comparisons
    var environmentStatistics = Configuration?.Analyze(environmentResults);
    if (environmentStatistics == null) return;

    environmentStatistics.FileName = $"{bindersDirectory}{clientName} - Statistics.xlsx";
    await Configuration?.ToExcelFile(environmentStatistics, templateFile);

    Console.Clear();
    PressReturn();
}
ClaudeDocument? GetMarkedDiscrpanciesFromExcelFile(string? excelFile, string? documentType = null)
{
    if (string.IsNullOrWhiteSpace(excelFile) || !File.Exists(excelFile))
        return null;

    documentType ??= "checklist";
    var checkListDocument = new ClaudeDocument { FileName = excelFile, DocumentType = documentType, Topics = new() };

    // Open the Excel file
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    using var package = new ExcelPackage(new FileInfo(excelFile));

    var worksheet = package.Workbook.Worksheets[0];
    checkListDocument.NamedInsured = worksheet.Cells["E3"].Value.ToSafeString();
    checkListDocument.Term = worksheet.Cells["E4"].Value.ToSafeString();
    checkListDocument.Lob = worksheet.Cells["E5"].Value.ToSafeString();
    checkListDocument.PolicyNumber = worksheet.Cells["E6"].Value.ToSafeString();

    var commonTopics = new Dictionary<int, string>
    {
        { 15, "namedInsured" },
        {16, "mailingAddress" },
        {17, "policyNumber" },
        {18, "term" },
        {19, "entityType" },
        {20, "market" },
        {21, "locationSchedule" },
        {22, "premium" },
        {23, "mep" },
        {24, "commission" },
        {25, "terrorism" },
        {26, "claims" }
    };

    var discrepancies = new List<Discrepancy>();
    var notApplicableTopics = new List<ComparisonTopic>();
    foreach (var (rowIndex, topicKey) in commonTopics)
    {
        if (!string.IsNullOrWhiteSpace(worksheet.Cells[$"D{rowIndex}"].Value.ToSafeString()))
            notApplicableTopics.Add(new ComparisonTopic { Key = topicKey, Type = "common" });

        if (!string.IsNullOrWhiteSpace(worksheet.Cells[$"E{rowIndex}"].Value.ToSafeString()) ||
            !string.IsNullOrWhiteSpace(worksheet.Cells[$"F{rowIndex}"].Value.ToSafeString()) ||
            !string.IsNullOrWhiteSpace(worksheet.Cells[$"G{rowIndex}"].Value.ToSafeString()))
        {
            discrepancies.Add(new Discrepancy
            {
                Binder = new ComparisonTopic { Key = topicKey, Type = "common", Result = worksheet.Cells[$"F{rowIndex}"].Value },
                Policy = new ComparisonTopic { Key = topicKey, Type = "common", Result = worksheet.Cells[$"E{rowIndex}"].Value },
                Other = new ComparisonTopic { Key = topicKey, Type = "common", Result = worksheet.Cells[$"G{rowIndex}"].Value },
            });
        }
    }

    if (discrepancies.Any())
        checkListDocument.Discrepancies = discrepancies;
    if (notApplicableTopics.Any())
        checkListDocument.NotApplicableTopics = notApplicableTopics;

    return checkListDocument;
}
async Task SaveChecklistResultsExcelFile(string? templateFile, List<ClaudeDocument> comparisonDocuments)
{
    if (string.IsNullOrWhiteSpace(templateFile) || !File.Exists(templateFile) ||
        comparisonDocuments == null || !comparisonDocuments.Any())
        return;

    // Open the Excel file
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    using var package = new ExcelPackage(new FileInfo(templateFile));

    var worksheet = package.Workbook.Worksheets[0];

    var binderDocument = comparisonDocuments.FirstOrDefault(document => document.DocumentType == "binder");
    if (binderDocument == null) return;

    var binderFile = binderDocument.FileName;
    var binderFileName = Path.GetFileName(binderFile);

    var fileKeyPattern = Configuration["keyPattern"];
    if (!string.IsNullOrWhiteSpace(fileKeyPattern) && !string.IsNullOrWhiteSpace(binderFileName))
    {
        var isValidMatch = Regex.Match(binderFileName, fileKeyPattern, RegexOptions.Singleline);
        if (isValidMatch.Success)
        {
            worksheet.Cells["E1"].Value = isValidMatch.Groups["key"].Value;
            worksheet.Cells["B1"].Value = Configuration["keyName"] ?? "Id:";
        }
    }

    var policyDocument = comparisonDocuments.FirstOrDefault(document => document.DocumentType == "policy");
    var checklistDocument = comparisonDocuments.FirstOrDefault(document => document.DocumentType == "checklist");

    worksheet.Cells["E3"].Value = checklistDocument?.NamedInsured ?? binderDocument.GetValue("namedInsured") ?? policyDocument?.GetValue("namedInsured");
    worksheet.Cells["E4"].Value = checklistDocument?.Term ?? binderDocument.GetValue("term") ?? policyDocument?.GetValue("term");
    worksheet.Cells["E5"].Value = checklistDocument?.Lob ?? "Cyber Liability";
    worksheet.Cells["E6"].Value = checklistDocument?.PolicyNumber ?? binderDocument.GetValue("policyNumber") ?? policyDocument?.GetValue("policyNumber");

    var topicRowIds = new Dictionary<string, int>
    {
        {"namedInsured", 10 },
        {"mailingAddress", 11 },
        {"policyNumber", 12 },
        {"term", 13 },
        {"entityType", 14 },
        {"market", 15 },
        {"locationSchedule", 16 },
        {"premium", 17 },
        {"mep", 18 },
        {"commission", 19 },
        {"terrorism", 20 },
        {"claims", 21 },
        {"mainCoverages", 26 },
        {"retentionsDeductibles", 27 },
        {"retroDate", 28 },
        {"litigationDate", 29 },
        {"continuityDate", 30 },
        {"additionalInterest", 31 },
        {"additionalCoverageExtensions", 32 }
    };

    var checkListTopics = checklistDocument?.NotApplicableTopics ?? new();
    foreach (var checklistTopic in checkListTopics)
        worksheet.Cells[$"D{topicRowIds[checklistTopic.Key ?? string.Empty]}"].Value = "X";

    var aiDiscrepanciesDocument = comparisonDocuments.FirstOrDefault(document => document.DocumentType == "discrepancies");
    if (aiDiscrepanciesDocument?.Discrepancies == null || !aiDiscrepanciesDocument.Discrepancies.Any()) return;

    var binderTopics = binderDocument.Topics ?? new();
    var policyTopics = policyDocument?.Topics ?? new();
    var checklistDiscrepancies = checklistDocument?.Discrepancies ?? new();
    
    var discrepancies = aiDiscrepanciesDocument.Discrepancies;
    foreach (var discrepancy in discrepancies)
    {
        if (!topicRowIds.ContainsKey(discrepancy.Key ?? string.Empty)) continue;

        var discrepancyKey = discrepancy.Key;
        try
        {
            worksheet.Cells[$"E{topicRowIds[discrepancyKey ?? string.Empty]}"].Value = discrepancy?.Policy.ToResultsString() ?? policyTopics.FirstOrDefault(policyTopic => policyTopic.Key == discrepancyKey)?.ToResultsString();
            worksheet.Cells[$"F{topicRowIds[discrepancyKey ?? string.Empty]}"].Value = discrepancy?.Binder.ToResultsString() ?? binderTopics.FirstOrDefault(binderTopic => binderTopic.Key == discrepancyKey)?.ToResultsString();
        }
        catch (Exception)
        {
            Console.WriteLine($"Error processing topic: {discrepancyKey} from results file {aiDiscrepanciesDocument.FileName}");
        }

        var checklistDiscrepancy = checklistDiscrepancies.FirstOrDefault(cd => cd.Key == discrepancyKey);
        if (checklistDiscrepancy == null) continue;

        try
        {
            worksheet.Cells[$"H{topicRowIds[discrepancyKey ?? string.Empty]}"].Value = checklistDiscrepancy.Policy.ToResultsString();
            worksheet.Cells[$"I{topicRowIds[discrepancyKey ?? string.Empty]}"].Value = checklistDiscrepancy.Binder.ToResultsString();
        }
        catch (Exception)
        {
            Console.WriteLine($"Error processing actual checklist topic: {discrepancyKey} from results file {aiDiscrepanciesDocument.FileName}");
        }
    }

    var documentValuesRowIds = new Dictionary<string, int>
    {
        {"namedInsured", 6 },
        {"mailingAddress", 7 },
        {"policyNumber", 8 },
        {"term", 9 },
        {"entityType", 10 },
        {"market", 11 },
        {"locationSchedule", 12 },
        {"premium", 13 },
        {"mep", 14 },
        {"commission", 15 },
        {"terrorism", 16 },
        {"claims", 17 },
        {"mainCoverages", 25 },
        {"retentionsDeductibles", 26 },
        {"retroDate", 27 },
        {"litigationDate", 28 },
        {"continuityDate", 29 },
        {"additionalInterest", 30 },
        {"additionalCoverageExtensions", 31 }
    };


    var binderWorksheet = package.Workbook.Worksheets[1];
    foreach (var binderTopic in binderTopics)
    {
        if (!documentValuesRowIds.ContainsKey(binderTopic.Key ?? string.Empty)) continue;

        var binderResultString = binderTopic.ToResultsString();
        if (!string.IsNullOrWhiteSpace(binderResultString))
            binderWorksheet.Cells[$"D{documentValuesRowIds[binderTopic.Key ?? string.Empty]}"].Value = binderResultString;
    }

    var policyWorksheet = package.Workbook.Worksheets[2];
    foreach (var policyTopic in policyTopics)
    {
        if (!documentValuesRowIds.ContainsKey(policyTopic.Key ?? string.Empty)) continue;

        var policyResultString = policyTopic.ToResultsString();
        if (!string.IsNullOrWhiteSpace(policyResultString))
            policyWorksheet.Cells[$"D{documentValuesRowIds[policyTopic.Key ?? string.Empty]}"].Value = policyResultString;
    }

    #region "Save Excel File"
    var excelFileName = aiDiscrepanciesDocument.FileName;
    if (string.IsNullOrWhiteSpace(excelFileName))
        return;

    var file = new FileInfo(excelFileName);
    await package.SaveAsAsync(file);
    #endregion

}
#endregion
#region "Event Handlers"
void OnDisplayMessage(object? sender, MessageEventArgs? e)
{
    var logKey = e?.LogKey ?? "general";
    var message = e?.Message;

    if (string.IsNullOrWhiteSpace(message)) return;
    processingLog[logKey] = message;
    PrintLog();
}
#endregion
