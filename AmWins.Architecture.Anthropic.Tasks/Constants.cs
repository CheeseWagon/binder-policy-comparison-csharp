using System.Reflection;

namespace AmWins.Architecture.Anthropic.Tasks;

internal class Constants
{
    #region "Constants"
    public const string HUMAN_PROMPT = @"#n#nHuman: ";
    public const string ASSISTANT_PROMPT = @"#n#nAssistant:";
    public const string LINE_FEED_TEMPLATE = @"#n#n";
    public const string TAB_TEMPLATE = @"#t";
    public const string DOUBLE_QUOTE_TEMPLATE = @"#q";

    public const decimal PROMPT_RATE = 11.02M;
    public const decimal RESPONSE_RATE = 32.68M;
    #endregion
    #region "Read-Only Properties"
    public static string? CyberDeclarations
    {
        get
        {
            var assembly = Assembly.GetExecutingAssembly();
            var stream =
                assembly.GetManifestResourceStream(
                    "AmWins.Architecture.Anthropic.Tasks.Resources.cyberDeclarations.json");
            if (stream == null)
                return null;


            using var sr = new StreamReader(stream);
            return sr.ReadToEnd();
        }
    }
    public static string? CommonDeclarations
    {
        get
        {
            var assembly = Assembly.GetExecutingAssembly();
            var stream =
                assembly.GetManifestResourceStream(
                    "AmWins.Architecture.Anthropic.Tasks.Resources.commonDeclarations.json");
            if (stream == null)
                return null;


            using var sr = new StreamReader(stream);
            return sr.ReadToEnd();
        }
    }
    #endregion
}
