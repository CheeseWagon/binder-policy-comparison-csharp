namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class MessageEventArgs : EventArgs
{
    #region "Properties"
    public string? LogKey { get; set; }
    public string? Message { get; set; }
    #endregion
}
