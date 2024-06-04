namespace AmWins.Architecture.Anthropic.Tasks.Models;

public class KeyedItem
{
    #region "Member Variables"
    private string? _key = string.Empty;
    #endregion
    #region "Properties"
    public virtual string? Key
    {
        get
        {
            if (string.IsNullOrWhiteSpace(_key))
                _key = UniqueId();
            return _key;
        }
        set => _key = value;
    }
    #endregion
    #region "Implementation"
    public static string UniqueId()
    {
        return Guid.NewGuid().ToString().ToUpper();
    }
    #endregion
}