namespace WxyXaf.DataDictionaries;

[AttributeUsage(AttributeTargets.Property)]
public class DataDictionaryAttribute : Attribute
{
    public string DataDictionaryName { get; }

    public DataDictionaryAttribute(string dataDictionaryName)
    {
        DataDictionaryName = dataDictionaryName;
    }
}
