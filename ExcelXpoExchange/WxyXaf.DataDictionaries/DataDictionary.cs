using DevExpress.Persistent.Base;
using DevExpress.Persistent.BaseImpl;
using DevExpress.Persistent.Validation;
using DevExpress.Xpo;
using WxyXpoExcel;

namespace WxyXaf.DataDictionaries;

[Persistent]
[NavigationItem]
[System.ComponentModel.DisplayName("数据字典")]
[ExcelImportExport]
public class DataDictionary : BaseObject
{
    private string _name;

    [RuleRequiredField]
    [RuleUniqueValue]
    [System.ComponentModel.DisplayName("名称")]
    public string Name
    {
        get => _name;
        set => SetPropertyValue(nameof(Name), ref _name, value);
    }

    [System.ComponentModel.DisplayName("字典项")]
    [Association("DataDictionary-Items"), Aggregated]
    public XPCollection<DataDictionaryItem> Items
    {
        get => GetCollection<DataDictionaryItem>(nameof(Items));
    }

    public DataDictionary(Session session)
        : base(session)
    {
    }
}
