using DevExpress.Persistent.Base;
using DevExpress.Persistent.BaseImpl;
using DevExpress.Persistent.Validation;
using DevExpress.Xpo;
using System.ComponentModel;
using System.Diagnostics;
using WxyXpoExcel;

namespace WxyXaf.DataDictionaries;

[Persistent]
[System.ComponentModel.DisplayName("字典项")]
[ExcelImportExport]
public class DataDictionaryItem : BaseObject
{
    private string _name;
    private string _code;
    private string _description;
    private int _order;
    private DataDictionary _dataDictionary;

    [RuleRequiredField]
    [System.ComponentModel.DisplayName("名称")]
    public string Name
    {
        get => _name;
        set => SetPropertyValue(nameof(Name), ref _name, value);
    }

    [System.ComponentModel.DisplayName("编码")]
    public string Code
    {
        get => _code;
        set => SetPropertyValue(nameof(Code), ref _code, value);
    }

    [System.ComponentModel.DisplayName("描述")]
    public string Description
    {
        get => _description;
        set => SetPropertyValue(nameof(Description), ref _description, value);
    }

    [System.ComponentModel.DisplayName("顺序")]
    public int Order
    {
        get => _order;
        set => SetPropertyValue(nameof(Order), ref _order, value);
    }

    [Association("DataDictionary-Items")]
    public DataDictionary DataDictionary
    {
        get => _dataDictionary;
        set => SetPropertyValue(nameof(DataDictionary), ref _dataDictionary, value);
    }

    [Browsable(false)]
    [RuleFromBoolProperty("字典项名称必须唯一", DefaultContexts.Save, CustomMessageTemplate = "字典项名称已存在")]
    public bool IsNameUnique
    {
        get => DataDictionary == null || !DataDictionary.Items.Any(item => item.Name == Name && item != this);
    }

    public DataDictionaryItem(Session session)
        : base(session)
    {
    }

    protected override void OnSaving()
    {
        Tracing.Tracer.LogText("=== DataDictionaryItem.OnSaving 开始执行 ===");
        Tracing.Tracer.LogText($"  对象ID: {Oid}, 名称: {Name}");
        Tracing.Tracer.LogText($"  初始DataDictionary: {(DataDictionary?.Name ?? "null")}");
        Tracing.Tracer.LogText($"  初始Order: {Order}");
        
        base.OnSaving();

        // 只有在DataDictionary已经设置的情况下才处理
        if (DataDictionary != null)
        {
            Tracing.Tracer.LogText($"  DataDictionary已设置: {DataDictionary.Name}");
            
            if (Order == 0 && DataDictionary.Items.Count > 1)
            {
                var maxOrder = DataDictionary.Items.Max(item => item.Order);
                Order = maxOrder + 1;
                Tracing.Tracer.LogText($"  自动设置Order: {Order} (基于最大Order: {maxOrder})");
            }
        }
        else
        {
            Tracing.Tracer.LogText("  DataDictionary为null，跳过自动关联");
        }

        Tracing.Tracer.LogText($"  设置DataDictionary后的状态: {(DataDictionary?.Name ?? "null")}");
        
        Tracing.Tracer.LogText("=== DataDictionaryItem.OnSaving 执行完成 ===");
    }
}