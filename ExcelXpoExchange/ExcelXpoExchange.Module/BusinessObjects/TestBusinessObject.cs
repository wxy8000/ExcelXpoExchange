using DevExpress.Persistent.Base;
using DevExpress.Persistent.BaseImpl;
using DevExpress.Xpo;
using WxyXaf.DataDictionaries;
using WxyXpoExcel;

namespace ExcelXpoExchange.Module.BusinessObjects
{
    [DefaultClassOptions]
    [ExcelImportExport]
    public class TestBusinessObject : BaseObject
    {
        public TestBusinessObject(Session session) : base(session) { }

        [Size(100)]
        [ExcelField(Caption = "名称", IsUnique = true)]
        public string Name
        {
            get => GetPropertyValue<string>(nameof(Name));
            set => SetPropertyValue(nameof(Name), value);
        }

        [DataDictionary("部门")]
        [ExcelField(Caption = "部门")]
        public DataDictionaryItem Department
        {
            get => GetPropertyValue<DataDictionaryItem>(nameof(Department));
            set => SetPropertyValue(nameof(Department), value);
        }

        [DataDictionary("职位")]
        [ExcelField(Caption = "职位")]
        public DataDictionaryItem Position
        {
            get => GetPropertyValue<DataDictionaryItem>(nameof(Position));
            set => SetPropertyValue(nameof(Position), value);
        }

        [ExcelField(Caption = "年龄")]
        public int Age
        {
            get => GetPropertyValue<int>(nameof(Age));
            set => SetPropertyValue(nameof(Age), value);
        }
    }
}