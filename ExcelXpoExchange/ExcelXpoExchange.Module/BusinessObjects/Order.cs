using DevExpress.Data.Filtering;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.DC;
using DevExpress.ExpressApp.Model;
using DevExpress.Persistent.Base;
using DevExpress.Persistent.BaseImpl;
using DevExpress.Persistent.Validation;
using DevExpress.Xpo;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using WxyXaf.DataDictionaries;
using WxyXpoExcel;

namespace ExcelXpoExchange.Module.BusinessObjects
{
    [DefaultClassOptions]
    [ExcelImportExport()]
    //[ImageName("BO_Contact")]
    //[DefaultProperty("DisplayMemberNameForLookupEditorsOfThisType")]
    //[DefaultListViewOptions(MasterDetailMode.ListViewOnly, false, NewItemRowPosition.None)]
    //[Persistent("DatabaseTableName")]
    // Specify more UI options using a declarative approach (https://docs.devexpress.com/eXpressAppFramework/112701/business-model-design-orm/data-annotations-in-data-model).
    public class Order : BaseObject
    { // Inherit from a different class to provide a custom primary key, concurrency and deletion behavior, etc. (https://docs.devexpress.com/eXpressAppFramework/113146/business-model-design-orm/business-model-design-with-xpo/base-persistent-classes).
        // Use CodeRush to create XPO classes and properties with a few keystrokes.
        // https://docs.devexpress.com/CodeRushForRoslyn/118557
        public Order(Session session)
            : base(session)
        {
        }
        public override void AfterConstruction()
        {
            base.AfterConstruction();
            // Place your initialization code here (https://docs.devexpress.com/eXpressAppFramework/112834/getting-started/in-depth-tutorial-winforms-webforms/business-model-design/initialize-a-property-after-creating-an-object-xpo?v=22.1).
        }
        /// <summary>
        /// 员工ID
        /// </summary>
        [Size(50)]
        [ExcelField(Caption = "员工ID", Order = 0,IsUnique =true)]
        [VisibleInDetailView(true)]
        [VisibleInListView(true)]
        [VisibleInLookupListView(true)]

        public string EmployeeId
        {
            get => fEmployeeId;
            set => SetPropertyValue(nameof(EmployeeId), ref fEmployeeId, value);
        }
        string fEmployeeId;

        /// <summary>
        /// 姓名
        /// </summary>
        [Size(100)]
        [ExcelField(Caption = "姓名", Order = 1)]
        [VisibleInDetailView(true)]
        [VisibleInListView(true)]
        [VisibleInLookupListView(true)]
        public string Name
        {
            get => fName;
            set => SetPropertyValue(nameof(Name), ref fName, value);
        }
        string fName;

        /// <summary>
        /// 年龄
        /// </summary>
        [ExcelField(Caption = "年龄", Order = 2)]
        [VisibleInDetailView(true)]
        [VisibleInListView(true)]
        public int Age
        {
            get => fAge;
            set => SetPropertyValue(nameof(Age), ref fAge, value);
        }
        int fAge;

        /// <summary>
        /// 邮箱
        /// </summary>
        [Size(200)]
        [ExcelField(Caption = "邮箱", Order = 3)]
        [VisibleInDetailView(true)]
        [VisibleInListView(true)]
        public string Email
        {
            get => fEmail;
            set => SetPropertyValue(nameof(Email), ref fEmail, value);
        }
        string fEmail;

        /// <summary>
        /// 电话
        /// </summary>
        [Size(50)]
        [ExcelField(Caption = "电话", Order = 4)]
        [VisibleInDetailView(true)]
        [VisibleInListView(true)]
        public string Phone
        {
            get => fPhone;
            set => SetPropertyValue(nameof(Phone), ref fPhone, value);
        }
        string fPhone;

        /// <summary>
        /// 部门
        /// </summary>
        [Size(100)]
        [ExcelField(Caption = "部门", Order = 5)]
        [VisibleInDetailView(true)]
        [VisibleInListView(true)]
        [DataDictionary("部门")]
        public DataDictionaryItem Department
        {
            get => fDepartment;
            set => SetPropertyValue(nameof(Department), ref fDepartment, value);
        }
        DataDictionaryItem fDepartment;


        /// <summary>
        /// 职位
        /// </summary>
        [Size(100)]
        [ExcelField(Caption = "职位", Order = 5)]
        [VisibleInDetailView(true)]
        [VisibleInListView(true)]
        [DataDictionary("职位")]
        public DataDictionaryItem ZhiWei
        {
            get => fZhiWei;
            set => SetPropertyValue(nameof(ZhiWei), ref fZhiWei, value);
        }
        DataDictionaryItem fZhiWei;

        /// <summary>
        /// 入职日期
        /// </summary>
        [ExcelField(Caption = "入职日期", Order = 6)]
        [VisibleInDetailView(true)]
        [VisibleInListView(true)]
        public DateTime HireDate
        {
            get => fHireDate;
            set => SetPropertyValue(nameof(HireDate), ref fHireDate, value);
        }
        DateTime fHireDate;

        /// <summary>
        /// 是否在职
        /// </summary>
        [ExcelField(Caption = "是否在职", Order = 7)]
        [VisibleInDetailView(true)]
        [VisibleInListView(true)]
        public bool IsActive
        {
            get => fIsActive;
            set => SetPropertyValue(nameof(IsActive), ref fIsActive, value);
        }
        bool fIsActive;

        //[Action(Caption = "My UI Action", ConfirmationMessage = "Are you sure?", ImageName = "Attention", AutoCommit = true)]
        //public void ActionMethod() {
        //    // Trigger a custom business logic for the current record in the UI (https://docs.devexpress.com/eXpressAppFramework/112619/ui-construction/controllers-and-actions/actions/how-to-create-an-action-using-the-action-attribute).
        //    this.PersistentProperty = "Paid";
        //}
    }
}