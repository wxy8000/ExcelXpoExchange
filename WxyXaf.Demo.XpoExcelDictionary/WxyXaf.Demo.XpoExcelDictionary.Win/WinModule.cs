using System.ComponentModel;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using DevExpress.ExpressApp.DC;
using DevExpress.ExpressApp.Editors;
using DevExpress.ExpressApp.Model;
using DevExpress.ExpressApp.Model.Core;
using DevExpress.ExpressApp.Model.DomainLogics;
using DevExpress.ExpressApp.Model.NodeGenerators;
using DevExpress.ExpressApp.Updating;

namespace WxyXaf.Demo.XpoExcelDictionary.Win
{
    [ToolboxItemFilter("Xaf.Platform.Win")]
    // For more typical usage scenarios, be sure to check out https://docs.devexpress.com/eXpressAppFramework/DevExpress.ExpressApp.ModuleBase.
    public sealed class ExcelXpoExchangeWinModule : ModuleBase
    {
        public ExcelXpoExchangeWinModule()
        {
            DevExpress.ExpressApp.Editors.FormattingProvider.UseMaskSettings = true;
            // ��ʽע��WinFormsƽ̨�ض��Ŀ�����
            AdditionalControllerTypes.Add(typeof(WxyXaf.Demo.XpoExcelDictionary.Win.Controllers.WinExcelImportExportViewController));
            
            // 注册WinForms特定的TreeListEditors模块
            RequiredModuleTypes.Add(typeof(DevExpress.ExpressApp.TreeListEditors.Win.TreeListEditorsWindowsFormsModule));
        }
        public override IEnumerable<ModuleUpdater> GetModuleUpdaters(IObjectSpace objectSpace, Version versionFromDB)
        {
            return ModuleUpdater.EmptyModuleUpdaters;
        }
        public override void Setup(XafApplication application)
        {
            base.Setup(application);
        }
    }
}
