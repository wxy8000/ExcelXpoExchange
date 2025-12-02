using DevExpress.ExpressApp;
using DevExpress.ExpressApp.ApplicationBuilder;
using DevExpress.ExpressApp.Blazor;
using DevExpress.ExpressApp.SystemModule;
using DevExpress.ExpressApp.Updating;
using DevExpress.ExpressApp.Xpo;
using WxyXaf.Demo.XpoExcelDictionary.Blazor.Server.Services;

namespace WxyXaf.Demo.XpoExcelDictionary.Blazor.Server
{
    public class ExcelXpoExchangeBlazorApplication : BlazorApplication
    {
        public ExcelXpoExchangeBlazorApplication()
        {
            ApplicationName = "ExcelXpoExchange";
            // 使用ModuleInfo来存储版本信息，遵循DevExpress官方的方式
            CheckCompatibilityType = DevExpress.ExpressApp.CheckCompatibilityType.ModuleInfo;
            DatabaseVersionMismatch += ExcelXpoExchangeBlazorApplication_DatabaseVersionMismatch;
        }
        protected override void OnSetupStarted()
        {
            base.OnSetupStarted();

            // 确保数据库连接字符串被正确配置
            // 设置为UpdateDatabaseAlways以自动更新数据库结构
            DatabaseUpdateMode = DatabaseUpdateMode.UpdateDatabaseAlways;
        }
        void ExcelXpoExchangeBlazorApplication_DatabaseVersionMismatch(object sender, DatabaseVersionMismatchEventArgs e)
        {
            // ����������¶��Զ��������ݿ⣬�����׳�����
            e.Updater.Update();
            e.Handled = true;
        }
    }
}
