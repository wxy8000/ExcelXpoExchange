using DevExpress.ExpressApp;
using DevExpress.ExpressApp.ApplicationBuilder;
using DevExpress.ExpressApp.Blazor;
using DevExpress.ExpressApp.SystemModule;
using DevExpress.ExpressApp.Updating;
using DevExpress.ExpressApp.Xpo;
using ExcelXpoExchange.Blazor.Server.Services;

namespace ExcelXpoExchange.Blazor.Server
{
    public class ExcelXpoExchangeBlazorApplication : BlazorApplication
    {
        public ExcelXpoExchangeBlazorApplication()
        {
            ApplicationName = "ExcelXpoExchange";
            // 使用ModuleInfo来检查兼容性，这是DevExpress推荐的方式
            CheckCompatibilityType = DevExpress.ExpressApp.CheckCompatibilityType.ModuleInfo;
            DatabaseVersionMismatch += ExcelXpoExchangeBlazorApplication_DatabaseVersionMismatch;
        }
        protected override void OnSetupStarted()
        {
            base.OnSetupStarted();

            // 确保在所有情况下都能正确更新数据库
            // 设置为UpdateDatabaseAlways以自动更新数据库架构
            DatabaseUpdateMode = DatabaseUpdateMode.UpdateDatabaseAlways;
        }
        void ExcelXpoExchangeBlazorApplication_DatabaseVersionMismatch(object sender, DatabaseVersionMismatchEventArgs e)
        {
            // 在所有情况下都自动更新数据库，避免抛出错误
            e.Updater.Update();
            e.Handled = true;
        }
    }
}
