using System.ComponentModel;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.ApplicationBuilder;
using DevExpress.ExpressApp.Updating;
using DevExpress.ExpressApp.Win;
using DevExpress.ExpressApp.Win.Utils;
using DevExpress.ExpressApp.Xpo;

namespace ExcelXpoExchange.Win
{
    // For more typical usage scenarios, be sure to check out https://docs.devexpress.com/eXpressAppFramework/DevExpress.ExpressApp.Win.WinApplication._members
    public class ExcelXpoExchangeWindowsFormsApplication : WinApplication
    {
        public ExcelXpoExchangeWindowsFormsApplication()
        {
            SplashScreen = new DXSplashScreen(typeof(XafSplashScreen), new DefaultOverlayFormOptions());
            ApplicationName = "ExcelXpoExchange";
            // 使用ModuleInfo来检查兼容性，这是DevExpress推荐的方式
            CheckCompatibilityType = DevExpress.ExpressApp.CheckCompatibilityType.ModuleInfo;
            UseOldTemplates = false;
            DatabaseVersionMismatch += ExcelXpoExchangeWindowsFormsApplication_DatabaseVersionMismatch;
            CustomizeLanguagesList += ExcelXpoExchangeWindowsFormsApplication_CustomizeLanguagesList;
        }
        void ExcelXpoExchangeWindowsFormsApplication_CustomizeLanguagesList(object sender, CustomizeLanguagesListEventArgs e)
        {
            string userLanguageName = System.Threading.Thread.CurrentThread.CurrentUICulture.Name;
            if (userLanguageName != "en-US" && e.Languages.IndexOf(userLanguageName) == -1)
            {
                e.Languages.Add(userLanguageName);
            }
        }
        void ExcelXpoExchangeWindowsFormsApplication_DatabaseVersionMismatch(object sender, DevExpress.ExpressApp.DatabaseVersionMismatchEventArgs e)
        {
            // 在所有情况下都自动更新数据库，避免抛出错误
            e.Updater.Update();
            e.Handled = true;
        }
    }
}
