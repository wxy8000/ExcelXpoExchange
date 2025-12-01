using System.Configuration;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.ApplicationBuilder;
using DevExpress.ExpressApp.Design;
using DevExpress.ExpressApp.Security;
using DevExpress.ExpressApp.Win;
using DevExpress.ExpressApp.Win.ApplicationBuilder;
using DevExpress.Persistent.Base;
using DevExpress.Persistent.BaseImpl;
using DevExpress.XtraEditors;

namespace ExcelXpoExchange.Win
{
    public class ApplicationBuilder : IDesignTimeApplicationFactory
    {
        public static WinApplication BuildApplication(string connectionString)
        {
            var builder = WinApplication.CreateBuilder();
            // Register custom services for Dependency Injection. For more information, refer to the following topic: https://docs.devexpress.com/eXpressAppFramework/404430/
            // builder.Services.AddScoped<CustomService>();
            // Register 3rd-party IoC containers (like Autofac, Dryloc, etc.)
            // builder.UseServiceProviderFactory(new DryIocServiceProviderFactory());
            // builder.UseServiceProviderFactory(new AutofacServiceProviderFactory());

            builder.UseApplication<ExcelXpoExchangeWindowsFormsApplication>();
            builder.Modules
                .AddConditionalAppearance()
                .AddValidation(options =>
                {
                    options.AllowValidationDetailsAccess = false;
                })
                .Add<ExcelXpoExchange.Module.ExcelXpoExchangeModule>()
                .Add<WxyXpoExcel.WxyXpoExcelModule>()
                .Add<ExcelXpoExchangeWinModule>();
            builder.ObjectSpaceProviders
                .AddXpo((application, options) =>
                {
                    options.ConnectionString = connectionString;
                })
                .AddNonPersistent();
            builder.AddBuildStep(application =>
            {
                application.ConnectionString = connectionString;
                
                // 确保在所有情况下都能正确更新数据库
                // 设置为UpdateDatabaseAlways以自动更新数据库架构
                application.DatabaseUpdateMode = DatabaseUpdateMode.UpdateDatabaseAlways;
                
                // 推荐使用ModuleInfo来检查兼容性，这是DevExpress推荐的方式
                application.CheckCompatibilityType = CheckCompatibilityType.ModuleInfo;
            });
            var winApplication = builder.Build();
            return winApplication;
        }

        XafApplication IDesignTimeApplicationFactory.Create()
            => BuildApplication(XafApplication.DesignTimeConnectionString);
    }
}
