using System.Configuration;
using System.IO;
using System.Reflection;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.ApplicationBuilder;
using DevExpress.ExpressApp.Security;
using DevExpress.ExpressApp.Utils;
using DevExpress.ExpressApp.Win;
using DevExpress.ExpressApp.Win.ApplicationBuilder;
using DevExpress.ExpressApp.Win.Utils;
using DevExpress.Persistent.Base;
using DevExpress.Persistent.BaseImpl;
using DevExpress.XtraEditors;

namespace ExcelXpoExchange.Win
{
    internal static class Program
    {
        static bool ContainsArgument(string[] args, string argument)
        {
            return args.Any(arg => arg.TrimStart('/').TrimStart('-').ToLower() == argument.ToLower());
        }
        
        /// <summary>
        /// 检查数据库兼容性，如果不兼容则删除旧数据库
        /// </summary>
        /// <param name="connectionString">数据库连接字符串</param>
        static void CheckAndDeleteIncompatibleDatabase(string connectionString)
        {
            try
            {
                // 提取SQLite数据库文件路径
                if (connectionString.Contains("XpoProvider=SQLite"))
                {
                    // 解析连接字符串，获取Data Source参数
                    var dataSourceParam = connectionString.Split(';')
                        .FirstOrDefault(p => p.Trim().StartsWith("Data Source", StringComparison.OrdinalIgnoreCase));
                    
                    if (dataSourceParam != null)
                    {
                        var dbFilePath = dataSourceParam.Split('=')[1].Trim();
                        
                        // 检查数据库文件是否存在
                        if (File.Exists(dbFilePath))
                        {
                            System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 检查数据库兼容性: {dbFilePath}");
                            
                            // 尝试打开数据库并检查兼容性
                            // 这里我们简化处理，直接尝试连接，如果失败则删除旧数据库
                            try
                            {
                                using (var connection = new System.Data.SQLite.SQLiteConnection($"Data Source={dbFilePath}"))
                                {
                                    connection.Open();
                                    
                                    // 检查数据库版本表是否存在
                                    using (var command = new System.Data.SQLite.SQLiteCommand(
                                        "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='XPObjectType'", 
                                        connection))
                                    {
                                        var count = Convert.ToInt32(command.ExecuteScalar());
                                        if (count == 0)
                                        {
                                            // 数据库结构不兼容，删除旧数据库
                                            System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 数据库结构不兼容，删除旧数据库: {dbFilePath}");
                                            connection.Close();
                                            File.Delete(dbFilePath);
                                            System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 旧数据库已删除");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                // 连接失败，数据库可能损坏或不兼容，删除旧数据库
                                System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 数据库连接失败，删除旧数据库: {dbFilePath}, 错误: {ex.Message}");
                                File.Delete(dbFilePath);
                                System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 旧数据库已删除");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 检查数据库兼容性时发生错误: {ex.Message}");
                // 忽略错误，继续启动应用程序
            }
        }
        
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static int Main(string[] args)
        {
            if (ContainsArgument(args, "help") || ContainsArgument(args, "h"))
            {
                Console.WriteLine("Updates the database when its version does not match the application's version.");
                Console.WriteLine();
                Console.WriteLine($"    {Assembly.GetExecutingAssembly().GetName().Name}.exe --updateDatabase [--forceUpdate --silent]");
                Console.WriteLine();
                Console.WriteLine("--forceUpdate - Marks that the database must be updated whether its version matches the application's version or not.");
                Console.WriteLine("--silent - Marks that database update proceeds automatically and does not require any interaction with the user.");
                Console.WriteLine();
                Console.WriteLine($"Exit codes: 0 - {DBUpdaterStatus.UpdateCompleted}");
                Console.WriteLine($"            1 - {DBUpdaterStatus.UpdateError}");
                Console.WriteLine($"            2 - {DBUpdaterStatus.UpdateNotNeeded}");
                return 0;
            }
            DevExpress.ExpressApp.FrameworkSettings.DefaultSettingsCompatibilityMode = DevExpress.ExpressApp.FrameworkSettingsCompatibilityMode.Latest;
#if EASYTEST
            DevExpress.ExpressApp.Win.EasyTest.EasyTestRemotingRegistration.Register();
#endif
            WindowsFormsSettings.LoadApplicationSettings();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            DevExpress.Utils.ToolTipController.DefaultController.ToolTipType = DevExpress.Utils.ToolTipType.SuperTip;
            if (Tracing.GetFileLocationFromSettings() == DevExpress.Persistent.Base.FileLocation.CurrentUserApplicationDataFolder)
            {
                Tracing.LocalUserAppDataPath = Application.LocalUserAppDataPath;
            }
            Tracing.Initialize();

            string connectionString = null;
            if (ConfigurationManager.ConnectionStrings["ConnectionString"] != null)
            {
                connectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
            }
#if EASYTEST
            if(ConfigurationManager.ConnectionStrings["EasyTestConnectionString"] != null) {
                connectionString = ConfigurationManager.ConnectionStrings["EasyTestConnectionString"].ConnectionString;
            }
#endif
            ArgumentNullException.ThrowIfNull(connectionString);
            
            // 检查数据库兼容性，如果不兼容则删除旧数据库
            CheckAndDeleteIncompatibleDatabase(connectionString);
            
            var winApplication = ApplicationBuilder.BuildApplication(connectionString);

            if (ContainsArgument(args, "updateDatabase"))
            {
                using var dbUpdater = new WinDBUpdater(() => winApplication);
                return dbUpdater.Update(
                    forceUpdate: ContainsArgument(args, "forceUpdate"),
                    silent: ContainsArgument(args, "silent"));
            }

            try
            {
                winApplication.Setup();
                winApplication.Start();
            }
            catch (Exception e)
            {
                winApplication.StopSplash();
                winApplication.HandleException(e);
            }
            return 0;
        }
    }
}
