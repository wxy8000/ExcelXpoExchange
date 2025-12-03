using System.Reflection;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Blazor.DesignTime;
using DevExpress.ExpressApp.Blazor.Services;
using DevExpress.ExpressApp.Design;
using DevExpress.ExpressApp.Utils;

namespace WxyXaf.Demo.XpoExcelDictionary.Blazor.Server
{
    public class Program : IDesignTimeApplicationFactory
    {
        static bool ContainsArgument(string[] args, string argument)
        {
            return args.Any(arg => arg.TrimStart('/').TrimStart('-').ToLower() == argument.ToLower());
        }
        
        /// <summary>
        /// ������ݿ�����ԣ������������ɾ�������ݿ�
        /// </summary>
        /// <param name="connectionString">���ݿ������ַ���</param>
        static void CheckAndDeleteIncompatibleDatabase(string connectionString)
        {
            try
            {
                // ��ȡSQLite���ݿ��ļ�·��
                if (connectionString.Contains("XpoProvider=SQLite"))
                {
                    // ���������ַ�������ȡData Source����
                    var dataSourceParam = connectionString.Split(';')
                        .FirstOrDefault(p => p.Trim().StartsWith("Data Source", StringComparison.OrdinalIgnoreCase));
                    
                    if (dataSourceParam != null)
                    {
                        var dbFilePath = dataSourceParam.Split('=')[1].Trim();
                        
                        Console.WriteLine($"[DB Check] 开始检查数据库文件: {dbFilePath}");
                        
                        // 检查数据库文件是否存在
                        if (File.Exists(dbFilePath))
                        {
                            Console.WriteLine($"[DB Check] 数据库文件已存在，开始检查结构完整性");
                            
                            // 尝试连接数据库并验证完整性
                            try
                            {
                                using (var connection = new System.Data.SQLite.SQLiteConnection($"Data Source={dbFilePath}"))
                                {
                                    connection.Open();
                                    
                                    // 检查数据库版本是否兼容
                                    using (var command = new System.Data.SQLite.SQLiteCommand(
                                        "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='XPObjectType'", 
                                        connection))
                                    {
                                        var count = Convert.ToInt32(command.ExecuteScalar());
                                        Console.WriteLine($"[DB Check] XPObjectType 表数量: {count}");
                                        
                                        if (count == 0)
                                        {
                                            // 数据库结构不完整，删除重建
                                            Console.WriteLine($"[DB Check] 数据库结构不完整，删除重建");
                                            connection.Close();
                                            File.Delete(dbFilePath);
                                            Console.WriteLine($"[DB Check] 旧数据库已删除");
                                        }
                                        else
                                        {
                                            Console.WriteLine($"[DB Check] 数据库结构正常");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                // 连接失败，数据库文件可能损坏，删除重建
                                Console.WriteLine($"[DB Check] 数据库连接失败，删除重建: {ex.Message}");
                                File.Delete(dbFilePath);
                                Console.WriteLine($"[DB Check] 损坏的数据库已删除");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"[DB Check] 数据库文件不存在，将自动创建");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DB Check] ������ݿ������ʱ��������: {ex.Message}");
                // ���Դ��󣬼�������Ӧ�ó���
            }
        }
        
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
            }
            else
            {
                DevExpress.ExpressApp.FrameworkSettings.DefaultSettingsCompatibilityMode = DevExpress.ExpressApp.FrameworkSettingsCompatibilityMode.Latest;
                IHost host = CreateHostBuilder(args).Build();
                if (ContainsArgument(args, "updateDatabase"))
                {
                    using (var serviceScope = host.Services.CreateScope())
                    {
                        return serviceScope.ServiceProvider.GetRequiredService<DevExpress.ExpressApp.Utils.IDBUpdater>().Update(ContainsArgument(args, "forceUpdate"), ContainsArgument(args, "silent"));
                    }
                }
                else
                {
                    host.Run();
                }
            }
            return 0;
        }
        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureWebHostDefaults(webBuilder =>
                {
                    webBuilder.UseStartup<Startup>();
                });
        XafApplication IDesignTimeApplicationFactory.Create()
        {
            IHostBuilder hostBuilder = CreateHostBuilder(Array.Empty<string>());
            return DesignTimeApplicationFactoryHelper.Create(hostBuilder);
        }
    }
}
