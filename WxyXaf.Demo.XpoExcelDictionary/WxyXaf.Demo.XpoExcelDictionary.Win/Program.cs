using System.Configuration;
using System.Data.SQLite;
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

namespace WxyXaf.Demo.XpoExcelDictionary.Win
{
    internal static class Program
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
                        
                        // ������ݿ��ļ��Ƿ����
                        if (File.Exists(dbFilePath))
                        {
                            System.Diagnostics.Debug.WriteLine($"������ݿ������: {dbFilePath}");
                            
                            // ���Դ����ݿⲢ��������
                            // �������Ǽ򻯴�����ֱ�ӳ������ӣ����ʧ����ɾ�������ݿ�
                            try
                            {
                                using (var connection = new System.Data.SQLite.SQLiteConnection($"Data Source={dbFilePath}"))
                                {
                                    connection.Open();
                                    
                                    // ������ݿ�汾���Ƿ����
                                    using (var command = new System.Data.SQLite.SQLiteCommand(
                                        "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='XPObjectType'", 
                                        connection))
                                    {
                                        var count = Convert.ToInt32(command.ExecuteScalar());
                                        if (count == 0)
                                        {
                                            // ���ݿ�ṹ�����ݣ�ɾ�������ݿ�
                                            System.Diagnostics.Debug.WriteLine($"���ݿ�ṹ�����ݣ�ɾ�������ݿ�: {dbFilePath}");
                                            connection.Close();
                                            File.Delete(dbFilePath);
                                            System.Diagnostics.Debug.WriteLine($"�����ݿ���ɾ��");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                // ����ʧ�ܣ����ݿ�����𻵻򲻼��ݣ�ɾ�������ݿ�
                                System.Diagnostics.Debug.WriteLine($"���ݿ�����ʧ�ܣ�ɾ�������ݿ�: {dbFilePath}, ����: {ex.Message}");
                                File.Delete(dbFilePath);
                                System.Diagnostics.Debug.WriteLine($"�����ݿ���ɾ��");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"������ݿ������ʱ��������: {ex.Message}");
                // ���Դ��󣬼�������Ӧ�ó���
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
            
            // ������ݿ�����ԣ������������ɾ�������ݿ�
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
