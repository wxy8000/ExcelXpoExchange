using DevExpress.ExpressApp.ApplicationBuilder;
using DevExpress.ExpressApp.Blazor.ApplicationBuilder;
using DevExpress.ExpressApp.Blazor.Services;
using WxyXaf.XpoExcel;
using WxyXaf.DataDictionaries;
using WxyXaf.Demo.XpoExcelDictionary.Blazor.Server.Services;
using Microsoft.AspNetCore.Components.Server.Circuits;

namespace WxyXaf.Demo.XpoExcelDictionary.Blazor.Server
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddSingleton(typeof(Microsoft.AspNetCore.SignalR.HubConnectionHandler<>), typeof(ProxyHubConnectionHandler<>));

            services.AddRazorPages();
            services.AddServerSideBlazor();
            services.AddHttpContextAccessor();
            services.AddScoped<CircuitHandler, CircuitHandlerProxy>();
            services.AddXaf(Configuration, builder =>
            {
                builder.UseApplication<WxyXaf.Demo.XpoExcelDictionary.Blazor.Server.ExcelXpoExchangeBlazorApplication>();
                builder.Modules
                    .AddConditionalAppearance()
                    .AddValidation(options =>
                    {
                        options.AllowValidationDetailsAccess = false;
                    })
                    .Add<WxyXaf.Demo.XpoExcelDictionary.Module.ExcelXpoExchangeModule>()
                    .Add<DataDictionariesModule>()
                    .Add<WxyXaf.Demo.XpoExcelDictionary.Blazor.Server.ExcelXpoExchangeBlazorModule>();
                builder.ObjectSpaceProviders
                    .AddXpo((serviceProvider, options) =>
                    {
                        string connectionString = null;
                        if (Configuration.GetConnectionString("ConnectionString") != null)
                        {
                            connectionString = Configuration.GetConnectionString("ConnectionString");
                        }
#if EASYTEST
                        if(Configuration.GetConnectionString("EasyTestConnectionString") != null) {
                            connectionString = Configuration.GetConnectionString("EasyTestConnectionString");
                        }
#endif
                        ArgumentNullException.ThrowIfNull(connectionString);
                        
                        // ������ݿ�����ԣ������������ɾ�������ݿ�
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
                                if (System.IO.File.Exists(dbFilePath))
                                {
                                    System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ������ݿ������: {dbFilePath}");
                                    
                                    // ���Դ����ݿⲢ��������
                                    try
                                    {
                                        using (var sqliteConnection = new System.Data.SQLite.SQLiteConnection($"Data Source={dbFilePath}"))
                                        {
                                            sqliteConnection.Open();
                                            
                                            // ������ݿ�汾���Ƿ����
                                            using (var command = new System.Data.SQLite.SQLiteCommand(
                                                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='XPObjectType'", 
                                                sqliteConnection))
                                            {
                                                var count = Convert.ToInt32(command.ExecuteScalar());
                                                if (count == 0)
                                                {
                                                    // ���ݿ�ṹ�����ݣ�ɾ�������ݿ�
                                                    System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ���ݿ�ṹ�����ݣ�ɾ�������ݿ�: {dbFilePath}");
                                                    sqliteConnection.Close();
                                                    System.IO.File.Delete(dbFilePath);
                                                    System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] �����ݿ���ɾ��");
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        // ����ʧ�ܣ����ݿ�����𻵻򲻼��ݣ�ɾ�������ݿ�
                                        System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ���ݿ�����ʧ�ܣ�ɾ�������ݿ�: {dbFilePath}, ����: {ex.Message}");
                                        System.IO.File.Delete(dbFilePath);
                                        System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] �����ݿ���ɾ��");
                                    }
                                }
                            }
                        }
                        
                        options.ConnectionString = connectionString;
                        options.ThreadSafe = true;
                        options.UseSharedDataStoreProvider = true;
                    })
                    .AddNonPersistent();
            });
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Error");
                // The default HSTS value is 30 days. To change this for production scenarios, see: https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }
            app.UseHttpsRedirection();
            app.UseRequestLocalization();
            app.UseStaticFiles();
            app.UseRouting();
            app.UseXaf();
            app.UseEndpoints(endpoints =>
            {
                endpoints.MapXafEndpoints();
                endpoints.MapBlazorHub();
                endpoints.MapFallbackToPage("/_Host");
                endpoints.MapControllers();
            });
        }
    }
}
