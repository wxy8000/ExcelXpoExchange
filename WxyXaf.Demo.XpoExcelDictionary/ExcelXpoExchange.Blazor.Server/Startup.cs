using DevExpress.ExpressApp.ApplicationBuilder;
using DevExpress.ExpressApp.Blazor.ApplicationBuilder;
using DevExpress.ExpressApp.Blazor.Services;
using WxyXaf.DataDictionaries;
using ExcelXpoExchange.Blazor.Server.Services;
using Microsoft.AspNetCore.Components.Server.Circuits;

namespace ExcelXpoExchange.Blazor.Server
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
                builder.UseApplication<ExcelXpoExchangeBlazorApplication>();
                builder.Modules
                    .AddConditionalAppearance()
                    .AddValidation(options =>
                    {
                        options.AllowValidationDetailsAccess = false;
                    })
                    .Add<ExcelXpoExchange.Module.ExcelXpoExchangeModule>()
                    .Add<DataDictionariesModule>()
                    .Add<ExcelXpoExchangeBlazorModule>();
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
                        
                        // 检查数据库兼容性，如果不兼容则删除旧数据库
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
                                if (System.IO.File.Exists(dbFilePath))
                                {
                                    System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 检查数据库兼容性: {dbFilePath}");
                                    
                                    // 尝试打开数据库并检查兼容性
                                    try
                                    {
                                        using (var sqliteConnection = new System.Data.SQLite.SQLiteConnection($"Data Source={dbFilePath}"))
                                        {
                                            sqliteConnection.Open();
                                            
                                            // 检查数据库版本表是否存在
                                            using (var command = new System.Data.SQLite.SQLiteCommand(
                                                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='XPObjectType'", 
                                                sqliteConnection))
                                            {
                                                var count = Convert.ToInt32(command.ExecuteScalar());
                                                if (count == 0)
                                                {
                                                    // 数据库结构不兼容，删除旧数据库
                                                    System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 数据库结构不兼容，删除旧数据库: {dbFilePath}");
                                                    sqliteConnection.Close();
                                                    System.IO.File.Delete(dbFilePath);
                                                    System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 旧数据库已删除");
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        // 连接失败，数据库可能损坏或不兼容，删除旧数据库
                                        System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 数据库连接失败，删除旧数据库: {dbFilePath}, 错误: {ex.Message}");
                                        System.IO.File.Delete(dbFilePath);
                                        System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 旧数据库已删除");
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
