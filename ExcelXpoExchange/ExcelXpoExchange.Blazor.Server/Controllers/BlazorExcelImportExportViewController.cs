using System;
using System.IO;
using System.Web;
using DevExpress.Data.Filtering;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using DevExpress.Persistent.Base;
using Microsoft.AspNetCore.Components;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.JSInterop;
using WxyXpoExcel;

namespace YourBlazorProject.Controllers
{
    /// <summary>
    /// Blazor版通用Excel导入导出控制器
    /// </summary>
    public class BlazorExcelImportExportViewController : ExcelImportExportViewController
    {
        private NavigationManager navigationManager;
        private IJSRuntime jsRuntime;

        /// <summary>
        /// 在控制器被激活时调用，此时Application属性已经被初始化
        /// </summary>
        protected override void OnActivated()
        {
            base.OnActivated();

            // 从服务提供程序获取所需服务
            if (Application != null && Application.ServiceProvider != null)
            {
                navigationManager = Application.ServiceProvider.GetService<NavigationManager>();
                jsRuntime = Application.ServiceProvider.GetService<IJSRuntime>();
            }
        }

        /// <summary>
        /// 执行导入操作，实现Blazor平台的Excel导入功能
        /// </summary>
        /// <param name="e">事件参数</param>
        protected override void ExecuteImportAction(SimpleActionExecuteEventArgs e)
        {
            try
            {
                // 使用导航到专门的导入页面，传递当前对象类型的名称作为参数
                if (navigationManager != null)
                {
                    var objectTypeName = ObjectType.FullName;
                    navigationManager.NavigateTo($"/import-excel-dialog?objectType={Uri.EscapeDataString(objectTypeName)}", forceLoad: true);
                }
                else
                {
                    Application.ShowViewStrategy.ShowMessage(
                        "Blazor版导入功能需要访问专门的页面，\n请在浏览器地址栏中输入：/import-excel-dialog",
                        InformationType.Info
                    );
                }
            }
            catch (Exception ex)
            {
                Application.ShowViewStrategy.ShowMessage($"导入Excel失败，{ex.Message}", InformationType.Error);
            }
        }

        /// <summary>
        /// 重写导出按钮点击事件，实现Blazor平台的Excel导出功能
        /// </summary>
        /// <param name="sender">事件发送者</param>
        /// <param name="e">事件参数</param>
        protected override void ExportAction_Execute(object sender, SimpleActionExecuteEventArgs e)
        {
            try
            {
                // 使用XpoExcelHelper导出数据
                var excelHelper = new XpoExcelHelper(Application, null);
                var exportOptions = ExcelImportExportAttribute?.ExportOptions ?? new XpoExcelExportOptions();

                // 导出数据到内存流
                var exportMethod = typeof(XpoExcelHelper).GetMethod("ExportToExcelStream", new[] { typeof(CriteriaOperator), typeof(XpoExcelExportOptions) });
                if (exportMethod == null)
                {
                    throw new InvalidOperationException("无法找到ExportToExcelStream方法");
                }

                var genericExportMethod = exportMethod.MakeGenericMethod(ObjectType);
                var stream = (MemoryStream)genericExportMethod.Invoke(excelHelper, new object[] { null, exportOptions });

                // 将内存流转换为字节数组
                stream.Position = 0;
                var bytes = stream.ToArray();

                // 使用JavaScript互操作下载文件
                if (jsRuntime != null)
                {
                    var fileName = $"{ObjectType.Name}_导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                    jsRuntime.InvokeVoidAsync("downloadFile", fileName, bytes);

                    // 显示成功消息
                    Application.ShowViewStrategy.ShowMessage(
                        $"数据已成功导出并下载",
                        InformationType.Success
                    );
                }
                else
                {
                    Application.ShowViewStrategy.ShowMessage(
                        "无法获取IJSRuntime服务，导出失败",
                        InformationType.Error
                    );
                }
            }
            catch (Exception ex)
            {
                // 显示错误消息
                Application.ShowViewStrategy.ShowMessage(
                    $"导出失败：{ex.Message}",
                    InformationType.Error
                );
            }
        }
    }
}
