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
using WxyXaf.XpoExcel;

namespace WxyXaf.Demo.XpoExcelDictionary.Blazor.Server.Controllers
{
    /// <summary>
    /// Blazor��ͨ��Excel���뵼��������
    /// </summary>
    public class BlazorExcelImportExportViewController : ExcelImportExportViewController
    {
        private NavigationManager navigationManager;
        private IJSRuntime jsRuntime;

        /// <summary>
        /// �ڿ�����������ʱ���ã���ʱApplication�����Ѿ�����ʼ��
        /// </summary>
        protected override void OnActivated()
        {
            base.OnActivated();

            // �ӷ����ṩ�����ȡ�������
            if (Application != null && Application.ServiceProvider != null)
            {
                navigationManager = Application.ServiceProvider.GetService<NavigationManager>();
                jsRuntime = Application.ServiceProvider.GetService<IJSRuntime>();
            }
        }

        /// <summary>
        /// ִ�е��������ʵ��Blazorƽ̨��Excel���빦��
        /// </summary>
        /// <param name="e">�¼�����</param>
        protected override void ExecuteImportAction(SimpleActionExecuteEventArgs e)
        {
            try
            {
                // ʹ�õ�����ר�ŵĵ���ҳ�棬���ݵ�ǰ�������͵�������Ϊ����
                if (navigationManager != null)
                {
                    var objectTypeName = ObjectType.FullName;
                    navigationManager.NavigateTo($"/import-excel-dialog?objectType={Uri.EscapeDataString(objectTypeName)}", forceLoad: true);
                }
                else
                {
                    Application.ShowViewStrategy.ShowMessage(
                        "Blazor平台导入功能需要使用专门的页面\n请访问以下地址进行导入：/import-excel-dialog",
                        InformationType.Info
                    );
                }
            }
            catch (Exception ex)
            {
                Application.ShowViewStrategy.ShowMessage($"导入Excel失败：{ex.Message}", InformationType.Error);
            }
        }

        /// <summary>
        /// 导出按钮点击事件，实现Blazor平台的Excel导出功能
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        protected override void ExportAction_Execute(object sender, SimpleActionExecuteEventArgs e)
        {
            try
            {
                // 使用XpoExcelHelper执行导出
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

                // 使用JavaScript下载Excel文件
                if (jsRuntime != null)
                {
                    var fileName = $"{ObjectType.Name}_导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                    // 直接传递中文文件名，现代浏览器支持UTF-8编码的文件名
                    jsRuntime.InvokeVoidAsync("downloadFile", fileName, bytes);

                    // 显示成功信息
                    Application.ShowViewStrategy.ShowMessage(
                        "数据已成功导出",
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
                // 显示错误信息
                Application.ShowViewStrategy.ShowMessage(
                    $"导出失败：{ex.Message}",
                    InformationType.Error
                );
            }
        }
    }
}
