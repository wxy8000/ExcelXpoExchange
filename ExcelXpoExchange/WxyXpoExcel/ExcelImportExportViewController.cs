using System;
using System.IO;
using System.Reflection;
using DevExpress.Data.Filtering;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using DevExpress.Persistent.Base;
using DevExpress.Xpo;

namespace WxyXpoExcel
{
    /// <summary>
    /// 通用Excel导入导出控制器，用于处理带有ExcelImportExportAttribute属性的XPO类
    /// </summary>
    public class ExcelImportExportViewController : ViewController<ListView>
    {
        /// <summary>
        /// 控制器激活条件
        /// </summary>
        protected override void OnViewControlsCreated()
        {
            base.OnViewControlsCreated();
            
            // 确保动作总是可用
            _exportAction.Active["HasExcelImportExportAttribute"] = true;
            _importAction.Active["HasExcelImportExportAttribute"] = true;
        }
        
        /// <summary>
        /// 控制器视图变更时
        /// </summary>
        protected override void OnViewChanged()
        {
            base.OnViewChanged();
            
            // 确保动作总是可用
            _exportAction.Active["HasExcelImportExportAttribute"] = true;
            _importAction.Active["HasExcelImportExportAttribute"] = true;
        }

        private SimpleAction _exportAction;
        private SimpleAction _importAction;
        private ExcelImportExportAttribute _excelImportExportAttribute;
        private Type _objectType;

        public ExcelImportExportViewController()
        {
            // 创建导出动作
            _exportAction = new SimpleAction(this, "ExportToExcel", PredefinedCategory.Edit)
            {
                Caption = "导出到Excel",
                ImageName = "Export",
                ToolTip = "将数据导出到Excel文件"
            };
            _exportAction.Execute += ExportAction_Execute;

            // 创建导入动作
            _importAction = new SimpleAction(this, "ImportFromExcel", PredefinedCategory.Edit)
            {
                Caption = "从Excel导入",
                ImageName = "Import",
                ToolTip = "从Excel文件导入数据"
            };
            _importAction.Execute += ImportAction_Execute;
        }

        protected override void OnActivated()
        {
            base.OnActivated();

            // 将日志输出到文件，方便调试
            string logPath = "c:\\temp\\ExcelXpoExchange.log";
            // 确保目录存在
            Directory.CreateDirectory(Path.GetDirectoryName(logPath));
            using (StreamWriter writer = new StreamWriter(logPath, true))
            {
                writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ExcelImportExportViewController.OnActivated 开始执行");
                writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 当前视图类型: {View.GetType().FullName}");
                writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 当前视图对象类型: {View.ObjectTypeInfo.Type.FullName}");

                // 获取当前视图的对象类型
                _objectType = View.ObjectTypeInfo.Type;

                // 检查对象类型是否带有ExcelImportExportAttribute属性
                writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 开始检查ExcelImportExportAttribute");
                _excelImportExportAttribute = _objectType.GetCustomAttribute<ExcelImportExportAttribute>();
                writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ExcelImportExportAttribute: {( _excelImportExportAttribute != null ? "找到" : "未找到" )}");

                // 检查是否有其他方式可以获取ExcelImportExportAttribute
                var attributes = _objectType.GetCustomAttributes(true);
                writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 对象类型的自定义属性数量: {attributes.Length}");
                foreach (var attr in attributes)
                {
                    writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 自定义属性: {attr.GetType().FullName}");
                }

                // 如果带有该属性，则激活动作，否则禁用动作
                bool hasAttribute = _excelImportExportAttribute != null;
                writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] hasAttribute: {hasAttribute}");
                
                // 强制激活动作，不管是否有ExcelImportExportAttribute，用于测试
                _exportAction.Active["HasExcelImportExportAttribute"] = true;
                _importAction.Active["HasExcelImportExportAttribute"] = true;
                
                writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 导出动作Active状态: {_exportAction.Active}");
                writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 导入动作Active状态: {_importAction.Active}");

                if (hasAttribute)
                {
                    // 配置动作
                    writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 激活动作");
                    ConfigureActions();
                }
                else
                {
                    writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 禁用动作，因为没有ExcelImportExportAttribute");
                }
                writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ExcelImportExportViewController.OnActivated 执行完成");
            }
        }

        private void ConfigureActions()
        {
            // 可以在这里添加额外的配置逻辑
        }

        /// <summary>
        /// 执行导出操作，由平台特定的控制器重写
        /// </summary>
        /// <param name="sender">事件发送者</param>
        /// <param name="e">事件参数</param>
        protected virtual void ExportAction_Execute(object sender, SimpleActionExecuteEventArgs e)
        {
            try
            {
                // 使用XpoExcelHelper导出数据
                var excelHelper = new XpoExcelHelper(Application, null);
                var exportOptions = _excelImportExportAttribute?.ExportOptions ?? new XpoExcelExportOptions();

                // 导出数据到内存流
                var exportMethod = typeof(XpoExcelHelper).GetMethod("ExportToExcelStream", new[] { typeof(CriteriaOperator), typeof(XpoExcelExportOptions) });
                if (exportMethod == null)
                {
                    throw new InvalidOperationException("无法找到ExportToExcelStream方法");
                }
                
                var genericExportMethod = exportMethod.MakeGenericMethod(_objectType);
                var stream = (MemoryStream)genericExportMethod.Invoke(excelHelper, new object[] { null, exportOptions });
                
                // 显示成功消息
                Application.ShowViewStrategy.ShowMessage(
                    $"数据已成功导出，共{stream.Length}字节",
                    InformationType.Success
                );
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

        private void ImportAction_Execute(object sender, SimpleActionExecuteEventArgs e)
        {
            try
            {
                // 调用平台特定的导入功能实现
                ExecuteImportAction(e);
            }
            catch (Exception ex)
            {
                // 显示错误消息
                Application.ShowViewStrategy.ShowMessage(
                    $"导入失败：{ex.Message}",
                    InformationType.Error
                );
            }
        }

        /// <summary>
        /// 获取当前视图的对象类型
        /// </summary>
        public Type ObjectType => _objectType;
        
        /// <summary>
        /// 获取当前对象类型的ExcelImportExport属性
        /// </summary>
        public ExcelImportExportAttribute ExcelImportExportAttribute => _excelImportExportAttribute;
        
        /// <summary>
        /// 执行导入操作，由平台特定的控制器重写
        /// </summary>
        /// <param name="e">事件参数</param>
        protected virtual void ExecuteImportAction(SimpleActionExecuteEventArgs e)
        {
            // 默认实现，显示提示信息
            Application.ShowViewStrategy.ShowMessage(
                "请在具体平台控制器中实现Excel导入功能",
                InformationType.Info
            );
        }
    }
}