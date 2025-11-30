using System;
using System.Reflection;
using DevExpress.Data.Filtering;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using DevExpress.Persistent.Base;
using DevExpress.Xpo;

namespace XpoExcelImportExport
{
    /// <summary>
    /// 通用Excel导入导出控制器，用于处理带有ExcelImportExportAttribute属性的XPO类
    /// </summary>
    public class ExcelImportExportViewController : ViewController<ListView>
    {
        private SimpleAction _exportAction;
        private SimpleAction _importAction;
        private ExcelImportExportAttribute _excelImportExportAttribute;
        private Type _objectType;

        public ExcelImportExportViewController()
        {
            // 创建导出动作
            _exportAction = new SimpleAction(this, "ExportToExcel", PredefinedCategory.Export)
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

            // 获取当前视图的对象类型
            _objectType = View.ObjectTypeInfo.Type;

            // 检查对象类型是否带有ExcelImportExportAttribute属性
            _excelImportExportAttribute = _objectType.GetCustomAttribute<ExcelImportExportAttribute>();

            // 如果带有该属性，则激活动作，否则禁用动作
            bool hasAttribute = _excelImportExportAttribute != null;
            _exportAction.Active["HasExcelImportExportAttribute"] = hasAttribute;
            _importAction.Active["HasExcelImportExportAttribute"] = hasAttribute;

            if (hasAttribute)
            {
                // 配置动作
                ConfigureActions();
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