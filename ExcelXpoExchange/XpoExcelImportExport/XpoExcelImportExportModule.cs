using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DevExpress.Data.Filtering;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using DevExpress.ExpressApp.Model;
using DevExpress.ExpressApp.SystemModule;
using DevExpress.ExpressApp.Xpo;
using DevExpress.Persistent.Base;
using DevExpress.Xpo;

namespace XpoExcelImportExport
{
    /// <summary>
    /// XpoExcelImportExport模块，用于自动注册带有ExcelImportExportAttribute标记的XPO类的导入导出功能
    /// </summary>
    public sealed class XpoExcelImportExportModule : ModuleBase
    {
        public XpoExcelImportExportModule()
        {
            RequiredModuleTypes.Add(typeof(SystemModule));
        }

        /// <summary>
        /// 当模块被初始化时调用
        /// </summary>
        /// <param name="application">应用程序实例</param>
        public override void Setup(XafApplication application)
        {
            base.Setup(application);
            application.SetupComplete += Application_SetupComplete;
        }

        /// <summary>
        /// 当应用程序设置完成后调用
        /// </summary>
        /// <param name="sender">事件发送者</param>
        /// <param name="e">事件参数</param>
        private void Application_SetupComplete(object sender, EventArgs e)
        {
            // 这里可以添加额外的初始化逻辑
        }
    }

    // 注意：原来的ImportExportController已被拆分为两个针对不同平台的控制器：
    // 1. ImportExportWinController - 针对WinForms平台
    // 2. ImportExportBlazorController - 针对Blazor平台
    // 这两个控制器会根据当前运行的平台自动激活
}
