# XpoExcelImportExport 通用模块

## 1. 项目简介

XpoExcelImportExport 是一个基于 DevExpress XAF (eXpressApp Framework) 和 XPO (eXpress Persistent Objects) 的通用 Excel 导入导出模块。它允许您通过简单的类装饰，为任何 XPO 类添加 Excel 导入导出功能，无需编写大量重复代码。

该模块提供了完整的 Blazor 和 WinForms 平台支持，支持多种导入模式，并提供了友好的用户界面。

## 2. 核心功能

- ✅ **通用导入导出**：支持任何带有 `[ExcelImportExport]` 属性的 XPO 类
- ✅ **平台支持**：同时支持 Blazor 和 WinForms 平台
- ✅ **多种导入模式**：支持仅创建、仅更新、创建或更新、替换所有数据
- ✅ **自动检测**：自动检测带有装饰符的类，无需手动注册
- ✅ **自定义配置**：支持自定义导出选项和导入选项
- ✅ **属性级控制**：支持在属性级别控制导入导出行为
- ✅ **友好的用户界面**：提供直观的导入导出对话框
- ✅ **完整的错误处理**：提供详细的导入结果和错误信息
- ✅ **支持大数据量**：优化的导入导出算法，支持处理大量数据

## 3. 技术架构

```
┌───────────────────────────────────────────────────────────────────┐
│                       XAF Application                           │
└───────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌───────────────────────────────────────────────────────────────────┐
│                    XpoExcelImportExportModule                    │
└───────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌───────────────────────────────────────────────────────────────────┐
│                  ExcelImportExportViewController                 │
└───────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌───────────────────────────────────────────────────────────────────┐
│                       XpoExcelHelper                            │
└───────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌───────────────────────────────────────────────────────────────────┐
│                          NPOI Library                            │
└───────────────────────────────────────────────────────────────────┘
```

## 4. 项目结构

```
XpoExcelImportExport/
├── ExcelImportExportViewController.cs   # 通用Excel导入导出控制器
├── IPlatformService.cs                # 平台服务接口
├── XpoExcelHelper.cs                  # Excel导入导出核心工具类
├── XpoExcelImportExport.csproj        # 项目文件
├── XpoExcelImportExportModule.cs      # 模块类
└── README.md                          # 本文件
```

## 5. 安装指南

### 5.1 项目拷贝步骤

#### 5.1.1 拷贝整个XpoExcelImportExport文件夹

将整个`XpoExcelImportExport`文件夹拷贝到您的解决方案目录中：

1. 在Windows资源管理器中，导航到`d:\projects\ExcelXpoExchange\ExcelXpoExchange\XpoExcelImportExport`
2. 右键点击该文件夹，选择"复制"
3. 导航到您的解决方案目录
4. 右键点击解决方案目录，选择"粘贴"

#### 5.1.2 在Visual Studio中添加项目

1. 打开您的解决方案
2. 在解决方案资源管理器中，右键点击解决方案，选择"添加" > "现有项目"
3. 导航到您刚刚拷贝的`XpoExcelImportExport`文件夹
4. 选择`XpoExcelImportExport.csproj`文件，点击"添加"

#### 5.1.3 添加项目引用

##### 5.1.3.1 模块项目引用

1. 在解决方案资源管理器中，右键点击您的模块项目，选择"添加引用"
2. 选择"项目"选项卡
3. 勾选`XpoExcelImportExport`项目，点击"确定"

##### 5.1.3.2 Blazor项目引用

1. 在解决方案资源管理器中，右键点击您的Blazor项目，选择"添加引用"
2. 选择"项目"选项卡
3. 勾选`XpoExcelImportExport`项目，点击"确定"

##### 5.1.3.3 WinForms项目引用

1. 在解决方案资源管理器中，右键点击您的WinForms项目，选择"添加引用"
2. 选择"项目"选项卡
3. 勾选`XpoExcelImportExport`项目，点击"确定"

## 6. 平台特定配置

### 6.1 Blazor平台配置

在Blazor项目中，您需要配置以下文件：

#### 6.1.1 更新_Host.cshtml文件

在Blazor项目的`Pages/_Host.cshtml`文件中添加`downloadFile` JavaScript函数：

```html
<script src="_framework/blazor.server.js"></script>
// 引入Blazor导出功能的JavaScript函数
<script>
    // 下载文件函数，用于Blazor导出功能
    function downloadFile(fileName, bytes) {
        // 将字节数组转换为Blob对象
        const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        // 创建下载链接
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        
        // 触发点击事件，自动打开浏览器的另存为对话框
        document.body.appendChild(link);
        link.click();
        
        // 清理临时资源
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }
</script>
```

#### 6.1.2 配置Startup.cs文件

在Blazor项目的`Startup.cs`文件中添加`XpoExcelImportExportModule`模块：

```csharp
services.AddXaf(Configuration, builder => {
    builder.UseApplication<YourBlazorApplication>();
    builder.Modules
        .AddConditionalAppearance()
        .AddValidation(options =>
        {
            options.AllowValidationDetailsAccess = false;
        })
        .Add<YourModule>()
        .Add<YourBlazorModule>()
        .Add<XpoExcelImportExport.XpoExcelImportExportModule>(); // 添加这一行
    // 其他配置...
});
```

#### 6.1.3 添加BlazorExcelImportExportViewController.cs

在Blazor项目的`Controllers`文件夹中添加`BlazorExcelImportExportViewController.cs`文件：

```csharp
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
using XpoExcelImportExport;

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
```

#### 6.1.4 添加ImportExcelDialog.razor

在Blazor项目的`Pages`文件夹中添加`ImportExcelDialog.razor`文件：

```razor
@page "/import-excel-dialog"
@using System.Reflection
@using System.Web
using DevExpress.ExpressApp
using DevExpress.ExpressApp.Blazor.Services
using XpoExcelImportExport
@inject IObjectSpaceFactory ObjectSpaceFactory
@inject IJSRuntime JSRuntime
@inject NavigationManager NavigationManager
@inject IServiceProvider ServiceProvider
@inject DevExpress.ExpressApp.Blazor.Services.IXafApplicationProvider XafApplicationProvider

<style>
    /* 模态对话框样式 */
    .modal-backdrop {
        background-color: rgba(0, 0, 0, 0.5);
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        z-index: 1040;
    }

    .modal {
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 1050;
        outline: 0;
    }

    .modal-dialog {
        position: relative;
        width: auto;
        max-width: 500px;
        margin: 1.75rem auto;
        pointer-events: none;
    }

    .modal-content {
        position: relative;
        display: flex;
        flex-direction: column;
        width: 100%;
        pointer-events: auto;
        background-color: #fff;
        background-clip: padding-box;
        border: 1px solid rgba(0, 0, 0, 0.2);
        border-radius: 0.5rem;
        box-shadow: 0 0.5rem 1rem rgba(0, 0, 0, 0.15);
        outline: 0;
    }

    .modal-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 1rem;
        border-bottom: 1px solid #dee2e6;
        border-top-left-radius: calc(0.5rem - 1px);
        border-top-right-radius: calc(0.5rem - 1px);
    }

    .modal-title {
        margin-bottom: 0;
        line-height: 1.5;
        font-size: 1.25rem;
        font-weight: 500;
    }

    .btn-close {
        box-sizing: content-box;
        width: 1em;
        height: 1em;
        padding: 0.25em 0.25em;
        color: #000;
        background: transparent url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16' fill='%23000'%3e%3cpath d='M.293.293a1 1 0 0 1 1.414 0L8 6.586 14.293.293a1 1 0 1 1 1.414 1.414L9.414 8l6.293 6.293a1 1 0 0 1-1.414 1.414L8 9.414l-6.293 6.293a1 1 0 0 1-1.414-1.414L6.586 8 .293 1.707a1 1 0 0 1 0-1.414z'/%3e%3c/svg%3e") center/1em auto no-repeat;
        border: 0;
        border-radius: 0.25rem;
        opacity: 0.5;
    }

    .btn-close:hover {
        color: #000;
        text-decoration: none;
        opacity: 0.75;
    }

    .modal-body {
        position: relative;
        flex: 1 1 auto;
        padding: 1rem;
    }

    .modal-footer {
        display: flex;
        align-items: center;
        justify-content: flex-end;
        padding: 1rem;
        border-top: 1px solid #dee2e6;
        border-bottom-left-radius: calc(0.5rem - 1px);
        border-bottom-right-radius: calc(0.5rem - 1px);
        gap: 0.5rem;
    }

    /* 文件选择美化 */
    .file-input-wrapper {
        position: relative;
        display: inline-block;
        width: 100%;
    }

    .file-input-custom {
        display: block;
        width: 100%;
        padding: 0.5rem 1rem;
        font-size: 0.875rem;
        font-weight: 400;
        line-height: 1.5;
        color: #495057;
        background-color: #fff;
        background-clip: padding-box;
        border: 1px solid #ced4da;
        border-radius: 0.375rem;
        transition: border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }

    .file-input-custom:hover {
        cursor: pointer;
        background-color: #f8f9fa;
    }

    .file-input {
        position: absolute;
        left: 0;
        top: 0;
        opacity: 0;
        width: 100%;
        height: 100%;
        cursor: pointer;
    }

    /* 导入模式radio样式 */
    .radio-group {
        display: flex;
        flex-direction: column;
        gap: 0.5rem;
    }

    .radio-item {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        padding: 0.5rem;
        border: 1px solid #e9ecef;
        border-radius: 0.375rem;
        transition: all 0.15s ease-in-out;
    }

    .radio-item:hover {
        background-color: #f8f9fa;
        border-color: #dee2e6;
    }

    .radio-item input[type="radio"] {
        accent-color: #0d6efd;
    }

    .radio-item label {
        margin: 0;
        cursor: pointer;
    }

    /* 导入结果样式 */
    .result-container {
        margin-top: 1rem;
        padding: 1rem;
        border-radius: 0.375rem;
    }

    .result-success {
        background-color: #d1fae5;
        border: 1px solid #a7f3d0;
        color: #065f46;
    }

    .result-error {
        background-color: #fee2e2;
        border: 1px solid #fecaca;
        color: #991b1b;
    }
</style>

<!-- 模态对话框内容 -->
<div class="modal-backdrop"></div>
<div class="modal" tabindex="-1" role="dialog" aria-labelledby="importModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="importModalLabel">从Excel导入数据</h5>
                <button type="button" class="btn-close" aria-label="Close" @onclick="Cancel"></button>
            </div>
            <div class="modal-body">
                <!-- 文件选择 -->
                <div class="mb-4">
                    <label class="form-label fw-semibold">选择Excel文件</label>
                    <div class="file-input-wrapper">
                        <div class="file-input-custom">
                            @(selectedFile != null ? selectedFile.Name : "点击选择文件")
                        </div>
                        <InputFile id="excelFile" OnChange="OnFileSelected" accept=".xlsx" class="file-input" />
                    </div>
                    <div class="text-muted text-sm mt-1">
                        支持 .xlsx 格式，最大 10MB
                    </div>
                </div>

                <!-- 导入模式 -->
                <div class="mb-4">
                    <label class="form-label fw-semibold">导入模式</label>
                    <div class="radio-group">
                        <div class="radio-item">
                            <input type="radio" id="mode-create" name="importMode" checked="@(selectedImportMode == "CreateOnly")" @onchange="HandleCreateOnlyChange" />
                            <label for="mode-create">仅创建新数据</label>
                        </div>
                        <div class="radio-item">
                            <input type="radio" id="mode-update" name="importMode" checked="@(selectedImportMode == "UpdateOnly")" @onchange="HandleUpdateOnlyChange" />
                            <label for="mode-update">仅更新已存在的数据</label>
                        </div>
                        <div class="radio-item">
                            <input type="radio" id="mode-create-update" name="importMode" checked="@(selectedImportMode == "CreateOrUpdate")" @onchange="HandleCreateOrUpdateChange" />
                            <label for="mode-create-update">创建或更新数据</label>
                        </div>
                        <div class="radio-item">
                            <input type="radio" id="mode-replace" name="importMode" checked="@(selectedImportMode == "Replace")" @onchange="HandleReplaceChange" />
                            <label for="mode-replace">替换所有数据</label>
                        </div>
                    </div>
                </div>

                <!-- 跳过重复数据 -->
                <div class="mb-4">
                    <div class="form-check">
                        <input type="checkbox" id="skipDuplicate" @bind="skipDuplicate" class="form-check-input" />
                        <label class="form-check-label" for="skipDuplicate">跳过重复数据</label>
                    </div>
                </div>

                <!-- 导入结果 -->
                @if (!string.IsNullOrEmpty(importResult))
                {
                    <div class="result-container @(isSuccess ? "result-success" : "result-error")">
                        @importResult
                    </div>
                }
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" @onclick="Cancel">取消</button>
                <button type="button" class="btn btn-primary" disabled="@isImporting" @onclick="ImportData">
                    @if (isImporting)
                    {
                        <span class="d-flex align-items-center gap-2">
                            <span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span>
                            导入中...
                        </span>
                    }
                    else
                    {
                        <span>导入数据</span>
                    }
                </button>
            </div>
        </div>
    </div>
</div>

<script>
    // 下载文件函数，用于Blazor导出功能
    function downloadFile(fileName, bytes) {
        // 将字节数组转换为Blob对象
        const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        // 创建下载链接
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        
        // 触发点击事件，自动打开浏览器的另存为对话框
        document.body.appendChild(link);
        link.click();
        
        // 清理临时资源
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }
</script>

@code {
    private IBrowserFile selectedFile;
    private bool isImporting;
    private string importResult;
    private bool isSuccess;
    private string selectedImportMode = "CreateOrUpdate";
    private bool skipDuplicate = true;
    private Type objectType;
    private string objectTypeName;
    private string listViewId;
    
    protected override void OnInitialized()
    {
        base.OnInitialized();
        
        // 解析URL参数
        var uri = new Uri(NavigationManager.Uri);
        var query = HttpUtility.ParseQueryString(uri.Query);
        objectTypeName = query["objectType"];
        
        if (!string.IsNullOrEmpty(objectTypeName))
        {
            // 根据类型名称获取Type对象
            // 遍历所有加载的程序集，查找匹配的类型
            foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                objectType = assembly.GetType(objectTypeName);
                if (objectType != null)
                {
                    break;
                }
            }
            
            if (objectType == null)
            {
                // 如果找不到类型，尝试使用当前程序集的命名空间作为前缀
                var currentAssembly = typeof(ImportExcelDialog).Assembly;
                var simpleTypeName = objectTypeName.Split('.').Last();
                objectType = currentAssembly.GetType($"{currentAssembly.GetName().Name}.BusinessObjects.{simpleTypeName}");
            }
            
            // 生成ListViewId
            if (objectType != null)
            {
                listViewId = $"{objectType.Name}_ListView";
            }
            else
            {
                // 如果找不到类型，显示错误信息
                importResult = $"无法找到类型：{objectTypeName}";
                isSuccess = false;
            }
        }
        else
        {
            // 如果没有传递objectType参数，显示错误信息
            importResult = "未指定要导入的对象类型";
            isSuccess = false;
        }
    }
    
    private void HandleCreateOnlyChange()
    {
        selectedImportMode = "CreateOnly";
    }
    
    private void HandleUpdateOnlyChange()
    {
        selectedImportMode = "UpdateOnly";
    }
    
    private void HandleCreateOrUpdateChange()
    {
        selectedImportMode = "CreateOrUpdate";
    }
    
    private void HandleReplaceChange()
    {
        selectedImportMode = "Replace";
    }
    
    private void OnFileSelected(InputFileChangeEventArgs e)
    {
        selectedFile = e.File;
    }
    
    private async Task ImportData()
    {
        if (selectedFile == null)
        {
            importResult = "请先选择Excel文件";
            isSuccess = false;
            return;
        }
        
        if (objectType == null)
        {
            importResult = "无法获取对象类型";
            isSuccess = false;
            return;
        }
        
        isImporting = true;
        importResult = string.Empty;
        
        try
        {
            // 1. 先将异步流转换为字节数组
            using var fileStream = selectedFile.OpenReadStream(maxAllowedSize: 10 * 1024 * 1024);
            using var memoryStream = new MemoryStream();
            await fileStream.CopyToAsync(memoryStream);
            memoryStream.Position = 0;
            
            // 2. 使用Task.Run将导入操作放在后台线程执行
            await Task.Run(() =>
            {
                // 获取应用程序实例
                var application = XafApplicationProvider.GetApplication();
                if (application == null)
                {
                    importResult = "无法获取应用程序实例";
                    isSuccess = false;
                    return;
                }
                
                // 创建XpoExcelHelper实例
                var platformService = ServiceProvider.GetService<IPlatformService>();
                var excelHelper = new XpoExcelHelper(application, platformService);
                
                // 保存文件到临时位置
                string tempFilePath = Path.GetTempFileName() + ".xlsx";
                using (var tempFileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
                {
                    memoryStream.Position = 0;
                    memoryStream.CopyTo(tempFileStream);
                }
                
                try
                {
                    // 执行导入
                    var importOptions = new XpoExcelImportOptions
                    {
                        Mode = selectedImportMode switch
                        {
                            "CreateOnly" => ImportMode.CreateOnly,
                            "UpdateOnly" => ImportMode.UpdateOnly,
                            "CreateOrUpdate" => ImportMode.CreateAndUpdate,
                            "Replace" => ImportMode.DeleteAndUpdate,
                            _ => ImportMode.CreateAndUpdate
                        },
                        StopOnError = false
                    };
                    
                    // 使用反射调用泛型方法
                    var importMethod = typeof(XpoExcelHelper).GetMethod("ImportFromExcel", new[] { typeof(string), typeof(XpoExcelImportOptions) });
                    if (importMethod == null)
                    {
                        importResult = "无法找到ImportFromExcel方法";
                        isSuccess = false;
                        return;
                    }
                    
                    var genericImportMethod = importMethod.MakeGenericMethod(objectType);
                    var result = (ImportResult)genericImportMethod.Invoke(excelHelper, new object[] { tempFilePath, importOptions });
                    
                    // 处理导入结果
                    importResult = result.HasErrors 
                        ? $"导入失败，成功{result.SuccessCount}条，失败{result.FailureCount}条，错误信息：{string.Join(Environment.NewLine, result.Errors.Select(e => e.ErrorMessage))}"
                        : $"导入成功，共{result.SuccessCount}条记录";
                    isSuccess = !result.HasErrors;
                }
                finally
                {
                    // 删除临时文件
                    if (File.Exists(tempFilePath))
                    {
                        File.Delete(tempFilePath);
                    }
                }
            });
        }
        catch (Exception ex)
        {
            importResult = $"导入失败：{ex.Message}{Environment.NewLine}{ex.StackTrace}";
            isSuccess = false;
        }
        finally
        {
            isImporting = false;
            
            // 导入完成后，等待2秒让用户看到结果，然后自动返回列表视图
            if (isSuccess)
            {
                await Task.Delay(2000);
                NavigationManager.NavigateTo($"/{listViewId}");
            }
        }
    }
    
    private void Cancel()
    {
        NavigationManager.NavigateTo($"/{listViewId}");
    }
}
```

### 6.2 WinForms平台配置

在WinForms项目中，您需要配置以下文件：

#### 6.2.1 配置Startup.cs文件

在WinForms项目的`Startup.cs`文件中添加`XpoExcelImportExportModule`模块：

```csharp
builder.Modules
    .AddConditionalAppearance()
    .AddValidation(options =>
    {
        options.AllowValidationDetailsAccess = false;
    })
    .Add<ExcelXpoExchange.Module.ExcelXpoExchangeModule>()
    .Add<ExcelXpoExchangeWinModule>()
    .Add<XpoExcelImportExport.XpoExcelImportExportModule>(); // 添加这一行
```

#### 6.2.2 添加WinExcelImportExportViewController.cs

在WinForms项目的`Controllers`文件夹中添加`WinExcelImportExportViewController.cs`文件：

```csharp
using System;
using System.IO;
using System.Windows.Forms;
using DevExpress.Data.Filtering;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using XpoExcelImportExport;

namespace YourWinFormsProject.Controllers
{
    /// <summary>
    /// WinForms版通用Excel导入导出控制器
    /// </summary>
    public class WinExcelImportExportViewController : ExcelImportExportViewController
    {
        /// <summary>
        /// 执行导入操作，实现WinForms平台的Excel导入功能
        /// </summary>
        /// <param name="e">事件参数</param>
        protected override void ExecuteImportAction(SimpleActionExecuteEventArgs e)
        {
            try
            {
                // 显示WinForms文件选择对话框
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Title = "选择Excel文件";
                    openFileDialog.Filter = "Excel文件 (*.xlsx)|*.xlsx|Excel 97-2003文件 (*.xls)|*.xls|所有文件 (*.*)|*.*";
                    openFileDialog.FilterIndex = 1;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = openFileDialog.FileName;

                        // 显示导入模式选择对话框
                        using (var dialog = new ImportModeDialog())
                        {
                            if (dialog.ShowDialog() == DialogResult.OK)
                            {
                                ImportMode importMode = dialog.SelectedMode;

                                // 创建XpoExcelHelper实例
                                var excelHelper = new XpoExcelHelper(Application, null);

                                // 执行导入
                                var importOptions = new XpoExcelImportOptions
                                {
                                    Mode = importMode,
                                    StopOnError = false
                                };

                                // 使用反射调用泛型方法
                                var importMethod = typeof(XpoExcelHelper).GetMethod("ImportFromExcel", new[] { typeof(string), typeof(XpoExcelImportOptions) });
                                if (importMethod == null)
                                {
                                    Application.ShowViewStrategy.ShowMessage(
                                        "无法找到ImportFromExcel方法",
                                        InformationType.Error
                                    );
                                    return;
                                }

                                var genericImportMethod = importMethod.MakeGenericMethod(ObjectType);
                                var result = (ImportResult)genericImportMethod.Invoke(excelHelper, new object[] { filePath, importOptions });

                                // 显示导入结果
                                Application.ShowViewStrategy.ShowMessage(
                                    result.HasErrors 
                                        ? $"导入失败，成功{result.SuccessCount}条，失败{result.FailureCount}条，错误信息：{string.Join(Environment.NewLine, result.Errors.Select(e => e.ErrorMessage))}"
                                        : $"导入成功，共{result.SuccessCount}条记录",
                                    result.HasErrors ? InformationType.Error : InformationType.Success
                                );

                                // 刷新视图，显示新导入的数据
                                View.RefreshDataSource();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Application.ShowViewStrategy.ShowMessage($"导入Excel失败，{ex.Message}", InformationType.Error);
            }
        }
        
        /// <summary>
        /// 重写导出按钮点击事件，实现WinForms平台的Excel导出功能
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

                // 显示WinForms文件保存对话框
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Title = "保存Excel文件";
                    saveFileDialog.Filter = "Excel文件 (*.xlsx)|*.xlsx|Excel 97-2003文件 (*.xls)|*.xls|所有文件 (*.*)|*.*";
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.RestoreDirectory = true;
                    saveFileDialog.FileName = $"{ObjectType.Name}_导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = saveFileDialog.FileName;

                        // 导出数据到文件
                        var exportMethod = typeof(XpoExcelHelper).GetMethod("ExportToExcel", new[] { typeof(string), typeof(CriteriaOperator), typeof(XpoExcelExportOptions) });
                        if (exportMethod == null)
                        {
                            throw new InvalidOperationException("无法找到ExportToExcel方法");
                        }
                        
                        var genericExportMethod = exportMethod.MakeGenericMethod(ObjectType);
                        genericExportMethod.Invoke(excelHelper, new object[] { filePath, null, exportOptions });
                        
                        // 显示成功消息
                        Application.ShowViewStrategy.ShowMessage(
                            $"数据已成功导出到：{filePath}",
                            InformationType.Success
                        );
                    }
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

        /// <summary>
        /// 导入模式选择对话框
        /// </summary>
        private class ImportModeDialog : Form
        {
            private RadioButton rbCreateOnly;
            private RadioButton rbUpdateOnly;
            private RadioButton rbCreateOrUpdate;
            private RadioButton rbReplace;
            private Button btnOK;
            private Button btnCancel;
            private Label label1;

            public ImportMode SelectedMode { get; private set; }

            public ImportModeDialog()
            {
                InitializeComponent();
                SelectedMode = ImportMode.CreateAndUpdate;
                rbCreateOrUpdate.Checked = true;
            }

            private void InitializeComponent()
            {
                this.label1 = new System.Windows.Forms.Label();
                this.rbCreateOnly = new System.Windows.Forms.RadioButton();
                this.rbUpdateOnly = new System.Windows.Forms.RadioButton();
                this.rbCreateOrUpdate = new System.Windows.Forms.RadioButton();
                this.rbReplace = new System.Windows.Forms.RadioButton();
                this.btnOK = new System.Windows.Forms.Button();
                this.btnCancel = new System.Windows.Forms.Button();
                this.SuspendLayout();
                // 
                // label1
                // 
                this.label1.AutoSize = true;
                this.label1.Location = new System.Drawing.Point(12, 18);
                this.label1.Name = "label1";
                this.label1.Size = new System.Drawing.Size(82, 15);
                this.label1.TabIndex = 0;
                this.label1.Text = "请选择导入模式：";
                // 
                // rbCreateOnly
                // 
                this.rbCreateOnly.AutoSize = true;
                this.rbCreateOnly.Location = new System.Drawing.Point(30, 47);
                this.rbCreateOnly.Name = "rbCreateOnly";
                this.rbCreateOnly.Size = new System.Drawing.Size(113, 19);
                this.rbCreateOnly.TabIndex = 1;
                this.rbCreateOnly.TabStop = true;
                this.rbCreateOnly.Text = "仅创建新数据";
                this.rbCreateOnly.UseVisualStyleBackColor = true;
                this.rbCreateOnly.CheckedChanged += new System.EventHandler(this.rbCreateOnly_CheckedChanged);
                // 
                // rbUpdateOnly
                // 
                this.rbUpdateOnly.AutoSize = true;
                this.rbUpdateOnly.Location = new System.Drawing.Point(30, 72);
                this.rbUpdateOnly.Name = "rbUpdateOnly";
                this.rbUpdateOnly.Size = new System.Drawing.Size(137, 19);
                this.rbUpdateOnly.TabIndex = 2;
                this.rbUpdateOnly.TabStop = true;
                this.rbUpdateOnly.Text = "仅更新已存在的数据";
                this.rbUpdateOnly.UseVisualStyleBackColor = true;
                this.rbUpdateOnly.CheckedChanged += new System.EventHandler(this.rbUpdateOnly_CheckedChanged);
                // 
                // rbCreateOrUpdate
                // 
                this.rbCreateOrUpdate.AutoSize = true;
                this.rbCreateOrUpdate.Location = new System.Drawing.Point(30, 97);
                this.rbCreateOrUpdate.Name = "rbCreateOrUpdate";
                this.rbCreateOrUpdate.Size = new System.Drawing.Size(113, 19);
                this.rbCreateOrUpdate.TabIndex = 3;
                this.rbCreateOrUpdate.TabStop = true;
                this.rbCreateOrUpdate.Text = "创建或更新数据";
                this.rbCreateOrUpdate.UseVisualStyleBackColor = true;
                this.rbCreateOrUpdate.CheckedChanged += new System.EventHandler(this.rbCreateOrUpdate_CheckedChanged);
                // 
                // rbReplace
                // 
                this.rbReplace.AutoSize = true;
                this.rbReplace.Location = new System.Drawing.Point(30, 122);
                this.rbReplace.Name = "rbReplace";
                this.rbReplace.Size = new System.Drawing.Size(113, 19);
                this.rbReplace.TabIndex = 4;
                this.rbReplace.TabStop = true;
                this.rbReplace.Text = "替换所有数据";
                this.rbReplace.UseVisualStyleBackColor = true;
                this.rbReplace.CheckedChanged += new System.EventHandler(this.rbReplace_CheckedChanged);
                // 
                // btnOK
                // 
                this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.btnOK.Location = new System.Drawing.Point(87, 156);
                this.btnOK.Name = "btnOK";
                this.btnOK.Size = new System.Drawing.Size(75, 23);
                this.btnOK.TabIndex = 5;
                this.btnOK.Text = "确定";
                this.btnOK.UseVisualStyleBackColor = true;
                // 
                // btnCancel
                // 
                this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                this.btnCancel.Location = new System.Drawing.Point(168, 156);
                this.btnCancel.Name = "btnCancel";
                this.btnCancel.Size = new System.Drawing.Size(75, 23);
                this.btnCancel.TabIndex = 6;
                this.btnCancel.Text = "取消";
                this.btnCancel.UseVisualStyleBackColor = true;
                // 
                // ImportModeDialog
                // 
                this.AcceptButton = this.btnOK;
                this.CancelButton = this.btnCancel;
                this.ClientSize = new System.Drawing.Size(255, 191);
                this.Controls.Add(this.btnCancel);
                this.Controls.Add(this.btnOK);
                this.Controls.Add(this.rbReplace);
                this.Controls.Add(this.rbCreateOrUpdate);
                this.Controls.Add(this.rbUpdateOnly);
                this.Controls.Add(this.rbCreateOnly);
                this.Controls.Add(this.label1);
                this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.Name = "ImportModeDialog";
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                this.Text = "选择导入模式";
                this.ResumeLayout(false);
                this.PerformLayout();
            }

            private void rbCreateOnly_CheckedChanged(object sender, EventArgs e)
            {
                if (rbCreateOnly.Checked)
                    SelectedMode = ImportMode.CreateOnly;
            }

            private void rbUpdateOnly_CheckedChanged(object sender, EventArgs e)
            {
                if (rbUpdateOnly.Checked)
                    SelectedMode = ImportMode.UpdateOnly;
            }

            private void rbCreateOrUpdate_CheckedChanged(object sender, EventArgs e)
            {
                if (rbCreateOrUpdate.Checked)
                    SelectedMode = ImportMode.CreateAndUpdate;
            }

            private void rbReplace_CheckedChanged(object sender, EventArgs e)
            {
                if (rbReplace.Checked)
                    SelectedMode = ImportMode.DeleteAndUpdate;
            }
        }
    }
}
```

## 7. 使用指南

### 7.1 类装饰配置

要为您的XPO类添加Excel导入导出功能，只需添加`[ExcelImportExport]`属性：

```csharp
using DevExpress.Xpo;
using XpoExcelImportExport;

[ExcelImportExport]
public class Person : XPObject
{
    public Person(Session session) : base(session) { }
    
    private string _Name;
    public string Name
    {
        get => _Name;
        set => SetPropertyValue(nameof(Name), ref _Name, value);
    }
    
    private DateTime _BirthDate;
    public DateTime BirthDate
    {
        get => _BirthDate;
        set => SetPropertyValue(nameof(BirthDate), ref _BirthDate, value);
    }
    
    private string _Email;
    public string Email
    {
        get => _Email;
        set => SetPropertyValue(nameof(Email), ref _Email, value);
    }
}
```

### 7.2 属性级配置

您可以使用`[ExcelField]`属性在属性级别控制导入导出行为：

```csharp
using DevExpress.Xpo;
using XpoExcelImportExport;

[ExcelImportExport]
public class Product : XPObject
{
    public Product(Session session) : base(session) { }
    
    private string _ProductCode;
    [ExcelField(Caption = "产品编码", Order = 1)]
    public string ProductCode
    {
        get => _ProductCode;
        set => SetPropertyValue(nameof(ProductCode), ref _ProductCode, value);
    }
    
    private string _ProductName;
    [ExcelField(Caption = "产品名称", Order = 2)]
    public string ProductName
    {
        get => _ProductName;
        set => SetPropertyValue(nameof(ProductName), ref _ProductName, value);
    }
    
    private decimal _Price;
    [ExcelField(Caption = "产品价格", Order = 3)]
    public decimal Price
    {
        get => _Price;
        set => SetPropertyValue(nameof(Price), ref _Price, value);
    }
    
    private bool _IsActive;
    [ExcelField(Caption = "是否启用", Order = 4)]
    public bool IsActive
    {
        get => _IsActive;
        set => SetPropertyValue(nameof(IsActive), ref _IsActive, value);
    }
    
    [ExcelField(Include = false)]
    public string InternalNotes
    {
        get => GetPropertyValue<string>(nameof(InternalNotes));
        set => SetPropertyValue(nameof(InternalNotes), value);
    }
}
```

### 7.3 导入导出选项配置

您可以在`[ExcelImportExport]`属性中配置导入导出选项：

```csharp
using DevExpress.Xpo;
using XpoExcelImportExport;

[ExcelImportExport(
    ExportOptions = new XpoExcelExportOptions {
        AutoFitColumns = true,
        ExcludedMembers = new[] { "InternalNotes", "CreatedBy" }
    },
    ImportOptions = new XpoExcelImportOptions {
        Mode = ImportMode.CreateAndUpdate,
        KeyMember = "ProductCode",
        StopOnError = false
    }
)]
public class Product : XPObject
{
    // 类定义...
}
```

#### 7.3.1 导出选项

| 属性 | 描述 | 默认值 |
|------|------|--------|
| `AutoFitColumns` | 是否自动调整列宽 | `true` |
| `ExcludedMembers` | 要排除的成员列表 | `null` |
| `IncludedMembers` | 要包含的成员列表 | `null` |

#### 7.3.2 导入选项

| 属性 | 描述 | 默认值 |
|------|------|--------|
| `Mode` | 导入模式 | `ImportMode.CreateAndUpdate` |
| `KeyMember` | 关键字段名，用于查找现有对象 | `null` |
| `StopOnError` | 是否在遇到错误时停止导入 | `false` |

#### 7.3.3 导入模式

| 模式 | 描述 |
|------|------|
| `CreateOnly` | 仅创建新对象，跳过现有对象 |
| `UpdateOnly` | 仅更新现有对象，跳过新对象 |
| `CreateAndUpdate` | 创建新对象并更新现有对象 |
| `DeleteAndUpdate` | 删除所有现有对象，然后创建新对象 |

### 7.4 运行和测试

1. 编译并运行您的应用程序
2. 导航到带有`[ExcelImportExport]`属性的类的列表视图
3. 您将看到"导出到Excel"和"从Excel导入"按钮
4. 点击"导出到Excel"按钮，将数据导出到Excel文件
5. 点击"从Excel导入"按钮，从Excel文件导入数据

## 8. 功能说明

### 8.1 导出功能

导出功能允许您将XPO对象的数据导出到Excel文件中。导出过程包括：

1. 自动检测带有`[ExcelImportExport]`属性的XPO类
2. 自动生成Excel表头，使用属性的Caption或名称
3. 自动调整列宽（可配置）
4. 支持关联对象的显示，显示关联对象的默认成员或关键字段
5. 支持属性级别的导出控制
6. 支持自定义导出选项

### 8.2 导入功能

导入功能允许您从Excel文件导入数据到XPO对象中。导入过程包括：

1. 自动检测Excel文件的格式（.xlsx或.xls）
2. 自动匹配Excel表头和XPO属性
3. 支持多种导入模式
4. 提供详细的导入结果和错误信息
5. 支持大数据量导入
6. 支持属性级别的导入控制
7. 支持自定义导入选项

### 8.3 导入结果

导入结果包括：

- 成功导入的记录数
- 导入失败的记录数
- 详细的错误信息，包括行索引、字段名和错误消息

## 9. 常见问题和解决方案

### 9.1 导入失败，错误信息："String '' was not recognized as a valid DateTime."

**解决方案**：模块已经处理了这种情况，空字符串会被转换为`DateTime.MinValue`，不会导致导入失败。

### 9.2 导入失败，错误信息："无法获取应用程序实例"

**解决方案**：确保您的应用程序正确注册了XafApplication实例。在Blazor平台上，模块使用`IXafApplicationProvider`服务获取应用程序实例。

### 9.3 导入失败，错误信息："无法找到类型"

**解决方案**：确保您的类型名称包含完整的命名空间，或者类型在当前程序集或常用命名空间中。

### 9.4 导入对话框不显示

**解决方案**：确保您的平台特定控制器已正确注册，并且您的类已添加`[ExcelImportExport]`属性。

### 9.5 导出失败，错误信息："无法找到ExportToExcelStream方法"

**解决方案**：确保您的XpoExcelHelper类中包含ExportToExcelStream方法，或者您的方法签名与控制器中调用的签名匹配。

### 9.6 Blazor平台导出时不显示另存为对话框

**解决方案**：确保您已在_Host.cshtml文件中添加了`downloadFile` JavaScript函数，并且该函数能够被正确调用。

## 10. 技术要求

- .NET 8.0 或更高版本
- DevExpress XAF 25.1 或更高版本
- DevExpress XPO 25.1 或更高版本
- NPOI 2.6.0 或更高版本

## 11. 支持的平台

- ✅ Blazor Server
- ✅ WinForms
- ✅ .NET 8.0+
- ✅ .NET 9.0+

## 12. 许可证

MIT License

## 13. 贡献指南

欢迎提交 Issue 和 Pull Request！

### 13.1 提交 Issue

1. 确保您的 Issue 描述清晰、具体
2. 包含重现步骤
3. 包含错误信息和堆栈跟踪
4. 包含您的环境信息（.NET 版本、DevExpress 版本等）

### 13.2 提交 Pull Request

1. Fork 仓库
2. 创建您的功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开 Pull Request

## 14. 联系方式

如有问题或建议，请通过以下方式联系我们：

- Email: your-email@example.com
- GitHub: https://github.com/your-username/xpo-excel-import-export

## 15. 更新日志

### v1.0.0

- 初始版本
- 支持 Blazor 和 WinForms 平台
- 支持通用导入导出功能
- 支持多种导入模式
- 支持属性级控制
- 支持自定义配置选项

## 16. 后续计划

- [ ] 支持更多 Excel 格式（如 CSV）
- [ ] 支持更复杂的关联对象导入
- [ ] 支持 Excel 模板生成
- [ ] 支持导入预览功能
- [ ] 支持批量导入优化
- [ ] 支持更多导入导出选项
- [ ] 支持数据验证

## 17. 附加资源

- [DevExpress XAF 文档](https://docs.devexpress.com/eXpressAppFramework/)
- [DevExpress XPO 文档](https://docs.devexpress.com/XPO/)
- [NPOI 文档](https://github.com/nissl-lab/npoi/wiki)
- [.NET 文档](https://docs.microsoft.com/zh-cn/dotnet/)

---

**XpoExcelImportExport 通用模块**

*让 Excel 导入导出变得简单！*