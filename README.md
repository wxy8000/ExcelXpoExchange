# WxyXaf 通用类移植操作指南

本文档详细介绍了将 WxyXaf 通用类移植到新解决方案时的操作步骤。

## 1. 准备工作

### 1.1 了解通用类项目结构

| 项目名称 | 描述 | 核心组件 |
|---------|------|----------|
| WxyXaf.Common | 提供共享接口和类型 | IRelatedObjectConverter 接口 |
| WxyXaf.XpoExcel | 提供 XPO 对象的 Excel 导入导出功能 | XpoExcelHelper, ExcelImportExportViewController, XpoExcelImportExportModule |
| WxyXaf.DataDictionaries | 提供数据字典功能 | DataDictionary, DataDictionaryItem, DataDictionaryAttribute, DataDictionaryUpdater |

### 1.2 新解决方案要求

- .NET 9.0 或更高版本
- DevExpress XAF 25.1.* 或兼容版本
- 已安装以下 NuGet 包：
  - NPOI 2.7.5
  - DevExpress.ExpressApp 25.1.*
  - DevExpress.ExpressApp.Xpo 25.1.*
  - DevExpress.Persistent.Base 25.1.*
  - DevExpress.Persistent.BaseImpl.Xpo 25.1.*
  - DevExpress.Xpo 25.1.*

## 2. 移植步骤

### 2.1 移植 WxyXaf.Common 项目

1. **复制项目文件**
   - 复制 `WxyXaf.Common` 文件夹到新解决方案目录
   - 包含文件：
     - IRelatedObjectConverter.cs
     - WxyXaf.Common.csproj

2. **添加到解决方案**
   - 在 Visual Studio 中打开新解决方案
   - 右键点击解决方案 → 添加 → 现有项目
   - 选择 `WxyXaf.Common.csproj` 文件

3. **验证项目配置**
   - 检查目标框架：net9.0
   - 检查 NuGet 引用：
     - DevExpress.ExpressApp 25.1.*
     - DevExpress.ExpressApp.Xpo 25.1.*
     - DevExpress.Persistent.Base 25.1.*
     - DevExpress.Xpo 25.1.*

### 2.2 移植 WxyXaf.XpoExcel 项目

1. **复制项目文件**
   - 复制 `WxyXaf.XpoExcel` 文件夹到新解决方案目录
   - 包含文件：
     - ExcelImportExportViewController.cs
     - IPlatformService.cs
     - XpoExcelHelper.cs
     - XpoExcelImportExportModule.cs
     - WxyXaf.XpoExcel.csproj

2. **添加到解决方案**
   - 右键点击解决方案 → 添加 → 现有项目
   - 选择 `WxyXaf.XpoExcel.csproj` 文件

3. **配置项目依赖**
   - 右键点击 `WxyXaf.XpoExcel` 项目 → 添加 → 项目引用
   - 选择 `WxyXaf.Common` 项目

4. **验证项目配置**
   - 检查目标框架：net9.0
   - 检查 NuGet 引用：
     - 与 `WxyXaf.Common` 相同的 DevExpress 包
     - NPOI 2.7.5

### 2.3 移植 WxyXaf.DataDictionaries 项目（可选）

1. **复制项目文件**
   - 复制 `WxyXaf.DataDictionaries` 文件夹到新解决方案目录
   - 包含核心文件：
     - DataDictionary.cs
     - DataDictionaryItem.cs
     - DataDictionaryAttribute.cs
     - DataDictionaryUpdater.cs
     - DataDictionariesModule.cs
     - WxyXaf.DataDictionaries.csproj

2. **添加到解决方案**
   - 右键点击解决方案 → 添加 → 现有项目
   - 选择 `WxyXaf.DataDictionaries.csproj` 文件

3. **配置项目依赖**
   - 右键点击 `WxyXaf.DataDictionaries` 项目 → 添加 → 项目引用
   - 选择 `WxyXaf.Common` 项目

### 2.4 配置应用程序模块依赖

#### Blazor 应用程序

1. 打开 Blazor 模块项目（如 `YourApp.Blazor.Server`）
2. 右键点击项目 → 添加 → 项目引用
3. 添加以下引用：
   - WxyXaf.Common
   - WxyXaf.XpoExcel
   - WxyXaf.DataDictionaries（如果使用）

4. 更新 `BlazorModule.cs`：
   ```csharp
   using WxyXaf.XpoExcel;
   using WxyXaf.DataDictionaries;

   public sealed partial class YourAppBlazorModule : ModuleBase
   {
       public YourAppBlazorModule()
       {
           InitializeComponent();
       }

       // 在 Setup 方法中添加模块依赖
       protected override void Setup(ApplicationModulesManager moduleManager)
       {
           base.Setup(moduleManager);
           // 确保 WxyXafXpoExcelModule 被加载
       }
   }
   ```

5. 更新 `Startup.cs`，注册 XpoExcel 和 DataDictionaries 模块：
   ```csharp
   using WxyXaf.XpoExcel;
   using WxyXaf.DataDictionaries;

   public void ConfigureServices(IServiceCollection services)
   {
       // ... 其他配置
       
       services.AddXaf(Configuration, builder =>
       {
           builder.UseApplication<YourAppBlazorApplication>();
           builder.Modules
               // ... 其他模块
               .Add<WxyXafXpoExcelModule>() // 添加XPO Excel导入导出模块
               .Add<DataDictionariesModule>() // 添加数据字典模块（如果使用）
               // ... 其他模块
       });
       
       // ... 其他配置
   }
   ```

#### WinForms 应用程序

1. 打开 WinForms 模块项目（如 `YourApp.Win`）
2. 右键点击项目 → 添加 → 项目引用
3. 添加以下引用：
   - WxyXaf.Common
   - WxyXaf.XpoExcel
   - WxyXaf.DataDictionaries（如果使用）

4. 更新 `WinModule.cs`，添加必要的模块依赖：
   ```csharp
   using WxyXaf.XpoExcel;
   using WxyXaf.DataDictionaries;

   public sealed partial class YourAppWinModule : ModuleBase
   {
       public YourAppWinModule()
       {
           InitializeComponent();
       }

       // 在 Setup 方法中添加模块依赖
       protected override void Setup(ApplicationModulesManager moduleManager)
       {
           base.Setup(moduleManager);
           // 确保 WxyXafXpoExcelModule 被加载
       }
   }
   ```

5. 更新 WinForms 应用程序的 `Program.cs` 或 `Startup.cs`，注册 XpoExcel 和 DataDictionaries 模块：
   ```csharp
   using WxyXaf.XpoExcel;
   using WxyXaf.DataDictionaries;

   public static class Program
   {
       [STAThread]
       public static void Main(string[] arguments)
       {
           // ... 其他配置
           
           winApplication.Modules
               // ... 其他模块
               .Add<WxyXafXpoExcelModule>() // 添加XPO Excel导入导出模块
               .Add<DataDictionariesModule>(); // 添加数据字典模块（如果使用）
           
           // ... 其他配置
       }
   }
   ```

## 3. 配置 Blazor 特定组件

### 3.1 复制 Blazor 控制器

1. 复制 `BlazorExcelImportExportViewController.cs` 到 Blazor 项目的 Controllers 文件夹
2. 确保命名空间正确更新为新解决方案的命名空间

### 3.2 复制 ImportExcelDialog.razor 组件

1. 复制 `ImportExcelDialog.razor` 到 Blazor 项目的 Pages 文件夹
2. 确保命名空间正确更新
3. 检查并更新 dialog 组件中的服务引用

### 3.3 配置 _Host.cshtml 文件

1. 打开 Blazor 项目的 `Pages/_Host.cshtml` 文件
2. 在 `</body>` 标签前添加以下 JavaScript 函数：
   ```html
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
3. 确保 `title` 标签中的应用名称正确

## 4. 配置 WinForms 特定组件

1. 复制 `WinExcelImportExportViewController.cs` 到 WinForms 项目的 Controllers 文件夹
2. 确保命名空间正确更新

## 5. 配置数据模型

### 5.1 配置业务对象的 Excel 导入导出

要为业务对象启用 Excel 导入导出功能，需要在类级别添加 `ExcelImportExportAttribute`，并在需要导入导出的属性上添加 `ExcelFieldAttribute`。

```csharp
using WxyXaf.XpoExcel;
using WxyXaf.DataDictionaries;

// 在类级别添加 ExcelImportExportAttribute 以启用导入导出功能
[ExcelImportExport()]
public class YourBusinessObject : XPObject
{
    public YourBusinessObject(Session session) : base(session) { }
    
    // 基本属性配置
    [Size(50)]
    [ExcelField(Caption = "员工ID", Order = 0, IsUnique = true)]
    public string EmployeeId
    {
        get { return GetPropertyValue<string>(nameof(EmployeeId)); }
        set { SetPropertyValue(nameof(EmployeeId), value); }
    }
    
    // 带有中文标题的属性
    [Size(100)]
    [ExcelField(Caption = "姓名", Order = 1)]
    public string Name
    {
        get { return GetPropertyValue<string>(nameof(Name)); }
        set { SetPropertyValue(nameof(Name), value); }
    }
    
    // 数字类型属性
    [ExcelField(Caption = "年龄", Order = 2)]
    public int Age
    {
        get { return GetPropertyValue<int>(nameof(Age)); }
        set { SetPropertyValue(nameof(Age), value); }
    }
    
    // 日期类型属性
    [ExcelField(Caption = "入职日期", Order = 3)]
    public DateTime HireDate
    {
        get { return GetPropertyValue<DateTime>(nameof(HireDate)); }
        set { SetPropertyValue(nameof(HireDate), value); }
    }
    
    // 布尔类型属性
    [ExcelField(Caption = "是否在职", Order = 4)]
    public bool IsActive
    {
        get { return GetPropertyValue<bool>(nameof(IsActive)); }
        set { SetPropertyValue(nameof(IsActive), value); }
    }
    
    // 数据字典关联属性
    [ExcelField(Caption = "部门", Order = 5)]
    [DataDictionary("部门")] // 关联到名为"部门"的数据字典
    public DataDictionaryItem Department
    {
        get { return GetPropertyValue<DataDictionaryItem>(nameof(Department)); }
        set { SetPropertyValue(nameof(Department), value); }
    }
    
    // 其他属性...
}
```

### 5.2 配置数据字典（可选）

如果使用数据字典功能，需要确保：

1. 业务对象已添加 `DataDictionaryAttribute` 关联到特定数据字典
2. 数据字典项属性类型为 `DataDictionaryItem`
3. 在 Startup.cs 中已注册 `DataDictionariesModule`

数据字典会在应用启动时自动初始化，也可以通过 `DataDictionaryUpdater` 手动更新数据字典项。

## 6. 测试和验证

### 6.1 编译解决方案

- 运行 `Build Solution` 命令，确保没有编译错误
- 解决任何依赖关系问题

### 6.2 测试 Excel 导入导出功能

1. 启动应用程序
2. 导航到包含导入导出功能的列表视图
3. 测试导出功能：
   - 点击导出按钮
   - 检查生成的 Excel 文件格式是否正确
   - 验证数据是否完整导出

4. 测试导入功能：
   - 准备测试数据的 Excel 文件
   - 使用不同导入模式测试：
     - CreateOnly（仅创建）
     - UpdateOnly（仅更新）
     - CreateAndUpdate（创建和更新）
     - DeleteAndUpdate（删除和更新）
   - 验证导入结果是否符合预期
   - 检查是否有错误信息显示

### 6.3 测试数据字典功能（可选）

1. 测试数据字典的创建和编辑
2. 测试数据字典项的唯一性验证
3. 测试数据字典更新器功能

## 7. 常见问题和解决方案

### 7.1 编译错误：找不到命名空间

**原因**：命名空间引用不正确
**解决方案**：
- 检查项目引用是否正确添加
- 确保 using 语句中的命名空间与项目实际命名空间匹配
- 重新生成解决方案

### 7.2 Excel 导入失败："Excel文件表头与对象字段不匹配"

**原因**：Excel 文件的列标题与业务对象的 ExcelFieldAttribute.DisplayName 不匹配
**解决方案**：
- 检查 Excel 文件的列标题是否与业务对象的 DisplayName 一致
- 确保 ExcelFieldAttribute 已正确应用到业务对象的属性上

### 7.3 导入时记录重复

**原因**：IsUnique 属性未正确设置或唯一性验证逻辑有问题
**解决方案**：
- 确保业务对象的唯一字段上设置了 `[ExcelField(IsUnique = true)]`
- 检查 XpoExcelHelper.cs 中的唯一性验证逻辑

### 7.4 导出文件名显示乱码

**原因**：文件名编码问题
**解决方案**：
- 检查控制器中的文件名生成逻辑，确保使用正确的中文编码
- 示例：`$"{ObjectType.Name}_导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx"`

## 8. 维护和更新

### 8.1 更新通用类项目

1. 定期从原始仓库获取最新版本的通用类项目
2. 比较差异并合并到当前解决方案
3. 重新测试所有功能

### 8.2 日志和监控

- 建议在关键操作点添加日志记录
- 监控导入导出操作的性能和错误率
- 定期备份数据，特别是在执行大规模导入操作前

## 9. 最佳实践

1. **模块化设计**：保持通用类项目的独立性，便于维护和更新
2. **接口抽象**：使用接口抽象关键功能，便于扩展和替换
3. **单元测试**：为通用类编写单元测试，确保功能稳定
4. **文档更新**：及时更新文档，反映最新的功能和变更
5. **版本控制**：使用 Git 等版本控制工具管理通用类项目的变更

## 10. 技术支持

- 如遇到问题，请查看项目中的示例代码
- 参考 DevExpress 官方文档
- 联系项目维护人员获取支持

---

**最后更新时间**：2025-12-03
**版本**