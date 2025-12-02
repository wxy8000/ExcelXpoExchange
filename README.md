# ExcelXpoExchange 通用模块

## 项目简�?

ExcelXpoExchange 是一个基�?DevExpress XAF (eXpressApp Framework) �?XPO (eXpress Persistent Objects) 的通用模块集合，提供了数据字典管理�?Excel 导入导出功能，旨在简化企业级应用开发�?

## 模块列表

| 模块名称 | 命名空间 | 主要功能 |
|---------|---------|---------|
| WxyXaf.Common | WxyXaf.Common | 基础通用模块，提供共享接口和类型 |
| WxyXaf.XpoExcel | WxyXaf.XpoExcel | XPO 对象�?Excel 导入导出功能 |
| WxyXaf.DataDictionaries | WxyXaf.DataDictionaries | 数据字典管理功能 |

## 安装方法

### 1. NuGet 包安�?

所有模块均已配置为 NuGet 包，可以通过以下方式安装�?

```bash
dotnet add package WxyXaf.Common
dotnet add package WxyXaf.XpoExcel
dotnet add package WxyXaf.DataDictionaries
```

### 2. 源码引用

可以直接引用源码项目�?

1. 将项目克隆到本地
2. 在解决方案中添加现有项目
3. 添加项目引用

## 快速开�?

### 1. 添加模块引用

�?XAF 应用程序�?Module.cs 文件中添加所需模块�?

```csharp
using DevExpress.ExpressApp;
using WxyXaf.XpoExcel;
using WxyXaf.DataDictionaries;

public sealed partial class YourModule : ModuleBase
{
    public YourModule()
    {
        InitializeComponent();
        
        // 添加 WxyXaf.XpoExcel 模块
        RequiredModuleTypes.Add(typeof(WxyXaf.XpoExcelModule));
        
        // 添加数据字典模块
        RequiredModuleTypes.Add(typeof(DataDictionariesModule));
    }
}
```

### 2. 配置 Excel 导入导出

在需要支�?Excel 导入导出�?XPO 类上添加 `ExcelImportExport` 特性：

```csharp
using WxyXaf.XpoExcel;

[ExcelImportExport]
public class YourBusinessObject : XPObject
{
    // 属性定�?
}
```

### 3. 使用数据字典

为业务对象属性添加数据字典支持：

```csharp
using WxyXaf.DataDictionaries;

public class YourBusinessObject : XPObject
{
    [DataDictionary("Department")] // 关联到名�?"Department" 的数据字�?
    public DataDictionaryItem Department
    {
        get => GetPropertyValue<DataDictionaryItem>(nameof(Department));
        set => SetPropertyValue(nameof(Department), value);
    }
}
```

## 模块详细使用方法

### 1. WxyXaf.Common 模块

#### 功能说明

提供共享接口和类型，是其他模块的基础�?

#### 主要类型

- **IRelatedObjectConverter**：关联对象转换器接口，用于自定义关联对象的转换逻辑
- **IPlatformService**：平台服务接口，用于获取平台相关信息

#### 使用示例

```csharp
using WxyXaf.Common;

// 实现自定义关联对象转换器
public class CustomRelatedObjectConverter : IRelatedObjectConverter
{
    public bool CanConvert(Type objectType)
    {
        return objectType == typeof(YourRelatedObject);
    }
    
    public object Convert(string value, Type objectType, IObjectSpace objectSpace)
    {
        // 实现转换逻辑
        return objectSpace.FindObject<YourRelatedObject>(CriteriaOperator.Parse("Name = ?", value));
    }
}
```

### 2. WxyXaf.XpoExcel 模块

#### 功能说明

提供 XPO 对象�?Excel 导入导出功能，支持多种导入模式和自定义配置�?

#### 主要特�?

- **ExcelImportExportAttribute**：标记类支持 Excel 导入导出
- **ExcelFieldAttribute**：配置字段的导入导出属�?
- **多种导入模式**：CreateOnly、UpdateOnly、CreateAndUpdate、DeleteAndUpdate
- **事务处理**：确保导入的原子�?
- **详细的导入结�?*：包含成功计数、失败计数和错误信息

#### 配置选项

##### ExcelImportExportAttribute

```csharp
[ExcelImportExport(ExportOptions = typeof(XpoExcelExportOptions), ImportOptions = typeof(XpoExcelImportOptions))]
public class YourBusinessObject : XPObject
{
    // 属性定�?
}
```

##### ExcelFieldAttribute

```csharp
public class YourBusinessObject : XPObject
{
    [ExcelField(Caption = "员工姓名", Order = 1, IsUnique = true)]
    public string EmployeeName
    {
        get => GetPropertyValue<string>(nameof(EmployeeName));
        set => SetPropertyValue(nameof(EmployeeName), value);
    }
    
    [ExcelField(Include = false)] // 不包含在导入导出�?
    public string InternalField
    {
        get => GetPropertyValue<string>(nameof(InternalField));
        set => SetPropertyValue(nameof(InternalField), value);
    }
}
```

#### 导入模式说明

| 模式名称 | 描述 |
|---------|------|
| CreateOnly | 仅创建新记录，跳过已存在的记�?|
| UpdateOnly | 仅更新现有记录，跳过不存在的记录 |
| CreateAndUpdate | 创建新记录并更新现有记录（默认模式） |
| DeleteAndUpdate | 删除现有记录并创建新记录 |

#### 使用示例

##### 导出数据�?Excel

```csharp
using WxyXaf.XpoExcel;

// 在控制器或服务中使用
var excelHelper = new XpoExcelHelper(Application, platformService);

// 导出到文�?
excelHelper.ExportToExcel<YourBusinessObject>("output.xlsx");

// 导出到内存流
var stream = excelHelper.ExportToExcelStream<YourBusinessObject>();
```

##### �?Excel 导入数据

```csharp
using WxyXaf.XpoExcel;

// 在控制器或服务中使用
var excelHelper = new XpoExcelHelper(Application, platformService);

// 准备导入选项
var importOptions = new XpoExcelImportOptions
{
    Mode = ImportMode.CreateAndUpdate, // 设置导入模式
    KeyMember = "EmployeeId" // 设置关键字段
};

// 从文件导�?
var result = excelHelper.ImportFromExcel<YourBusinessObject>("input.xlsx", importOptions);

// 处理导入结果
if (result.SuccessCount > 0)
{
    // 导入成功处理
}

if (result.Errors.Count > 0)
{
    // 处理错误
    foreach (var error in result.Errors)
    {
        Console.WriteLine($"�?{error.RowIndex}: {error.FieldName} - {error.ErrorMessage}");
    }
}
```

### 3. WxyXaf.DataDictionaries 模块

#### 功能说明

提供数据字典管理功能，允许动态管理业务对象的枚举值�?

#### 主要类型

- **DataDictionary**：数据字典类，包含多个字典项
- **DataDictionaryItem**：字典项类，包含名称、编码、描述等属�?
- **DataDictionaryAttribute**：用于关联业务对象属性和数据字典

#### 使用示例

##### 1. 创建数据字典

1. 运行应用程序
2. 在导航菜单中找到 "数据字典" �?
3. 创建新的数据字典，例�?"Department"（部门）
4. 在数据字典中添加字典项，例如 "销售部"�?技术部"�?人力资源�?

##### 2. 关联业务对象属�?

```csharp
using WxyXaf.DataDictionaries;

public class Employee : XPObject
{
    [DataDictionary("Department")] // 关联到名�?"Department" 的数据字�?
    public DataDictionaryItem Department
    {
        get => GetPropertyValue<DataDictionaryItem>(nameof(Department));
        set => SetPropertyValue(nameof(Department), value);
    }
}
```

##### 3. 在界面中使用

关联后，业务对象的属性将显示为下拉列表，包含数据字典中的所有项�?

##### 4. 数据字典项验�?

数据字典项名称在同一字典内自动验证唯一性，确保数据完整性�?

## 配置选项

### 1. WxyXaf.XpoExcel 配置

#### 导出选项

```csharp
var exportOptions = new XpoExcelExportOptions
{
    AutoFitColumns = true // 自动调整列宽
};

excelHelper.ExportToExcel<YourBusinessObject>("output.xlsx", null, exportOptions);
```

#### 导入选项

```csharp
var importOptions = new XpoExcelImportOptions
{
    Mode = ImportMode.CreateAndUpdate, // 导入模式
    KeyMember = "Id", // 关键字段
    StopOnError = false, // 是否遇到错误停止导入
    Encoding = Encoding.UTF8 // 文件编码
};

excelHelper.ImportFromExcel<YourBusinessObject>("input.xlsx", importOptions);
```

## 依赖关系

| 模块 | 依赖 |
|------|------|
| WxyXaf.Common | DevExpress.ExpressApp, DevExpress.ExpressApp.Xpo, DevExpress.Persistent.Base, DevExpress.Xpo |
| WxyXaf.XpoExcel | WxyXaf.Common, NPOI |
| WxyXaf.DataDictionaries | WxyXaf.Common, WxyXaf.XpoExcel |

## 常见问题

### 1. Excel 导入时出�?"字段值已存在" 错误

**原因**：启用了 `IsUnique` 属性，且导入的字段值在数据库中已存�?

**解决方案**�?
- 检�?Excel 数据，确保唯一字段值不重复
- 调整导入模式�?`UpdateOnly` �?`CreateAndUpdate`
- 移除字段�?`IsUnique` 属�?

### 2. 数据字典项不显示在下拉列表中

**原因**�?
- 数据字典名称�?`DataDictionaryAttribute` 中指定的名称不匹�?
- 数据字典中没有添加字典项
- 权限问题

**解决方案**�?
- 检查数据字典名称是否正�?
- 确保数据字典中已添加字典�?
- 检查用户权�?

### 3. Excel 导出时字段顺序不正确

**原因**：未设置 `ExcelFieldAttribute` �?`Order` 属�?

**解决方案**�?

```csharp
[ExcelField(Order = 1)]
public string Field1 { get; set; }

[ExcelField(Order = 2)]
public string Field2 { get; set; }
```

## 贡献指南

1. Fork 项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开 Pull Request

## 许可�?

本项目采�?MIT 许可�?- 查看 [LICENSE](LICENSE) 文件了解详情�?

## 联系方式

如有问题或建议，请通过以下方式联系�?

- 项目地址：https://github.com/yourusername/ExcelXpoExchange
- 邮箱：your.email@example.com

## 更新日志

### v1.0.0

- 初始发布
- 包含 WxyXaf.Common、WxyXpoExcel、WxyXaf.DataDictionaries 模块
- 实现数据字典管理功能
- 实现 XPO 对象�?Excel 导入导出功能

---

**版权所�?© 2025 ExcelXpoExchange 团队**
