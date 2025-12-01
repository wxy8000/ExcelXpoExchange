# 通用模块NuGet包移植操作指南

## 1. NuGet包概述

### 1.1 包列表

| 包名称 | 版本 | 描述 | 依赖关系 |
|--------|------|------|----------|
| Wxy.Common | 1.0.0 | 提供共享接口和类型 | DevExpress.ExpressApp, DevExpress.Xpo |
| WxyXaf.DataDictionaries | 1.0.0 | 提供数据字典管理功能 | Wxy.Common, DevExpress.ExpressApp |
| WxyXpoExcel | 1.0.0 | 提供XPO对象的Excel导入导出功能 | Wxy.Common, DevExpress.ExpressApp, NPOI |

### 1.2 包位置

NuGet包生成在以下目录：
```
d:\projects\ExcelXpoExchange\ExcelXpoExchange\nupkg\
```

## 2. 安装NuGet包

### 2.1 方式一：使用本地NuGet源

#### 步骤1：添加本地NuGet源
1. 打开 Visual Studio
2. 选择 "工具" → "NuGet包管理器" → "包管理器设置"
3. 在 "NuGet包管理器" → "包源" 中，点击 "添加"
4. 名称：输入 "LocalNuGet"
5. 源：浏览到 `d:\projects\ExcelXpoExchange\ExcelXpoExchange\nupkg\` 目录
6. 点击 "更新" → "确定"

#### 步骤2：安装NuGet包
1. 在新解决方案中，右键点击项目 → "管理NuGet包"
2. 选择 "LocalNuGet" 包源
3. 搜索并安装以下包：
   - Wxy.Common
   - WxyXaf.DataDictionaries
   - WxyXpoExcel
4. 点击 "安装"，接受许可证协议

### 2.2 方式二：直接安装.nupkg文件

#### 步骤1：复制.nupkg文件
1. 将 `nupkg` 目录中的.nupkg文件复制到新解决方案的任意目录

#### 步骤2：安装包
1. 在新解决方案中，右键点击项目 → "管理NuGet包"
2. 点击 "浏览" → "本地" → "选择文件"
3. 浏览到复制的.nupkg文件，依次安装三个包

## 3. 配置模块

### 3.1 模块注册

在新解决方案的Module.cs文件中注册模块：

```csharp
using System;
using DevExpress.ExpressApp;
using WxyXaf.DataDictionaries;
using WxyXpoExcel;

namespace NewSolution.Module
{
    public sealed class NewSolutionModule : ModuleBase
    {
        public NewSolutionModule()
        {
            RequiredModuleTypes.Add(typeof(SystemModule.SystemModule));
            RequiredModuleTypes.Add(typeof(DevExpress.ExpressApp.Objects.BusinessClassLibraryCustomizationModule));
            RequiredModuleTypes.Add(typeof(DevExpress.ExpressApp.ConditionalAppearance.ConditionalAppearanceModule));
            RequiredModuleTypes.Add(typeof(DevExpress.ExpressApp.Validation.ValidationModule));
            
            // 注册自定义模块
            RequiredModuleTypes.Add(typeof(DataDictionariesModule));
            RequiredModuleTypes.Add(typeof(WxyXpoExcelModule));
        }
        // 其他代码...
    }
}
```

### 3.2 配置文件设置

在App.config/Web.config中添加以下配置：

```xml
<appSettings>
    <!-- 设置日志编码为UTF-8，解决中文乱码问题 -->
    <add key="eXpressAppFrameworkTraceEncoding" value="UTF-8"/>
    <!-- 设置日志级别 -->
    <add key="eXpressAppFrameworkTraceLevel" value="3"/>
</appSettings>
```

### 3.3 相关对象转换器注册

在ExcelImportExportViewController中注册DataDictionaryItemConverter：

```csharp
using System;
using System.IO;
using System.Windows.Forms;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using WxyXpoExcel;
using WxyXaf.DataDictionaries;

namespace NewSolution.Win.Controllers
{
    public class WinExcelImportExportViewController : ExcelImportExportViewController
    {
        protected override void ExecuteImportAction(SimpleActionExecuteEventArgs e)
        {
            try
            {
                // 显示文件选择对话框
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Title = "选择Excel文件";
                    openFileDialog.Filter = "Excel文件 (*.xlsx)|*.xlsx|Excel 97-2003文件 (*.xls)|*.xls";
                    
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = openFileDialog.FileName;
                        
                        // 创建XpoExcelHelper实例，并注册DataDictionaryItemConverter
                        var dataDictionaryItemConverter = new DataDictionaryItemConverter();
                        var excelHelper = new XpoExcelHelper(Application, null, new[] { dataDictionaryItemConverter });
                        
                        // 执行导入操作...
                    }
                }
            }
            catch (Exception ex)
            {
                Application.ShowViewStrategy.ShowMessage(
                    $"导入失败：{ex.Message}",
                    InformationType.Error
                );
            }
        }
        // 其他代码...
    }
}
```

## 4. 使用模块功能

### 4.1 数据字典使用示例

#### 步骤1：在业务对象中使用DataDictionaryAttribute

```csharp
using System;
using DevExpress.Persistent.Base;
using DevExpress.Persistent.BaseImpl;
using DevExpress.Xpo;
using WxyXaf.DataDictionaries;
using WxyXpoExcel;

namespace NewSolution.Module.BusinessObjects
{
    [DefaultClassOptions]
    [ExcelImportExport()]
    public class Employee : BaseObject
    {
        public Employee(Session session)
            : base(session)
        {
        }
        
        [Size(50)]
        [ExcelField(Caption = "员工ID", Order = 0)]
        public string EmployeeId
        {
            get => fEmployeeId;
            set => SetPropertyValue(nameof(EmployeeId), ref fEmployeeId, value);
        }
        string fEmployeeId;
        
        [Size(100)]
        [ExcelField(Caption = "姓名", Order = 1)]
        public string Name
        {
            get => fName;
            set => SetPropertyValue(nameof(Name), ref fName, value);
        }
        string fName;
        
        [ExcelField(Caption = "部门", Order = 2)]
        [DataDictionary("部门")]
        public DataDictionaryItem Department
        {
            get => fDepartment;
            set => SetPropertyValue(nameof(Department), ref fDepartment, value);
        }
        DataDictionaryItem fDepartment;
        
        [ExcelField(Caption = "职位", Order = 3)]
        [DataDictionary("职位")]
        public DataDictionaryItem Position
        {
            get => fPosition;
            set => SetPropertyValue(nameof(Position), ref fPosition, value);
        }
        DataDictionaryItem fPosition;
    }
}
```

#### 步骤2：运行应用程序，管理数据字典

1. 启动应用程序
2. 在导航菜单中，找到并打开 "数据字典" 和 "数据字典项"
3. 创建数据字典（如：部门、职位）
4. 为每个数据字典添加数据项

### 4.2 Excel导入导出示例

#### 步骤1：使用Excel导入功能

1. 在导航菜单中，打开需要导入数据的列表视图
2. 点击 "编辑" → "从Excel导入"
3. 选择Excel文件，点击 "打开"
4. 选择导入模式，点击 "确定"
5. 等待导入完成，查看导入结果

#### 步骤2：使用Excel导出功能

1. 在导航菜单中，打开需要导出数据的列表视图
2. 点击 "编辑" → "导出到Excel"
3. 选择保存位置和文件名，点击 "保存"
4. 等待导出完成，查看Excel文件

## 5. 故障排除

### 5.1 模块注册失败

**症状**：应用程序启动时，提示 "无法加载模块"

**解决方案**：
1. 检查Module.cs中是否正确注册了所有模块
2. 确保所有NuGet包版本一致
3. 检查项目引用是否正确

### 5.2 数据字典导入失败

**症状**：从Excel导入数据时，数据字典项没有正确关联

**解决方案**：
1. 确保DataDictionaryItemConverter已正确注册
2. 检查Excel文件中的数据字典名称与系统中是否一致
3. 检查DataDictionaryAttribute是否正确应用

### 5.3 日志中文乱码

**症状**：eXpressAppFramework.log文件中的中文显示为乱码

**解决方案**：
1. 确保配置文件中设置了正确的日志编码：
   ```xml
   <add key="eXpressAppFrameworkTraceEncoding" value="UTF-8"/>
   ```
2. 使用支持UTF-8编码的文本编辑器查看日志文件

### 5.4 依赖项冲突

**症状**：构建时提示依赖项版本冲突

**解决方案**：
1. 确保所有包使用相同版本的DevExpress组件
2. 在app.config中添加绑定重定向：
   ```xml
   <runtime>
       <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
           <dependentAssembly>
               <assemblyIdentity name="DevExpress.ExpressApp" publicKeyToken="b88d1754d700e49a" culture="neutral" />
               <bindingRedirect oldVersion="0.0.0.0-25.1.7.0" newVersion="25.1.7.0" />
           </dependentAssembly>
           <!-- 添加其他DevExpress组件的绑定重定向 -->
       </assemblyBinding>
   </runtime>
   ```

## 6. 最佳实践

### 6.1 模块设计原则

1. **接口解耦**：模块间通过接口通信，避免直接依赖
2. **依赖注入**：使用依赖注入管理模块间的依赖关系
3. **配置驱动**：通过配置文件控制模块行为
4. **日志完善**：提供详细的日志记录，便于调试
5. **异常处理**：完善的异常处理机制，提高系统稳定性

### 6.2 性能优化

1. **批量操作**：Excel导入时使用批量操作，提高性能
2. **索引优化**：为数据字典相关字段添加索引
3. **缓存机制**：对频繁访问的数据字典项进行缓存
4. **延迟加载**：合理使用延迟加载，减少内存占用

### 6.3 安全性

1. **权限控制**：为数据字典管理功能添加权限控制
2. **输入验证**：对Excel导入数据进行严格验证
3. **日志审计**：记录数据字典的增删改操作
4. **防止注入**：使用参数化查询，防止SQL注入

## 7. 版本更新

### 7.1 更新NuGet包

1. 在原始解决方案中修改代码
2. 增加项目文件中的版本号：
   ```xml
   <Version>1.0.1</Version>
   ```
3. 重新构建并生成NuGet包：
   ```
   dotnet build --configuration Release
   dotnet pack --configuration Release
   ```
4. 在新解决方案中，更新NuGet包

### 7.2 版本兼容

- 主版本号：不兼容的API变更
- 次版本号：向下兼容的功能新增
- 修订号：向下兼容的问题修复

## 8. 联系方式

如有任何问题或建议，请联系：

- 邮箱：your.email@example.com
- 电话：123-456-7890
- 文档：https://your-documentation-url.com

## 9. 许可证

本模块采用MIT许可证，详情请查看LICENSE文件。

---

**更新日期**：2025-12-01
**版本**：1.0.0
**作者**：Your Name
