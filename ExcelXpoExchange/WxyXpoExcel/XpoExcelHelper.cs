using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Drawing;
using DevExpress.Data.Filtering;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.DC;
using DevExpress.ExpressApp.Model;
using DevExpress.ExpressApp.SystemModule;
using DevExpress.ExpressApp.Xpo;
using DevExpress.Persistent.Base;
using static DevExpress.Persistent.Base.Tracing;
using DevExpress.Xpo;
using DevExpress.Xpo.Metadata;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;


using Wxy.Common;

namespace WxyXpoExcel
{
    
    /// <summary>
    /// 标记XPO类支持Excel导入导出功能的装饰器
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class ExcelImportExportAttribute : Attribute
    {
        /// <summary>
        /// 导出选项
        /// </summary>
        public XpoExcelExportOptions ExportOptions { get; set; } = new XpoExcelExportOptions();
        
        /// <summary>
        /// 导入选项
        /// </summary>
        public XpoExcelImportOptions ImportOptions { get; set; } = new XpoExcelImportOptions();
    }
    
    /// <summary>
    /// 标记XPO属性支持Excel导入导出功能的装饰器
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelFieldAttribute : Attribute
    {
        /// <summary>
        /// 是否包含在导入导出中
        /// </summary>
        public bool Include { get; set; } = true;
        
        /// <summary>
        /// 列标题
        /// </summary>
        public string Caption { get; set; }
        
        /// <summary>
        /// 列顺序
        /// </summary>
        public int Order { get; set; } = 0;
        
        /// <summary>
        /// 导入时是否要求值唯一
        /// </summary>
        public bool IsUnique { get; set; } = false;
    }
    
    public class XpoExcelHelper
    {
        private readonly XafApplication _application;
        private readonly IPlatformService _platformService;
        private readonly List<IRelatedObjectConverter> _relatedObjectConverters;
        
        public XpoExcelHelper(XafApplication application, IPlatformService platformService)
            : this(application, platformService, Enumerable.Empty<IRelatedObjectConverter>())
        {
        }
        
        public XpoExcelHelper(XafApplication application, IPlatformService platformService, IEnumerable<IRelatedObjectConverter> relatedObjectConverters)
        {
            _application = application;
            _platformService = platformService;
            _relatedObjectConverters = new List<IRelatedObjectConverter>();
            
            // 注册所有提供的关联对象转换器
            if (relatedObjectConverters != null)
            {
                foreach (var converter in relatedObjectConverters)
                {
                    RegisterRelatedObjectConverter(converter);
                }
            }
        }
        
        /// <summary>
        /// 注册关联对象转换器
        /// </summary>
        /// <param name="converter">转换器实例</param>
        public void RegisterRelatedObjectConverter(IRelatedObjectConverter converter)
        {
            if (converter != null && !_relatedObjectConverters.Contains(converter))
            {
                _relatedObjectConverters.Add(converter);
            }
        }
        
        /// <summary>
        /// 获取适合指定类型的关联对象转换器
        /// </summary>
        /// <param name="objectType">要转换的对象类型</param>
        /// <returns>适合的转换器，如果没有则返回null</returns>
        private IRelatedObjectConverter GetRelatedObjectConverter(Type objectType)
        {
            return _relatedObjectConverters.FirstOrDefault(converter => converter.CanConvert(objectType));
        }
        
        /// <summary>
        /// 导出XPO对象到Excel文件（非泛型方法）
        /// </summary>
        /// <param name="objectType">XPO对象类型</param>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="criteria">过滤条件</param>
        /// <param name="options">导出选项</param>
        public void ExportToExcel(Type objectType, string filePath, CriteriaOperator criteria = null, XpoExcelExportOptions options = null)
        {
            // 使用反射调用泛型方法，明确指定方法参数类型以避免歧义
            var methods = typeof(XpoExcelHelper).GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.Name == "ExportToExcel" && m.IsGenericMethod)
                .ToList();
            
            var method = methods.FirstOrDefault(m => {
                var parameters = m.GetParameters();
                return parameters.Length == 3 && 
                       parameters[0].ParameterType == typeof(string) &&
                       parameters[1].ParameterType == typeof(CriteriaOperator) &&
                       parameters[2].ParameterType == typeof(XpoExcelExportOptions);
            });
            
            if (method == null)
            {
                throw new InvalidOperationException("无法找到匹配的ExportToExcel方法");
            }
            
            var genericMethod = method.MakeGenericMethod(objectType);
            genericMethod.Invoke(this, new object[] { filePath, criteria, options });
        }
        
        /// <summary>
        /// 导出XPO对象到Excel文件（内部泛型方法）
        /// </summary>
        /// <typeparam name="T">XPO对象类型</typeparam>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="criteria">过滤条件</param>
        /// <param name="options">导出选项</param>
        private void ExportToExcelGeneric<T>(string filePath, CriteriaOperator criteria = null, XpoExcelExportOptions options = null) where T : XPBaseObject
        {
            ExportToExcel<T>(filePath, criteria, options);
        }
        
        /// <summary>
        /// 导出XPO对象到Excel文件
        /// </summary>
        /// <typeparam name="T">XPO对象类型</typeparam>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="criteria">过滤条件</param>
        /// <param name="options">导出选项</param>
        public void ExportToExcel<T>(string filePath, CriteriaOperator criteria = null, XpoExcelExportOptions options = null) where T : XPBaseObject
        {
            // 创建工作簿
            IWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet(typeof(T).Name);
            
            ExportToWorksheet<T>(worksheet, criteria, options);
            
            // 保存文件
            using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }
        }
        
        /// <summary>
        /// 导出XPO对象到Excel内存流
        /// </summary>
        /// <typeparam name="T">XPO对象类型</typeparam>
        /// <param name="criteria">过滤条件</param>
        /// <param name="options">导出选项</param>
        /// <returns>包含Excel数据的内存流</returns>
        public MemoryStream ExportToExcelStream<T>(CriteriaOperator criteria = null, XpoExcelExportOptions options = null) where T : XPBaseObject
        {
            // 创建工作簿
            IWorkbook workbook = new XSSFWorkbook();
            
            try
            {
                ISheet worksheet = workbook.CreateSheet(typeof(T).Name);
                
                ExportToWorksheet<T>(worksheet, criteria, options);
                
                // 创建结果流
                MemoryStream resultStream = new MemoryStream();
                
                // 将工作簿写入结果流
                workbook.Write(resultStream, true); // 第二个参数为true表示不关闭流
                
                // 重置结果流位置
                resultStream.Position = 0;
                
                return resultStream;
            }
            catch
            {
                // 如果出现异常，确保释放workbook资源
                workbook.Close();
                throw;
            }
        }
        
        /// <summary>
        /// 导出XPO对象到Excel工作表
        /// </summary>
        /// <typeparam name="T">XPO对象类型</typeparam>
        /// <param name="worksheet">Excel工作表</param>
        /// <param name="criteria">过滤条件</param>
        /// <param name="options">导出选项</param>
        public void ExportToWorksheet<T>(ISheet worksheet, CriteriaOperator criteria = null, XpoExcelExportOptions options = null) where T : XPBaseObject
        {
            options = options ?? new XpoExcelExportOptions();
            
            using (var objectSpace = _application.CreateObjectSpace(typeof(T)) as XPObjectSpace)
            {
                // 特殊处理DataDictionary类型（使用反射避免直接引用）
                bool isDataDictionaryType = typeof(T).FullName == "WxyXaf.DataDictionaries.DataDictionary";
                if (isDataDictionaryType)
                {
                    // 获取所有DataDictionary对象
                    var objects = objectSpace.GetObjects<T>(criteria);
                    
                    // 创建字典项的表头样式
                    ICellStyle headerStyle = worksheet.Workbook.CreateCellStyle();
                    IFont headerFont = worksheet.Workbook.CreateFont();
                    headerFont.IsBold = true;
                    headerFont.Color = IndexedColors.White.Index;
                    headerStyle.SetFont(headerFont);
                    headerStyle.FillForegroundColor = IndexedColors.Orange.Index;
                    headerStyle.FillPattern = FillPattern.SolidForeground;
                    
                    // 创建字典名称的标题样式
                    ICellStyle titleStyle = worksheet.Workbook.CreateCellStyle();
                    IFont titleFont = worksheet.Workbook.CreateFont();
                    titleFont.IsBold = true;
                    titleFont.FontHeightInPoints = 14;
                    titleFont.Color = IndexedColors.Black.Index;
                    titleStyle.SetFont(titleFont);
                    titleStyle.Alignment = HorizontalAlignment.Center;
                    
                    int currentRowIndex = 0;
                    
                    // 遍历所有数据字典
                    foreach (var dataDictionary in objects)
                    {
                        // 使用反射获取字典名称
                        string dictionaryName = (string)dataDictionary.GetMemberValue("Name");
                        
                        // 1. 写入字典名称作为标题第一行
                        IRow titleRow = worksheet.CreateRow(currentRowIndex);
                        ICell titleCell = titleRow.CreateCell(0);
                        titleCell.SetCellValue(dictionaryName);
                        titleCell.CellStyle = titleStyle;
                        // 合并单元格，使其横跨所有列
                        worksheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(currentRowIndex, currentRowIndex, 0, 3));
                        currentRowIndex++;
                        
                        // 2. 写入字典项的表头
                        IRow headerRow = worksheet.CreateRow(currentRowIndex);
                        headerRow.CreateCell(0).SetCellValue("名称");
                        headerRow.CreateCell(1).SetCellValue("编码");
                        headerRow.CreateCell(2).SetCellValue("描述");
                        headerRow.CreateCell(3).SetCellValue("顺序");
                        
                        // 设置表头样式
                        for (int i = 0; i < 4; i++)
                        {
                            headerRow.Cells[i].CellStyle = headerStyle;
                        }
                        currentRowIndex++;
                        
                        // 3. 使用反射获取字典项数据
                        var items = (System.Collections.IEnumerable)dataDictionary.GetMemberValue("Items");
                        if (items != null)
                        {
                            foreach (var item in items)
                            {
                                // 将item转换为XPBaseObject类型，因为GetMemberValue是XPBaseObject的方法
                                var xpItem = item as XPBaseObject;
                                if (xpItem != null)
                                {
                                    IRow dataRow = worksheet.CreateRow(currentRowIndex);
                                    dataRow.CreateCell(0).SetCellValue((string)xpItem.GetMemberValue("Name"));
                                    dataRow.CreateCell(1).SetCellValue((string)xpItem.GetMemberValue("Code") ?? string.Empty);
                                    dataRow.CreateCell(2).SetCellValue((string)xpItem.GetMemberValue("Description") ?? string.Empty);
                                    dataRow.CreateCell(3).SetCellValue((int)xpItem.GetMemberValue("Order"));
                                    currentRowIndex++;
                                }
                            }
                        }
                        
                        // 每个字典之间空一行
                        currentRowIndex++;
                    }
                    
                    // 自动调整列宽
                    if (options.AutoFitColumns)
                    {
                        for (int i = 0; i < 4; i++)
                        {
                            worksheet.AutoSizeColumn(i);
                        }
                    }
                }
                else
                {
                    // 原有逻辑，处理其他类型
                    var objects = objectSpace.GetObjects<T>(criteria);
                    var typeInfo = _application.Model.BOModel.GetClass(typeof(T));
                    
                    // 获取要导出的成员
                    var members = GetExportMembers(typeInfo, options);
                    
                    // 创建表头样式
                    ICellStyle headerStyle = worksheet.Workbook.CreateCellStyle();
                    IFont headerFont = worksheet.Workbook.CreateFont();
                    headerFont.IsBold = true;
                    headerFont.Color = IndexedColors.White.Index;
                    headerStyle.SetFont(headerFont);
                    headerStyle.FillForegroundColor = IndexedColors.Orange.Index;
                    headerStyle.FillPattern = FillPattern.SolidForeground;
                    
                    // 写入表头
                    IRow headerRow = worksheet.CreateRow(0);
                    for (int i = 0; i < members.Length; i++)
                    {
                        ICell cell = headerRow.CreateCell(i);
                        
                        // 检查是否有ExcelFieldAttribute并使用其Caption
                        var fieldAttribute = members[i].MemberInfo.MemberTypeInfo.Type.GetProperty(members[i].Name)?.GetCustomAttribute<ExcelFieldAttribute>();
                        var caption = fieldAttribute?.Caption ?? members[i].Caption;
                        cell.SetCellValue(caption);
                        cell.CellStyle = headerStyle;
                    }
                    
                    // 写入数据
                    int rowIndex = 1;
                    foreach (var obj in objects)
                    {
                        IRow dataRow = worksheet.CreateRow(rowIndex);
                        for (int i = 0; i < members.Length; i++)
                        {
                            var member = members[i];
                            var value = obj.GetMemberValue(member.Name);
                            
                            // 处理关联对象
                            if (value is XPBaseObject relatedObj)
                            {
                                var displayMember = member.MemberInfo.MemberTypeInfo.DefaultMember;
                                value = relatedObj.GetMemberValue(displayMember?.Name ?? relatedObj.ClassInfo.KeyProperty.Name);
                            }
                            
                            ICell cell = dataRow.CreateCell(i);
                            SetCellValue(cell, value);
                        }
                        rowIndex++;
                    }
                    
                    // 自动调整列宽
                    if (options.AutoFitColumns)
                    {
                        for (int i = 0; i < members.Length; i++)
                        {
                            worksheet.AutoSizeColumn(i);
                        }
                    }
                }
            }
        }
        
        /// <summary>
        /// 设置单元格值，根据数据类型自动转换
        /// </summary>
        private void SetCellValue(ICell cell, object value)
        {
            if (value == null)
            {
                cell.SetCellValue(string.Empty);
                return;
            }

            if (value is string strValue)
            {
                cell.SetCellValue(strValue);
            }
            else if (value is int intValue)
            {
                cell.SetCellValue((double)intValue);
            }
            else if (value is long longValue)
            {
                cell.SetCellValue((double)longValue);
            }
            else if (value is decimal decimalValue)
            {
                cell.SetCellValue((double)decimalValue);
            }
            else if (value is double doubleValue)
            {
                cell.SetCellValue(doubleValue);
            }
            else if (value is float floatValue)
            {
                cell.SetCellValue((double)floatValue);
            }
            else if (value is bool boolValue)
            {
                cell.SetCellValue(boolValue);
            }
            else if (value is DateTime dateTimeValue)
            {
                cell.SetCellValue(dateTimeValue);
            }
            else
            {
                cell.SetCellValue(value.ToString());
            }
        }
        
        /// <summary>
        /// 从Excel文件导入XPO对象
        /// </summary>
        /// <typeparam name="T">XPO对象类型</typeparam>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="options">导入选项</param>
        /// <returns>导入结果</returns>
        public ImportResult ImportFromExcel<T>(string filePath, XpoExcelImportOptions options = null) where T : XPBaseObject
        {
            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromExcel 开始导入文件: {filePath}");
            
            try
            {
                // 检查文件是否存在
                if (!File.Exists(filePath))
                {
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromExcel 文件不存在: {filePath}");
                    return new ImportResult { Errors = { new ImportError { ErrorMessage = $"文件不存在: {filePath}" } } };
                }
                
                IWorkbook workbook;
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromExcel 文件流已打开");
                    
                    // 根据文件扩展名创建对应的工作簿
                    if (Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromExcel 创建XSSF工作簿");
                        workbook = new XSSFWorkbook(fileStream);
                    }
                    else
                    {
                        Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromExcel 创建HSSF工作簿");
                        workbook = new HSSFWorkbook(fileStream);
                    }
                }
                
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromExcel 工作簿创建完成，工作表数量: {workbook.NumberOfSheets}");
                
                if (workbook.NumberOfSheets <= 0)
                {
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromExcel 工作簿中没有工作表");
                    return new ImportResult { Errors = { new ImportError { ErrorMessage = "Excel文件中没有工作表" } } };
                }
                
                var worksheet = workbook.GetSheetAt(0);
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromExcel 获取第一个工作表: {worksheet.SheetName}");
                
                var result = ImportFromWorksheet<T>(worksheet, options);
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromExcel 导入完成，成功: {result.SuccessCount}, 失败: {result.FailureCount}, 错误数: {result.Errors.Count}");
                
                return result;
            }
            catch (Exception ex)
            {
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromExcel 导入失败: {ex.Message}\n{ex.StackTrace}");
                return new ImportResult { Errors = { new ImportError { ErrorMessage = $"导入失败: {ex.Message}" } } };
            }
        }
        
        /// <summary>
        /// 从Excel文件导入XPO对象（非泛型方法）
        /// </summary>
        /// <param name="objectType">XPO对象类型</param>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="options">导入选项</param>
        /// <returns>导入结果</returns>
        public ImportResult ImportFromExcel(Type objectType, string filePath, XpoExcelImportOptions options = null)
        {
            // 使用反射调用泛型方法，明确指定方法参数类型以避免歧义
            var methods = typeof(XpoExcelHelper).GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.Name == "ImportFromExcel" && m.IsGenericMethod)
                .ToList();
            
            var method = methods.FirstOrDefault(m => {
                var parameters = m.GetParameters();
                return parameters.Length == 2 && 
                       parameters[0].ParameterType == typeof(string) &&
                       parameters[1].ParameterType == typeof(XpoExcelImportOptions);
            });
            
            if (method == null)
            {
                throw new InvalidOperationException("无法找到匹配的ImportFromExcel方法");
            }
            
            var genericMethod = method.MakeGenericMethod(objectType);
            return (ImportResult)genericMethod.Invoke(this, new object[] { filePath, options });
        }
        
        /// <summary>
        /// 从Excel文件导入XPO对象（内部泛型方法）
        /// </summary>
        /// <typeparam name="T">XPO对象类型</typeparam>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="options">导入选项</param>
        /// <returns>导入结果</returns>
        private ImportResult ImportFromExcelGeneric<T>(string filePath, XpoExcelImportOptions options = null) where T : XPBaseObject
        {
            return ImportFromExcel<T>(filePath, options);
        }
        
        /// <summary>
        /// 从Excel工作表导入XPO对象
        /// </summary>
        /// <typeparam name="T">XPO对象类型</typeparam>
        /// <param name="worksheet">Excel工作表</param>
        /// <param name="options">导入选项</param>
        /// <returns>导入结果</returns>
        public ImportResult ImportFromWorksheet<T>(ISheet worksheet, XpoExcelImportOptions options = null) where T : XPBaseObject
        {
            options = options ?? new XpoExcelImportOptions();
            var result = new ImportResult();
            
            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 开始导入类型: {typeof(T).Name}");
            
            try
            {
                using (var objectSpace = _application.CreateObjectSpace(typeof(T)) as XPObjectSpace)
                {
                    if (objectSpace == null)
                    {
                        Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 无法创建对象空间");
                        result.Errors.Add(new ImportError { ErrorMessage = "无法创建对象空间" });
                        return result;
                    }
                    
                    var typeInfo = _application.Model.BOModel.GetClass(typeof(T));
                    
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 获取类型信息: {typeInfo?.Name ?? "null"}");
                    
                    if (typeInfo == null)
                    {
                        Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 无法获取类型信息");
                        result.Errors.Add(new ImportError { ErrorMessage = "无法获取类型信息" });
                        return result;
                    }
                    
                    // 读取表头，建立字段映射
                    var headerRow = worksheet.GetRow(0);
                    if (headerRow == null)
                    {
                        Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 表头为空");
                        result.Errors.Add(new ImportError { ErrorMessage = "Excel文件表头为空" });
                        return result;
                    }
                    
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 表头行单元格数: {headerRow.LastCellNum}");
                    
                    var fieldMappings = new Dictionary<int, IModelMember>();
                    for (int c = 0; c < headerRow.LastCellNum; c++)
                    {
                        var cell = headerRow.GetCell(c);
                        if (cell != null)
                        {
                            var fieldCaption = cell.StringCellValue;
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 处理表头单元格 {c}: {fieldCaption}");
                            
                            var member = typeInfo.AllMembers.SingleOrDefault(x => x.Caption == fieldCaption);
                            if (member != null)
                            {
                                fieldMappings.Add(c, member);
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 字段映射: {fieldCaption} -> {member.Name}");
                            }
                            else
                            {
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 未找到字段映射: {fieldCaption}");
                            }
                        }
                        else
                        {
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 表头单元格 {c} 为空");
                        }
                    }
                    
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 字段映射完成，共 {fieldMappings.Count} 个字段");
                    
                    // 如果没有找到任何字段映射，返回错误
                    if (fieldMappings.Count == 0)
                    {
                        Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 没有找到任何字段映射");
                        result.Errors.Add(new ImportError { ErrorMessage = "Excel文件表头与对象字段不匹配，请检查列标题是否正确" });
                        return result;
                    }
                    
                    // 读取数据行
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 数据行范围: 1 - {worksheet.LastRowNum}");
                    
                    int processedRowCount = 0;
                    for (int r = 1; r <= worksheet.LastRowNum; r++)
                    {
                        var dataRow = worksheet.GetRow(r);
                        if (dataRow == null)
                        {
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 第 {r} 行为空，跳过");
                            continue;
                        }
                        
                        processedRowCount++;
                        Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 处理第 {r} 行数据，单元格数: {dataRow.LastCellNum}");
                        
                        T obj = null;
                        bool isNewObject = false;
                        
                        // 根据导入模式处理对象创建/查找
                        switch (options.Mode)
                        {
                            case ImportMode.CreateOnly:
                                // 仅新建
                                obj = objectSpace.CreateObject<T>();
                                isNewObject = true;
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 创建新对象(仅新建模式)");
                                break;
                                
                            case ImportMode.UpdateOnly:
                                // 仅更新现有记录
                                obj = FindExistingObject<T>(worksheet, r, fieldMappings, objectSpace, options);
                                if (obj == null)
                                {
                                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 未找到现有对象，跳过更新");
                                    result.FailureCount++;
                                    continue;
                                }
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 找到现有对象进行更新");
                                break;
                                
                            case ImportMode.CreateAndUpdate:
                                // 新建和更新（默认模式）
                                obj = FindExistingObject<T>(worksheet, r, fieldMappings, objectSpace, options) ?? 
                                      objectSpace.CreateObject<T>();
                                isNewObject = (obj != null && objectSpace.IsNewObject(obj));
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet {(isNewObject ? "创建新对象" : "更新现有对象")}(新建和更新模式)");
                                break;
                                
                            case ImportMode.DeleteAndUpdate:
                                // 删除后更新
                                var existingObj = FindExistingObject<T>(worksheet, r, fieldMappings, objectSpace, options);
                                if (existingObj != null)
                                {
                                    // 删除现有对象
                                    objectSpace.Delete(existingObj);
                                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 删除现有对象");
                                }
                                // 创建新对象
                                obj = objectSpace.CreateObject<T>();
                                isNewObject = true;
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 删除后创建新对象");
                                break;
                        }
                        
                        if (obj == null)
                        {
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 对象创建失败");
                            result.FailureCount++;
                            result.Errors.Add(new ImportError
                            {
                                RowIndex = r,
                                FieldName = "ObjectCreation",
                                ErrorMessage = "无法创建对象实例"
                            });
                            continue;
                        }
                        
                        bool rowHasError = false;
                        int filledFieldCount = 0;
                        
                        // 填充字段值
                        foreach (var mapping in fieldMappings)
                        {
                            var columnIndex = mapping.Key;
                            var member = mapping.Value;
                            var cell = dataRow.GetCell(columnIndex);
                            
                            if (cell != null)
                            {
                                try
                                {
                                    var value = ConvertCellValue(cell, member, objectSpace);
                                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 设置字段值: {member.Name} = {value?.ToString() ?? "null"}");
                                    
                                    // 检查字段是否要求唯一
                                    var propertyInfo = member.MemberInfo.MemberTypeInfo.Type.GetProperty(member.Name);
                                    var fieldAttribute = propertyInfo?.GetCustomAttribute<ExcelFieldAttribute>();
                                    if (fieldAttribute != null && fieldAttribute.IsUnique)
                                    {
                                        // 构建唯一性查询条件
                                        CriteriaOperator criteria = new BinaryOperator(member.Name, value);
                                        // 如果是更新模式，排除当前对象
                                        if (options.Mode == ImportMode.UpdateOnly || options.Mode == ImportMode.CreateAndUpdate)
                                        {
                                            // 获取对象的主键信息（使用XPO ClassInfo）
                                            var classInfo = objectSpace.Session.GetClassInfo(obj.GetType());
                                            var keyProperty = classInfo.KeyProperty;
                                            if (keyProperty != null)
                                            {
                                                var keyValue = obj.GetMemberValue(keyProperty.Name);
                                                criteria = GroupOperator.And(
                                                    criteria,
                                                    new BinaryOperator(keyProperty.Name, keyValue, BinaryOperatorType.NotEqual)
                                                );
                                            }
                                        }
                                        
                                        // 检查是否已存在相同值
                                        var existingObject = objectSpace.FindObject(obj.GetType(), criteria);
                                        if (existingObject != null)
                                        {
                                            throw new InvalidOperationException($"字段 '{member.Name}' 的值 '{value}' 已存在，要求唯一");
                                        }
                                    }
                                    
                                    obj.SetMemberValue(member.Name, value);
                                    filledFieldCount++;
                                }
                                catch (Exception ex)
                                {
                                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 设置字段值失败: {member.Name}, 错误: {ex.Message}");
                                    result.Errors.Add(new ImportError
                                    {
                                        RowIndex = r,
                                        FieldName = member.Name,
                                        ErrorMessage = ex.Message
                                    });
                                    rowHasError = true;
                                }
                            }
                            else
                            {
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 单元格 {columnIndex} 为空，跳过字段: {member.Name}");
                            }
                        }
                        
                        Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 第 {r} 行填充了 {filledFieldCount} 个字段");
                        
                        if (!rowHasError)
                        {
                            result.SuccessCount++;
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 第 {r} 行处理成功");
                        }
                        else
                        {
                            result.FailureCount++;
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 第 {r} 行处理失败");
                            if (options.StopOnError)
                            {
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 遇到错误停止导入");
                                break;
                            }
                        }
                    }
                    
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 数据处理完成，处理行数: {processedRowCount}，成功: {result.SuccessCount}, 失败: {result.FailureCount}");
                    
                    // 提交事务 - 即使有错误也要提交成功的记录
                    if (result.SuccessCount > 0)
                    {
                        try
                        {
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 开始提交事务");
                            objectSpace.CommitChanges();
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 事务提交成功");
                        }
                        catch (Exception commitEx)
                        {
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 事务提交失败: {commitEx.Message}");
                            // 如果提交失败，记录错误并回滚
                            result.Errors.Add(new ImportError
                            {
                                RowIndex = -1,
                                FieldName = "Commit",
                                ErrorMessage = $"提交事务失败: {commitEx.Message}"
                            });
                            objectSpace.Rollback();
                        }
                    }
                    else
                    {
                        Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 没有成功记录，跳过事务提交");
                    }
                }
            }
            catch (Exception ex)
            {
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 处理过程中发生异常: {ex.Message}\n{ex.StackTrace}");
                result.Errors.Add(new ImportError
                {
                    RowIndex = -1,
                    FieldName = "General",
                    ErrorMessage = $"处理过程中发生异常: {ex.Message}"
                });
            }
            
            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ImportFromWorksheet 导入完成，成功: {result.SuccessCount}, 失败: {result.FailureCount}, 错误数: {result.Errors.Count}");
            
            return result;
        }
        
        /// <summary>
        /// 查找现有对象
        /// </summary>
        private T FindExistingObject<T>(ISheet worksheet, int rowIndex, Dictionary<int, IModelMember> fieldMappings, XPObjectSpace objectSpace, XpoExcelImportOptions options) where T : XPBaseObject
        {
            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 开始查找现有对象，行索引: {rowIndex}");
            
            if (!string.IsNullOrEmpty(options.KeyMember))
            {
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 使用指定关键字段: {options.KeyMember}");
                
                // 使用指定的关键字段查找
                var keyMapping = fieldMappings.FirstOrDefault(x => x.Value.Name == options.KeyMember);
                if (keyMapping.Value != null)
                {
                    var dataRow = worksheet.GetRow(rowIndex);
                    if (dataRow != null)
                    {
                        var cell = dataRow.GetCell(keyMapping.Key);
                        if (cell != null)
                        {
                            var keyValue = ConvertCellValue(cell, keyMapping.Value, objectSpace);
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 关键字段值: {keyValue}");
                            
                            var criteria = new BinaryOperator(options.KeyMember, keyValue?.ToString() ?? string.Empty);
                            var result = objectSpace.FindObject<T>(criteria);
                            
                            if (result != null)
                            {
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 找到现有对象");
                            }
                            else
                            {
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 未找到现有对象");
                            }
                            
                            return result;
                        }
                        else
                        {
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 关键字段单元格为空");
                        }
                    }
                    else
                    {
                        Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 数据行为空");
                    }
                }
                else
                {
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 未找到关键字段映射");
                }
            }
            
            // 使用主键查找 - 简化实现，直接使用第一个字段
            // 实际应用中可能需要更复杂的逻辑来确定主键
            var typeInfo = _application.Model.BOModel.GetClass(typeof(T));
            // 优先使用对象的主键字段
            IModelMember keyMember = null;
            if (!string.IsNullOrEmpty(options.KeyMember))
            {
                keyMember = typeInfo.AllMembers.FirstOrDefault(m => m.Name == options.KeyMember);
            }
            
            if (keyMember == null)
            {
                // 尝试使用类型的主键属性
                // 注意：KeyProperty可能直接返回属性名称字符串而不是对象
                if (typeInfo.KeyProperty != null)
                {
                    // 检查KeyProperty是否为字符串类型
                    var keyPropertyName = typeInfo.KeyProperty.ToString();
                    if (!string.IsNullOrWhiteSpace(keyPropertyName))
                    {
                        keyMember = typeInfo.AllMembers.FirstOrDefault(m => m.Name == keyPropertyName);
                    }
                }
                
                // 如果还是没有找到，使用第一个成员
                if (keyMember == null)
                {
                    keyMember = typeInfo.AllMembers.FirstOrDefault();
                }
            }
                            
            if (keyMember != null)
            {
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 使用默认关键字段: {keyMember.Name}");
                
                var keyMapping = fieldMappings.FirstOrDefault(x => x.Value.Name == keyMember.Name);
                if (keyMapping.Value != null)
                {
                    var dataRow = worksheet.GetRow(rowIndex);
                    if (dataRow != null)
                    {
                        var cell = dataRow.GetCell(keyMapping.Key);
                        if (cell != null)
                        {
                            var keyValue = ConvertCellValue(cell, keyMapping.Value, objectSpace);
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 关键字段值: {keyValue}");
                            
                            var criteria = new BinaryOperator(keyMember.Name, keyValue?.ToString() ?? string.Empty);
                            var result = objectSpace.FindObject<T>(criteria);
                            
                            if (result != null)
                            {
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 找到现有对象");
                            }
                            else
                            {
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 未找到现有对象");
                            }
                            
                            return result;
                        }
                        else
                        {
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 关键字段单元格为空");
                        }
                    }
                    else
                    {
                        Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 数据行为空");
                    }
                }
                else
                {
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 未找到关键字段映射");
                }
            }
            else
            {
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 未找到类型关键字段");
            }
            
            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] FindExistingObject 返回null");
            return null;
        }
        
        /// <summary>
        /// 转换单元格值到指定类型
        /// </summary>
        private object ConvertCellValue(ICell cell, IModelMember member, XPObjectSpace objectSpace)
        {
            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 开始转换单元格值: 列={cell.ColumnIndex}, 行={cell.RowIndex}, 字段={member.Name}");
            
            var memberType = member.MemberInfo.MemberType;
            
            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 成员类型: {memberType.FullName}");
            
            // 处理可空类型
            if (memberType.IsValueType && memberType.IsGenericType && 
                memberType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                memberType = memberType.GetGenericArguments()[0];
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 处理可空类型，基础类型: {memberType.FullName}");
            }
            
            // 获取单元格值
            object cellValue = null;
            
            switch (cell.CellType)
            {
                case CellType.String:
                    cellValue = cell.StringCellValue;
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 字符串类型值: {cellValue}");
                    break;
                case CellType.Numeric:
                    // 对于数值类型，需要特别处理
                    // Excel将所有数值存储为从1900年1月1日开始的天数偏移量
                    // 需要判断是否真的是日期还是普通数值
                    
                    // 首先尝试判断是否为真正的日期格式
                    try
                    {
                        var dateValue = cell.DateCellValue;
                        // 检查是否为有效的日期（不是从1900年开始的简单数值）
                        if (dateValue.HasValue && dateValue.Value != DateTime.MinValue && dateValue.Value != DateTime.MaxValue && 
                            dateValue.Value.Year > 1900) // 如果年份大于1900，很可能是真正的日期
                        {
                            cellValue = dateValue.Value;
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 日期类型值: {cellValue}");
                        }
                        else
                        {
                            // 不是有效的日期，作为数字处理
                            cellValue = cell.NumericCellValue;
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 数字类型值: {cellValue}");
                        }
                    }
                    catch
                    {
                        // 转换失败，作为数字处理
                        cellValue = cell.NumericCellValue;
                        Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 数字类型值(异常处理): {cellValue}");
                    }
                    break;
                case CellType.Boolean:
                    cellValue = cell.BooleanCellValue;
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 布尔类型值: {cellValue}");
                    break;
                case CellType.Formula:
                    // 计算公式结果
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.String:
                            cellValue = cell.StringCellValue;
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 公式字符串类型值: {cellValue}");
                            break;
                        case CellType.Numeric:
                            // 对于公式计算的数值，同样需要判断是否为日期
                            try
                            {
                                var dateValue = cell.DateCellValue;
                                // 检查是否为有效的日期
                                if (dateValue.HasValue && dateValue.Value != DateTime.MinValue && dateValue.Value != DateTime.MaxValue && 
                                    dateValue.Value.Year > 1900)
                                {
                                    cellValue = dateValue.Value;
                                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 公式日期类型值: {cellValue}");
                                }
                                else
                                {
                                    // 不是有效的日期，作为数字处理
                                    cellValue = cell.NumericCellValue;
                                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 公式数字类型值: {cellValue}");
                                }
                            }
                            catch
                            {
                                // 转换失败，作为数字处理
                                cellValue = cell.NumericCellValue;
                                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 公式数字类型值(异常处理): {cellValue}");
                            }
                            break;
                        case CellType.Boolean:
                            cellValue = cell.BooleanCellValue;
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 公式布尔类型值: {cellValue}");
                            break;
                        default:
                            cellValue = cell.StringCellValue;
                            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 公式默认类型值: {cellValue}");
                            break;
                    }
                    break;
                default:
                    cellValue = string.Empty;
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 默认类型值: {cellValue}");
                    break;
            }
            
            // 处理关联对象
            if (typeof(XPBaseObject).IsAssignableFrom(memberType))
            {
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 处理关联对象类型: {memberType.Name}");
                
                // 如果单元格值为空，返回null
                if (cellValue == null || (cellValue is string str && string.IsNullOrWhiteSpace(str)))
                {
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 关联对象单元格值为空，返回null");
                    return null;
                }
                
                // 查找适合当前类型的关联对象转换器
                var converter = GetRelatedObjectConverter(memberType);
                if (converter != null)
                {
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 找到适合的关联对象转换器: {converter.GetType().Name}");
                    // 使用转换器转换值
                    return converter.Convert(cellValue, member, objectSpace);
                }
                else
                {
                    // 没有找到适合的转换器，返回null
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 没有找到适合的关联对象转换器，返回null");
                    return null;
                }
            }
            
            // 转换基本类型
            if (memberType == typeof(DateTime))
            {
                // 处理空字符串情况
                if (cellValue == null || (cellValue is string str && string.IsNullOrWhiteSpace(str)))
                {
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 空字符串转换为DateTime，返回DateTime.MinValue");
                    return DateTime.MinValue;
                }
                
                var result = Convert.ToDateTime(cellValue);
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 转换为DateTime: {result}");
                return result;
            }
            else if (memberType == typeof(bool))
            {
                var result = Convert.ToBoolean(cellValue);
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 转换为Boolean: {result}");
                return result;
            }
            else if (memberType.IsEnum)
            {
                if (cellValue is double numericValue)
                {
                    var result = Enum.ToObject(memberType, numericValue);
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 转换为枚举(数字): {result}");
                    return result;
                }
                else
                {
                    var result = Enum.Parse(memberType, cellValue.ToString());
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 转换为枚举(字符串): {result}");
                    return result;
                }
            }
            else if (memberType == typeof(decimal) || memberType == typeof(double) || 
                     memberType == typeof(float) || memberType == typeof(int) || 
                     memberType == typeof(long))
            {
                // 特别处理数值类型
                if (cellValue is DateTime dateTimeValue)
                {
                    // 如果是从Excel读取的日期格式数值，转换为对应的数字
                    // Excel日期是从1900年1月1日开始计算的天数
                    var excelEpoch = new DateTime(1900, 1, 1);
                    var numericValue = (dateTimeValue - excelEpoch).TotalDays;
                    // Excel有一个特殊的错误，认为1900年是闰年，所以需要加1天（除了1900年1月1日本身）
                    if (numericValue > 1)
                        numericValue += 1;
                    
                    var result = Convert.ChangeType(numericValue, memberType);
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 转换Excel日期数值为{memberType.Name}: {result}");
                    return result;
                }
                else
                {
                    var result = Convert.ChangeType(cellValue, memberType);
                    Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 转换为数值类型: {result}");
                    return result;
                }
            }
            else if (memberType.IsValueType || memberType == typeof(string))
            {
                var result = Convert.ChangeType(cellValue, memberType);
                Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 转换为基本类型: {result}");
                return result;
            }
            
            Tracing.Tracer.LogText($"[{DateTime.Now:HH:mm:ss.fff}] ConvertCellValue 返回原始值: {cellValue}");
            return cellValue;
        }
        
        /// <summary>
        /// 获取要导出的成员
        /// </summary>
        private IModelMember[] GetExportMembers(IModelClass typeInfo, XpoExcelExportOptions options)
        {
            var members = typeInfo.AllMembers.Where(x =>
                !x.MemberInfo.IsAutoGenerate &&
                !x.IsCalculated &&
                !x.MemberInfo.IsList &&
                !x.MemberInfo.IsReadOnly
            );
            
            // 检查ExcelFieldAttribute
            var membersWithAttributes = members.Select(member =>
            {
                var fieldAttribute = member.MemberInfo.MemberTypeInfo.Type.GetProperty(member.Name)?.GetCustomAttribute<ExcelFieldAttribute>();
                return new { Member = member, Attribute = fieldAttribute };
            });
            
            // 过滤掉标记为不包含的成员
            membersWithAttributes = membersWithAttributes.Where(x => x.Attribute == null || x.Attribute.Include);
            
            // 应用排除列表
            if (options.ExcludedMembers != null && options.ExcludedMembers.Any())
            {
                membersWithAttributes = membersWithAttributes.Where(x => !options.ExcludedMembers.Contains(x.Member.Name));
            }
            
            // 应用包含列表
            if (options.IncludedMembers != null && options.IncludedMembers.Any())
            {
                membersWithAttributes = membersWithAttributes.Where(x => options.IncludedMembers.Contains(x.Member.Name));
            }
            
            // 根据ExcelFieldAttribute.Order排序
            var sortedMembers = membersWithAttributes.OrderBy(x => x.Attribute?.Order ?? 0)
                .Select(x => x.Member)
                .ToArray();
            
            return sortedMembers;
        }
    }
    
    /// <summary>
    /// 导入结果
    /// </summary>
    public class ImportResult
    {
        /// <summary>
        /// 成功导入的记录数
        /// </summary>
        public int SuccessCount { get; set; }
        
        /// <summary>
        /// 失败的记录数
        /// </summary>
        public int FailureCount { get; set; }
        
        /// <summary>
        /// 导入过程中发生的错误
        /// </summary>
        public List<ImportError> Errors { get; set; } = new List<ImportError>();
        
        /// <summary>
        /// 是否包含错误
        /// </summary>
        public bool HasErrors => Errors != null && Errors.Count > 0;
    }
    
    /// <summary>
    /// 导入错误
    /// </summary>
    public class ImportError
    {
        /// <summary>
        /// 错误行索引
        /// </summary>
        public int RowIndex { get; set; }
        
        /// <summary>
        /// 错误字段名称
        /// </summary>
        public string FieldName { get; set; }
        
        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }
    
    /// <summary>
    /// 导入模式
    /// </summary>
    public enum ImportMode
    {
        /// <summary>
        /// 仅创建新记录
        /// </summary>
        CreateOnly,
        
        /// <summary>
        /// 仅更新现有记录
        /// </summary>
        UpdateOnly,
        
        /// <summary>
        /// 创建新记录并更新现有记录
        /// </summary>
        CreateAndUpdate,
        
        /// <summary>
        /// 删除现有记录并创建新记录
        /// </summary>
        DeleteAndUpdate
    }
    
    /// <summary>
    /// 导入选项
    /// </summary>
    public class XpoExcelImportOptions
    {
        /// <summary>
        /// 导入模式
        /// </summary>
        public ImportMode Mode { get; set; } = ImportMode.CreateAndUpdate;
        
        /// <summary>
        /// 关键字段名称
        /// </summary>
        public string KeyMember { get; set; }
        
        /// <summary>
        /// 是否遇到错误时停止导入
        /// </summary>
        public bool StopOnError { get; set; } = false;
    }
    
    /// <summary>
    /// 导出选项
    /// </summary>
    public class XpoExcelExportOptions
    {
        /// <summary>
        /// 是否自动调整列宽
        /// </summary>
        public bool AutoFitColumns { get; set; } = true;
        
        /// <summary>
        /// 要排除的成员名称列表
        /// </summary>
        public List<string> ExcludedMembers { get; set; }
        
        /// <summary>
        /// 要包含的成员名称列表
        /// </summary>
        public List<string> IncludedMembers { get; set; }
    }
}