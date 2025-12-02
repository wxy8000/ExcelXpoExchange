using System;
using System.IO;
using System.Reflection;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Xpo;
using DevExpress.Xpo;
using WxyXaf.DataDictionaries;
using WxyXaf.XpoExcel;

namespace TestDataDictionaryImport
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== 测试数据字典Excel导入功能 ===");
            
            try
            {
                // 创建临时Excel文件用于测试
                string excelFilePath = Path.Combine(Path.GetTempPath(), "TestDataDictionary.xlsx");
                Console.WriteLine($"  创建临时Excel文件: {excelFilePath}");
                
                // 创建测试数据
                CreateTestExcelFile(excelFilePath);
                
                // 初始化XAF应用程序
                Console.WriteLine("  初始化XAF应用程序...");
                var application = new BlazorApplication();
                
                // 添加必要的模�?
                application.Modules.Add(new WxyXaf.XpoExcel.WxyXaf.XpoExcel.odule());
                application.Modules.Add(new WxyXaf.DataDictionaries.DataDictionariesModule());
                
                // 设置数据库连接（使用内存数据库）
                application.ConnectionString = @"XpoProvider=SQLite;Data Source=:memory:;Pooling=true;Max Pool Size=100;";
                
                // 初始化应用程�?
                application.Setup();
                application.Start();
                
                // 创建对象空间
                using (var objectSpace = application.CreateObjectSpace(typeof(DataDictionary)) as XPObjectSpace)
                {
                    if (objectSpace == null)
                    {
                        Console.WriteLine("  无法创建对象空间");
                        return;
                    }
                    
                    // 初始化ExcelHelper
                    var excelHelper = new XpoExcelHelper(application, null);
                    
                    // 导入数据字典
                    Console.WriteLine("  导入数据字典...");
                    var result = excelHelper.ImportFromExcel<DataDictionary>(excelFilePath);
                    
                    // 显示导入结果
                    Console.WriteLine($"  数据字典导入结果: 成功={result.SuccessCount}, 失败={result.FailureCount}");
                    if (result.HasErrors)
                    {
                        Console.WriteLine("  导入错误:");
                        foreach (var error in result.Errors)
                        {
                            Console.WriteLine($"    �?{error.RowIndex}, 字段 {error.FieldName}: {error.ErrorMessage}");
                        }
                    }
                    
                    // 导入字典�?
                    Console.WriteLine("  导入字典�?..");
                    result = excelHelper.ImportFromExcel<DataDictionaryItem>(excelFilePath);
                    
                    // 显示导入结果
                    Console.WriteLine($"  字典项导入结�? 成功={result.SuccessCount}, 失败={result.FailureCount}");
                    if (result.HasErrors)
                    {
                        Console.WriteLine("  导入错误:");
                        foreach (var error in result.Errors)
                        {
                            Console.WriteLine($"    �?{error.RowIndex}, 字段 {error.FieldName}: {error.ErrorMessage}");
                        }
                    }
                    
                    // 验证导入结果
                    Console.WriteLine("  验证导入结果...");
                    var dataDictionaries = objectSpace.GetObjects<DataDictionary>();
                    Console.WriteLine($"  导入后数据字典数�? {dataDictionaries.Count}");
                    
                    foreach (var dict in dataDictionaries)
                    {
                        Console.WriteLine($"    数据字典: {dict.Name}, 字典项数�? {dict.Items.Count}");
                        foreach (var item in dict.Items)
                        {
                            Console.WriteLine($"      字典�? {item.Name}, 编码: {item.Code}, 描述: {item.Description}");
                        }
                    }
                }
                
                // 清理
                Console.WriteLine("  清理临时文件...");
                File.Delete(excelFilePath);
                
                Console.WriteLine("=== 测试完成 ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  测试失败: {ex.Message}");
                Console.WriteLine($"  堆栈跟踪: {ex.StackTrace}");
            }
            
            Console.ReadKey();
        }
        
        static void CreateTestExcelFile(string filePath)
        {
            // 使用NPOI创建Excel文件
            using (var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook())
            {
                // 创建数据字典工作�?
                var dataDictSheet = workbook.CreateSheet("DataDictionary");
                
                // 创建表头
                var headerRow = dataDictSheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue("名称");
                
                // 创建测试数据
                var row1 = dataDictSheet.CreateRow(1);
                row1.CreateCell(0).SetCellValue("测试字典1");
                
                var row2 = dataDictSheet.CreateRow(2);
                row2.CreateCell(0).SetCellValue("测试字典2");
                
                // 创建字典项工作表
                var dictItemSheet = workbook.CreateSheet("DataDictionaryItem");
                
                // 创建表头
                headerRow = dictItemSheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue("名称");
                headerRow.CreateCell(1).SetCellValue("编码");
                headerRow.CreateCell(2).SetCellValue("描述");
                headerRow.CreateCell(3).SetCellValue("顺序");
                headerRow.CreateCell(4).SetCellValue("DataDictionary");
                
                // 创建测试数据
                row1 = dictItemSheet.CreateRow(1);
                row1.CreateCell(0).SetCellValue("字典�?");
                row1.CreateCell(1).SetCellValue("ITEM1");
                row1.CreateCell(2).SetCellValue("这是测试字典�?");
                row1.CreateCell(3).SetCellValue(1);
                row1.CreateCell(4).SetCellValue("测试字典1");
                
                row2 = dictItemSheet.CreateRow(2);
                row2.CreateCell(0).SetCellValue("字典�?");
                row2.CreateCell(1).SetCellValue("ITEM2");
                row2.CreateCell(2).SetCellValue("这是测试字典�?");
                row2.CreateCell(3).SetCellValue(2);
                row2.CreateCell(4).SetCellValue("测试字典1");
                
                // 保存Excel文件
                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fileStream);
                }
            }
        }
    }
}
