using System;
using System.Reflection;
using WxyXaf.DataDictionaries;

namespace TestDataDictionaryUpdater
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== 测试DataDictionary属性收�?===");
            
            // 加载ExcelXpoExchange.Module程序�?
            Assembly moduleAssembly = Assembly.LoadFrom(@"D:\projects\ExcelXpoExchange\ExcelXpoExchange\ExcelXpoExchange.Module\bin\Debug\net9.0\ExcelXpoExchange.Module.dll");
            Console.WriteLine($"  加载程序�? {moduleAssembly.FullName}");
            
            // 获取Order�?
            Type orderType = moduleAssembly.GetType("ExcelXpoExchange.Module.BusinessObjects.Order");
            if (orderType == null)
            {
                Console.WriteLine("  找不到Order�?);
                return;
            }
            Console.WriteLine($"  找到Order�? {orderType.FullName}");
            
            // 获取Order类的所有属�?
            PropertyInfo[] properties = orderType.GetProperties();
            Console.WriteLine($"  找到 {properties.Length} 个属�?);
            
            // 遍历所有属性，查找带有DataDictionaryAttribute的属�?
            foreach (PropertyInfo property in properties)
            {
                Console.WriteLine($"  处理属�? {property.Name}");
                
                // 获取属性上的DataDictionaryAttribute
                object[] attributes = property.GetCustomAttributes(typeof(DataDictionaryAttribute), true);
                if (attributes.Length > 0)
                {
                    var dataDictionaryAttribute = attributes[0] as DataDictionaryAttribute;
                    if (dataDictionaryAttribute != null)
                    {
                        Console.WriteLine($"  找到DataDictionaryAttribute，字典名�? {dataDictionaryAttribute.DataDictionaryName}");
                    }
                }
                else
                {
                    Console.WriteLine($"  未找到DataDictionaryAttribute");
                }
            }
            
            Console.WriteLine("=== 测试完成 ===");
            Console.ReadKey();
        }
    }
}
