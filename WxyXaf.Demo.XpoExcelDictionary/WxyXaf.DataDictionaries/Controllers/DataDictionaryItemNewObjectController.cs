using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Editors;
using DevExpress.ExpressApp.SystemModule;
using DevExpress.Persistent.Base;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace WxyXaf.DataDictionaries.Controllers;

/// <summary>
/// 处理从下拉列表中新建DataDictionaryItem的控制器
/// 确保新建的DataDictionaryItem关联到正确的DataDictionary
/// </summary>
public class DataDictionaryItemNewObjectController : ViewController
{
    private NewObjectViewController _newObjectController;

    public DataDictionaryItemNewObjectController()
    {
        TargetObjectType = typeof(DataDictionaryItem);
    }

    protected override void OnActivated()
    {
        base.OnActivated();
        
        Tracing.Tracer.LogText("=== DataDictionaryItemNewObjectController.OnActivated 执行 ===");
        
        _newObjectController = Frame.GetController<NewObjectViewController>();
        if (_newObjectController != null)
        {
            Tracing.Tracer.LogText("  找到NewObjectViewController，注册事件处理程序");
            _newObjectController.ObjectCreated += NewObjectController_ObjectCreated;
        }
    }

    private void NewObjectController_ObjectCreated(object sender, ObjectCreatedEventArgs e)
    {
        if (e.CreatedObject is DataDictionaryItem dataDictionaryItem)
        {
            Tracing.Tracer.LogText("=== NewObjectController_ObjectCreated 执行 ===");
            Tracing.Tracer.LogText($"  新建DataDictionaryItem: {dataDictionaryItem.Name}");
            
            // 从事件参数中获取正确的ObjectSpace，而不是从View.ObjectSpace获取
            var objectSpaceForNewItem = e.ObjectSpace;
            Tracing.Tracer.LogText($"  新对象的ObjectSpace: {objectSpaceForNewItem.GetHashCode()}");
            Tracing.Tracer.LogText($"  视图的ObjectSpace: {View.ObjectSpace.GetHashCode()}");
            
            DataDictionary dataDictionary = null;
            
            // 1. 首先检查当前Frame是否是嵌套Frame，且ViewItem是PropertyEditor
            if (Frame is NestedFrame nestedFrame && nestedFrame.ViewItem is PropertyEditor propertyEditor)
            {
                Tracing.Tracer.LogText("  当前Frame是嵌套Frame");
                
                // 获取父对象
                var parentObject = propertyEditor.CurrentObject;
                Tracing.Tracer.LogText($"  父对象类型: {parentObject?.GetType().FullName}");
                
                if (parentObject != null)
                {
                    // 获取PropertyEditor的PropertyName
                    var propertyName = propertyEditor.PropertyName;
                    Tracing.Tracer.LogText($"  PropertyEditor的PropertyName: {propertyName}");
                    
                    // 获取父对象的Type
                    var parentObjectType = parentObject.GetType();
                    
                    // 获取父对象的属性信息
                    var propertyInfo = parentObjectType.GetProperty(propertyName);
                    if (propertyInfo != null)
                    {
                        Tracing.Tracer.LogText($"  找到父对象的属性: {propertyInfo.Name}");
                        
                        // 获取属性上的DataDictionaryAttribute
                        var dataDictionaryAttribute = propertyInfo.GetCustomAttributes(typeof(DataDictionaryAttribute), true).FirstOrDefault() as DataDictionaryAttribute;
                        if (dataDictionaryAttribute != null)
                        {
                            Tracing.Tracer.LogText($"  找到DataDictionaryAttribute，字典名称: {dataDictionaryAttribute.DataDictionaryName}");
                            
                            // 使用属性上的DataDictionaryAttribute的DataDictionaryName查找对应的DataDictionary对象
                            dataDictionary = objectSpaceForNewItem.FindObject<DataDictionary>(
                                DevExpress.Data.Filtering.CriteriaOperator.Parse($"Name = '{dataDictionaryAttribute.DataDictionaryName}'"));
                            if (dataDictionary != null)
                            {
                                Tracing.Tracer.LogText($"  成功找到DataDictionary: {dataDictionary.Name}");
                            }
                            else
                            {
                                Tracing.Tracer.LogText($"  没有找到名称为'{dataDictionaryAttribute.DataDictionaryName}'的DataDictionary");
                            }
                        }
                    }
                    // 如果父对象是DataDictionary，直接使用它
                    else if (parentObject is DataDictionary parentDataDictionary)
                    {
                        Tracing.Tracer.LogText($"  父对象是DataDictionary: {parentDataDictionary.Name}");
                        
                        // 使用父对象的DataDictionary
                        dataDictionary = objectSpaceForNewItem.GetObject<DataDictionary>(parentDataDictionary);
                        Tracing.Tracer.LogText($"  使用父对象的DataDictionary: {dataDictionary.Name}");
                    }
                }
            }
            // 2. 如果不是嵌套视图，尝试从当前视图的CollectionSource中获取过滤条件
            else if (View is ListView listView2 && listView2.CollectionSource != null)
            {
                Tracing.Tracer.LogText("  当前视图是ListView");
                
                // 检查CollectionSource是否有过滤条件
                var criteriaDictionary = listView2.CollectionSource.Criteria;
                if (criteriaDictionary != null && criteriaDictionary.Count > 0)
                {
                    Tracing.Tracer.LogText($"  CollectionSource有{criteriaDictionary.Count}个过滤条件");
                    
                    foreach (var key in criteriaDictionary.Keys)
                    {
                        var criteriaOperator = criteriaDictionary[key];
                        Tracing.Tracer.LogText($"    过滤条件[{key}]: {criteriaOperator}");
                        
                        // 尝试从过滤条件中提取DataDictionary名称
                        string criteriaString = criteriaOperator.ToString();
                        Tracing.Tracer.LogText($"    过滤条件字符串: {criteriaString}");
                        
                        // 使用正则表达式提取DataDictionary名称，支持多种过滤条件格式
                        string dataDictionaryName = ExtractDataDictionaryName(criteriaString);
                        if (!string.IsNullOrEmpty(dataDictionaryName))
                        {
                            Tracing.Tracer.LogText($"    从过滤条件提取到DataDictionary名称: {dataDictionaryName}");
                            
                            // 使用提取到的DataDictionary名称查找对应的DataDictionary对象
                            dataDictionary = objectSpaceForNewItem.FindObject<DataDictionary>(
                                DevExpress.Data.Filtering.CriteriaOperator.Parse($"Name = '{dataDictionaryName}'"));
                            if (dataDictionary != null)
                            {
                                Tracing.Tracer.LogText($"    成功找到DataDictionary: {dataDictionary.Name}");
                            }
                            else
                            {
                                Tracing.Tracer.LogText($"    没有找到名称为'{dataDictionaryName}'的DataDictionary");
                            }
                            break;
                        }
                    }
                }
                else
                {
                    Tracing.Tracer.LogText("  CollectionSource没有过滤条件");
                }
            }
            
            // 3. 如果还是没有获取到DataDictionary，尝试从当前上下文中推断
            if (dataDictionary == null)
            {
                Tracing.Tracer.LogText("  尝试从当前上下文中推断DataDictionary");
                
                // 获取所有DataDictionary
                var allDataDictionaries = objectSpaceForNewItem.GetObjects<DataDictionary>();
                Tracing.Tracer.LogText($"    找到{allDataDictionaries.Count}个DataDictionary");
                
                if (allDataDictionaries.Count == 1)
                {
                    // 如果只有一个DataDictionary，直接使用它
                    dataDictionary = allDataDictionaries[0];
                    Tracing.Tracer.LogText($"    只有一个DataDictionary，直接使用: {dataDictionary.Name}");
                }
            }
            
            // 4. 如果找到了DataDictionary，设置到新建的DataDictionaryItem上
            if (dataDictionary != null)
            {
                // 设置新建的DataDictionaryItem的DataDictionary属性
                dataDictionaryItem.DataDictionary = dataDictionary;
                Tracing.Tracer.LogText($"    成功设置DataDictionary: {dataDictionary.Name}");
            }
            else
            {
                Tracing.Tracer.LogText("    没有找到合适的DataDictionary，将在保存时处理");
            }
            
            Tracing.Tracer.LogText("=== NewObjectController_ObjectCreated 执行完成 ===");
        }
    }
    
    /// <summary>
    /// 使用正则表达式从过滤条件字符串中提取DataDictionary名称
    /// </summary>
    /// <param name="criteriaString">过滤条件字符串</param>
    /// <returns>提取到的DataDictionary名称，失败则返回null</returns>
    private string ExtractDataDictionaryName(string criteriaString)
    {
        try
        {
            // 简化的正则表达式，避免复杂的转义
            string pattern = "DataDictionary\\.Name.*?['\"](.*?)['\"]";
            Match match = Regex.Match(criteriaString, pattern, RegexOptions.IgnoreCase);
            
            if (match.Success && match.Groups.Count > 1)
            {
                return match.Groups[1].Value;
            }
            
            return null;
        }
        catch (Exception ex)
        {
            Tracing.Tracer.LogText($"    正则表达式提取失败: {ex.Message}");
            return null;
        }
    }

    protected override void OnDeactivated()
    {
        if (_newObjectController != null)
        {
            _newObjectController.ObjectCreated -= NewObjectController_ObjectCreated;
            _newObjectController = null;
        }
        base.OnDeactivated();
    }
}