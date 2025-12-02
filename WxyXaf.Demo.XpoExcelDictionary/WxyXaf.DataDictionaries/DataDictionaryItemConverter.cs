using System;
using DevExpress.Data.Filtering;
using DevExpress.ExpressApp.Model;
using DevExpress.ExpressApp.Xpo;
using WxyXaf.Common;

namespace WxyXaf.DataDictionaries
{
    /// <summary>
    /// DataDictionaryItem类型的关联对象转换器
    /// </summary>
    public class DataDictionaryItemConverter : IRelatedObjectConverter
    {
        /// <summary>
        /// 判断当前转换器是否支持指定类�?
        /// </summary>
        /// <param name="objectType">要转换的对象类型</param>
        /// <returns>是否支持</returns>
        public bool CanConvert(Type objectType)
        {
            return objectType == typeof(DataDictionaryItem);
        }
        
        /// <summary>
        /// 将Excel单元格值转换为DataDictionaryItem对象
        /// </summary>
        /// <param name="cellValue">Excel单元格�?/param>
        /// <param name="member">模型成员信息</param>
        /// <param name="objectSpace">对象空间</param>
        /// <returns>转换后的DataDictionaryItem对象</returns>
        public object Convert(object cellValue, IModelMember member, XPObjectSpace objectSpace)
        {
            if (cellValue == null)
            {
                return null;
            }
            
            // 获取显示�?
            string displayValue = cellValue.ToString();
            if (string.IsNullOrWhiteSpace(displayValue))
            {
                return null;
            }
            
            // 查找或创建DataDictionaryItem
            // 1. 首先获取DataDictionaryName（从DataDictionaryAttribute�?
            DataDictionaryAttribute dataDictionaryAttribute = null;
            string dictionaryName = string.Empty;
            
            // 方式1：使用反射直接从属性获取DataDictionaryAttribute
            // 注意：member.MemberInfo.Owner.Type 是包含该属性的类（如Order），而不是属性类型（如DataDictionaryItem�?
            var propertyInfo = member.MemberInfo.Owner.Type.GetProperty(member.Name);
            if (propertyInfo != null)
            {
                dataDictionaryAttribute = propertyInfo.GetCustomAttributes(typeof(DataDictionaryAttribute), true)
                    .FirstOrDefault() as DataDictionaryAttribute;
                
                if (dataDictionaryAttribute != null)
                {
                    dictionaryName = dataDictionaryAttribute.DataDictionaryName;
                }
            }
            
            // 方式2：如果方�?失败，尝试使用FindAttribute方法获取
            if (dataDictionaryAttribute == null)
            {
                // 使用FindAttribute方法查找DataDictionaryAttribute
                dataDictionaryAttribute = member.MemberInfo.FindAttribute<DataDictionaryAttribute>();
                if (dataDictionaryAttribute != null)
                {
                    dictionaryName = dataDictionaryAttribute.DataDictionaryName;
                }
            }
            
            // 方式3：如果方�?和方�?都失败，尝试从MemberTypeInfo获取
            if (dataDictionaryAttribute == null)
            {
                // 从MemberTypeInfo获取属性信�?
                var typeInfo = member.MemberInfo.MemberTypeInfo;
                if (typeInfo != null)
                {
                    var memberInfo = typeInfo.FindMember(member.Name);
                    if (memberInfo != null)
                    {
                        dataDictionaryAttribute = memberInfo.FindAttribute<DataDictionaryAttribute>();
                        if (dataDictionaryAttribute != null)
                        {
                            dictionaryName = dataDictionaryAttribute.DataDictionaryName;
                        }
                    }
                }
            }
            
            if (!string.IsNullOrWhiteSpace(dictionaryName))
            {
                // 2. 查找或创建对应的DataDictionary
                var dataDictionary = objectSpace.FindObject<DataDictionary>(
                    new BinaryOperator("Name", dictionaryName));
                
                if (dataDictionary == null)
                {
                    // 如果DataDictionary不存在，创建�?
                    dataDictionary = objectSpace.CreateObject<DataDictionary>();
                    dataDictionary.Name = dictionaryName;
                    // 确保DataDictionary被保存，以便后续查询能找到它
                    objectSpace.CommitChanges();
                }
                
                // 3. 重新查询DataDictionary，确保它在对象空间中
                dataDictionary = objectSpace.FindObject<DataDictionary>(
                    new BinaryOperator("Name", dictionaryName));
                
                if (dataDictionary == null)
                {
                    // 如果仍然找不到，可能是创建失败，返回null
                    return null;
                }
                
                // 4. 查找或创建对应的DataDictionaryItem
                var dictionaryItem = objectSpace.FindObject<DataDictionaryItem>(
                    GroupOperator.And(
                        new BinaryOperator("Name", displayValue),
                        new BinaryOperator("DataDictionary.Oid", dataDictionary.Oid)));
                
                if (dictionaryItem != null)
                {
                    return dictionaryItem;
                }
                else
                {
                    // 5. 如果找不到，创建新的DataDictionaryItem
                    var newItem = objectSpace.CreateObject<DataDictionaryItem>();
                    newItem.Name = displayValue;
                    newItem.DataDictionary = dataDictionary;
                    // 双向关联，确保关系正确建�?
                    dataDictionary.Items.Add(newItem);
                    return newItem;
                }
            }
            
            return null;
        }
    }
}
