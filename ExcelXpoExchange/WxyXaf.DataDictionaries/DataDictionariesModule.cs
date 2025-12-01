using DevExpress.ExpressApp;
using DevExpress.ExpressApp.DC;
using DevExpress.ExpressApp.Updating;
using DevExpress.Persistent.Base;
using DevExpress.Xpo;
using System.ComponentModel;
using System.Collections;

namespace WxyXaf.DataDictionaries;

public sealed class DataDictionariesModule : ModuleBase
{
    public DataDictionariesModule()
    {
        RequiredModuleTypes.Add(typeof(DevExpress.ExpressApp.SystemModule.SystemModule));
        RequiredModuleTypes.Add(typeof(WxyXpoExcel.WxyXpoExcelModule));
    }

    public override IEnumerable<ModuleUpdater> GetModuleUpdaters(IObjectSpace objectSpace, Version versionFromDB)
    {
        return new ModuleUpdater[]
        {
            new DataDictionaryUpdater(objectSpace, versionFromDB)
        };
    }

    public override void CustomizeTypesInfo(ITypesInfo typesInfo)
    {
        base.CustomizeTypesInfo(typesInfo);

        Tracing.Tracer.LogText("=== DataDictionariesModule === === DataDictionariesModule.CustomizeTypesInfo 开始执行 ===");

        ITypeInfo dataDictionaryItemTypeInfo = null;

        // typesInfo.PersistentTypes包含持久化类型与非持久化类型，这一块在文章中有介绍
        // https://www.cnblogs.com/haoxj/p/16834882.html
        foreach (var persistentTypeInfo in typesInfo.PersistentTypes.Where(p => p.IsPersistent))
        {
            Tracing.Tracer.LogText($"=== DataDictionariesModule === 处理持久化类型: {persistentTypeInfo.Type.FullName}");
            
            // 遍历所有属性成员，不进行类型过滤
            var members = persistentTypeInfo.Members
                .Where(m => m.IsProperty);

            foreach (var member in members)
            {
                Tracing.Tracer.LogText($"=== DataDictionariesModule ===  处理属性: {member.Name}, 类型: {member.MemberType.Name}");
                
                // 符合数据字典的要求是在属性中添加DataDictionaryAttribute，并且其DataDictionaryName属性不为空
                var attribute = member.FindAttribute<DataDictionaryAttribute>();
                if (attribute != null && !string.IsNullOrWhiteSpace(attribute.DataDictionaryName))
                {
                    Tracing.Tracer.LogText($"=== DataDictionariesModule ===    找到DataDictionaryAttribute，字典名称: {attribute.DataDictionaryName}");
                    
                    // 在TypesInfo中查找出DataDictionaryItem对应的TypeInfo
                    dataDictionaryItemTypeInfo ??= typesInfo.FindTypeInfo(typeof(DataDictionaryItem));

                    // 检查属性类型是否是DataDictionaryItem或其子类，或者是集合类型
                    bool isDataDictionaryItemType = false;
                    Type memberType = member.MemberType;
                    
                    // 检查是否是DataDictionaryItem类型或其子类
                    if (typeof(DataDictionaryItem).IsAssignableFrom(memberType))
                    {
                        isDataDictionaryItemType = true;
                    }
                    // 检查是否是集合类型，且元素类型是DataDictionaryItem或其子类
                    else if (memberType.IsGenericType && typeof(IEnumerable<>).IsAssignableFrom(memberType.GetGenericTypeDefinition()))
                    {
                        var elementType = memberType.GetGenericArguments()[0];
                        if (typeof(DataDictionaryItem).IsAssignableFrom(elementType))
                        {
                            isDataDictionaryItemType = true;
                        }
                    }
                    
                    if (isDataDictionaryItemType)
                    {
                        Tracing.Tracer.LogText($"=== DataDictionariesModule ===    属性类型是DataDictionaryItem或其子类，继续处理");
                        
                        // 在DataDictionaryItem中添加成员，名称为当前成员与其所在类的信息组合，是为了保证唯一
                        // 字典项成员类型为XPCollection<>，它们之间的关系为一对多
                        var dictItemMember = dataDictionaryItemTypeInfo.CreateMember(
                            $"{persistentTypeInfo.Name}_{member.Name}",
                            typeof(XPCollection<>).MakeGenericType(persistentTypeInfo.Type));

                        // 关系名也是一个信息的组合，也是为了保证唯一
                        var associationName = $"{persistentTypeInfo.Name}_{member.Name}_{nameof(DataDictionaryItem)}";
                        Tracing.Tracer.LogText($"=== DataDictionariesModule ===    创建关联: {associationName}");

                        // 向字典项成员中添加AssociationAttribute，注意其中的类型为elementType，不是字典项成员的类型
                        // 操作成员时，都选择了跳过刷新，刷新会放到后面一起执行（可以减少性能的损耗）
                        dictItemMember.AddAttribute(new AssociationAttribute(associationName), true);

                        // 在DataDictionaryItem中隐藏新添加的成员是必要的，因为不需要对其显示及操作
                        dictItemMember.AddAttribute(new BrowsableAttribute(false), true);

                        // 在XPO中AssociationAttribute是双向的，所关联的属性都要有
                        member.AddAttribute(new AssociationAttribute(associationName), true);

                        // 当前成员中如果已有数据来源，则采用已有的，没有则添加一个
                        // 默认数据源的条件为：[DataDictionary.Name]='所对应的字典名'
                        if (member.FindAttribute<DataSourceCriteriaAttribute>() == null
                            && member.FindAttribute<DataSourceCriteriaPropertyAttribute>() == null
                            && member.FindAttribute<DataSourcePropertyAttribute>() == null)
                        {
                            var criteria = $"[DataDictionary.Name]='{attribute.DataDictionaryName}'";
                            Tracing.Tracer.LogText($"=== DataDictionariesModule ===    添加数据源过滤条件: {criteria}");
                            member.AddAttribute(new DataSourceCriteriaAttribute(criteria), true);
                        }

                        // 统一刷新成员
                        ((XafMemberInfo)member).Refresh();
                        ((XafMemberInfo)dictItemMember).Refresh();
                        
                        Tracing.Tracer.LogText($"=== DataDictionariesModule ===    属性处理完成: {member.Name}");
                    }
                    else
                    {
                        Tracing.Tracer.LogText($"=== DataDictionariesModule ===    属性类型不是DataDictionaryItem或其子类，跳过处理");
                    }
                }
            }
        }
        
        Tracing.Tracer.LogText("=== DataDictionariesModule === === DataDictionariesModule.CustomizeTypesInfo 执行完成 ===");
    }
}