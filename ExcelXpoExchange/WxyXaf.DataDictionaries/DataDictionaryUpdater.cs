using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Updating;
using DevExpress.Xpo;

using DevExpress.Persistent.Base;

namespace WxyXaf.DataDictionaries;

public class DataDictionaryUpdater : ModuleUpdater
{
    public DataDictionaryUpdater(IObjectSpace objectSpace, Version currentDBVersion)
        : base(objectSpace, currentDBVersion)
    {
    }

    public override void UpdateDatabaseAfterUpdateSchema()
    {
        base.UpdateDatabaseAfterUpdateSchema();

        Tracing.Tracer.LogText("=== DataDictionaryUpdater.UpdateDatabaseAfterUpdateSchema 开始执行 ===");

        // 收集XAF中的所有数据字典名称，并创建对应的数据字典
        // 收集思路与Module类中CustomizeTypesInfo方法类似，请参考其中的注释
        // 注意：这里并没有对其进行过滤，只要有DataDictionaryAttribute都会被收集

        var typesInfo = ObjectSpace.TypesInfo;
        var dataDictionaryNames = new List<string>();

        Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===  找到 {typesInfo.PersistentTypes.Count()} 个持久类型");
        
        // 遍历所有持久类型
        foreach (var persistentTypeInfo in typesInfo.PersistentTypes.Where(p => p.IsPersistent))
        {
            // 获取实际的CLR类型
            Type clrType = persistentTypeInfo.Type;
            Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===  处理CLR类型: {clrType.FullName}");
            
            // 特别检查Order类
            if (clrType.Name == "Order")
            {
                Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===  找到Order类，开始处理其属性");
            }
            
            // 遍历类型的所有属性成员
            var properties = persistentTypeInfo.Members.Where(m => m.IsProperty).ToList();
            Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    找到 {properties.Count} 个属性");
            
            foreach (var member in properties)
            {
                Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    处理属性: {member.Name}");
                
                // 特别检查Department和ZhiWei属性
                if (member.Name == "Department" || member.Name == "ZhiWei")
                {
                    Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    找到目标属性: {member.Name}");
                }
                
                // 尝试获取DataDictionaryAttribute
                DataDictionaryAttribute attribute = null;
                
                // 方式1: 直接从CLR属性获取Attribute
                var propertyInfo = clrType.GetProperty(member.Name);
                if (propertyInfo != null)
                {
                    Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    方式2: 尝试从CLR属性获取Attribute");
                    var attributes = propertyInfo.GetCustomAttributes(typeof(DataDictionaryAttribute), true);
                    Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    方式2: 找到 {attributes.Length} 个DataDictionaryAttribute");
                    attribute = attributes.FirstOrDefault() as DataDictionaryAttribute;
                    if (attribute != null)
                    {
                        Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    方式2: 从CLR属性获取到DataDictionaryAttribute，字典名称: {attribute.DataDictionaryName}");
                    }
                    else
                    {
                        Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    方式2: 未从CLR属性获取到DataDictionaryAttribute");
                    }
                }
                else
                {
                    Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    方式2: 无法获取CLR属性信息");
                }
                
                if (attribute != null && !string.IsNullOrWhiteSpace(attribute.DataDictionaryName))
                {
                    // 如果找到有效的DataDictionaryAttribute，添加到列表中
                    if (!dataDictionaryNames.Contains(attribute.DataDictionaryName))
                    {
                        dataDictionaryNames.Add(attribute.DataDictionaryName);
                        Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    添加数据字典名称: {attribute.DataDictionaryName}");
                    }
                }
                else
                {
                    Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    未找到有效的DataDictionaryAttribute");
                }
            }
        }
        
        // 手动添加Order类中的数据字典名称，确保它们被创建
        // 这是一个临时解决方案，用于确保数据字典被创建
        if (!dataDictionaryNames.Contains("部门"))
        {
            dataDictionaryNames.Add("部门");
            Tracing.Tracer.LogText("=== DataDictionaryUpdater ===    手动添加数据字典名称: 部门");
        }
        if (!dataDictionaryNames.Contains("职位"))
        {
            dataDictionaryNames.Add("职位");
            Tracing.Tracer.LogText("=== DataDictionaryUpdater ===    手动添加数据字典名称: 职位");
        }

        Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===  收集到 {dataDictionaryNames.Count} 个数据字典名称");
        foreach (var name in dataDictionaryNames)
        {
            Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    数据字典名称: {name}");
        }

        // 遍历所有数据字典名称，创建对应的数据字典
        foreach (var dataDictionaryName in dataDictionaryNames)
        {
            Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===  处理数据字典名称: {dataDictionaryName}");
            
            // 检查数据字典是否已存在
            // 使用ObjectSpace.FindObject方法代替FirstOrDefault方法
            var dataDictionary = ObjectSpace.FindObject<DataDictionary>(
                new DevExpress.Data.Filtering.BinaryOperator("Name", dataDictionaryName));
            if (dataDictionary == null)
            {
                // 如果不存在，创建新的数据字典
                Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    数据字典不存在，创建新的数据字典: {dataDictionaryName}");
                dataDictionary = ObjectSpace.CreateObject<DataDictionary>();
                dataDictionary.Name = dataDictionaryName;
            }
            else
            {
                Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===    数据字典已存在: {dataDictionaryName}");
            }
        }

        // 提交更改
        Tracing.Tracer.LogText($"=== DataDictionaryUpdater ===  提交更改");
        ObjectSpace.CommitChanges();
        Tracing.Tracer.LogText($"=== DataDictionaryUpdater === 执行完成 ===");
    }
}
