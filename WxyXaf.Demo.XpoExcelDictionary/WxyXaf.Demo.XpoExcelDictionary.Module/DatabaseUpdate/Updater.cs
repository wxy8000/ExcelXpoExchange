using DevExpress.Data.Filtering;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Updating;
using DevExpress.ExpressApp.Xpo;
using DevExpress.Persistent.Base;
using DevExpress.Persistent.BaseImpl;
using DevExpress.Persistent.Validation;
using DevExpress.Xpo;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using WxyXaf.DataDictionaries;
using WxyXaf.Demo.XpoExcelDictionary.Module.BusinessObjects;

namespace WxyXaf.Demo.XpoExcelDictionary.Module.DatabaseUpdate
{
    // For more typical usage scenarios, be sure to check out https://docs.devexpress.com/eXpressAppFramework/DevExpress.ExpressApp.Updating.ModuleUpdater
    public class Updater : ModuleUpdater
    {
        public Updater(IObjectSpace objectSpace, Version currentDBVersion) :
            base(objectSpace, currentDBVersion)
        {
        }
        public override void UpdateDatabaseAfterUpdateSchema()
        {
            base.UpdateDatabaseAfterUpdateSchema();
            
            Console.WriteLine("[Updater] 开始更新数据库...");
            
            // 清空所有现有数据
            ClearAllData();
            
            // 创建模拟数据
            CreateDataDictionaries();
            Console.WriteLine("[Updater] 数据字典创建完成");
            
            // 创建产品和客户
            var products = CreateProducts();
            Console.WriteLine("[Updater] 产品数据创建完成，数量: " + products.Count);
            
            var customers = CreateCustomers();
            Console.WriteLine("[Updater] 客户数据创建完成，数量: " + customers.Count);
            
            // 创建订单，传入产品和客户列表
            CreateOrders(products, customers);
            Console.WriteLine("[Updater] 订单数据创建完成");

            ObjectSpace.CommitChanges(); // 持久化创建的对象
            Console.WriteLine("[Updater] 数据库更新完成");
        }
        
        /// <summary>
        /// 清空所有现有数据
        /// </summary>
        private void ClearAllData()
        {
            // 先清空订单数据（有外键约束，需要先于产品和客户删除）
            var orders = ObjectSpace.GetObjects<订单>().ToList(); // 转换为列表，避免枚举时修改集合
            foreach (var order in orders)
            {
                ObjectSpace.Delete(order);
            }
            
            // 清空产品数据
            var products = ObjectSpace.GetObjects<产品>().ToList(); // 转换为列表，避免枚举时修改集合
            foreach (var product in products)
            {
                ObjectSpace.Delete(product);
            }
            
            // 清空客户数据
            var customers = ObjectSpace.GetObjects<客户>().ToList(); // 转换为列表，避免枚举时修改集合
            foreach (var customer in customers)
            {
                ObjectSpace.Delete(customer);
            }
            
            // 清空数据字典项
            var dictionaryItems = ObjectSpace.GetObjects<DataDictionaryItem>().ToList(); // 转换为列表，避免枚举时修改集合
            foreach (var item in dictionaryItems)
            {
                ObjectSpace.Delete(item);
            }
            
            // 清空数据字典
            var dictionaries = ObjectSpace.GetObjects<DataDictionary>().ToList(); // 转换为列表，避免枚举时修改集合
            foreach (var dict in dictionaries)
            {
                ObjectSpace.Delete(dict);
            }
            
            // 提交删除操作
            ObjectSpace.CommitChanges();
        }
        
        /// <summary>
        /// 创建数据字典和数据字典项
        /// </summary>
        private void CreateDataDictionaries()
        {
            // 产品分类
            var productCategoryDict = GetOrCreateDataDictionary("产品分类");
            GetOrCreateDataDictionaryItem(productCategoryDict, "电子产品", "Electronic");
            GetOrCreateDataDictionaryItem(productCategoryDict, "服装", "Clothing");
            GetOrCreateDataDictionaryItem(productCategoryDict, "食品", "Food");
            GetOrCreateDataDictionaryItem(productCategoryDict, "家居用品", "Home");
            
            // 品牌
            var brandDict = GetOrCreateDataDictionary("品牌");
            GetOrCreateDataDictionaryItem(brandDict, "苹果", "Apple");
            GetOrCreateDataDictionaryItem(brandDict, "三星", "Samsung");
            GetOrCreateDataDictionaryItem(brandDict, "华为", "Huawei");
            GetOrCreateDataDictionaryItem(brandDict, "小米", "Xiaomi");
            GetOrCreateDataDictionaryItem(brandDict, "联想", "Lenovo");
            
            // 计量单位
            var unitDict = GetOrCreateDataDictionary("计量单位");
            GetOrCreateDataDictionaryItem(unitDict, "件", "Piece");
            GetOrCreateDataDictionaryItem(unitDict, "台", "Unit");
            GetOrCreateDataDictionaryItem(unitDict, "个", "Item");
            GetOrCreateDataDictionaryItem(unitDict, "斤", "Jin");
            
            // 产品状态
            var productStatusDict = GetOrCreateDataDictionary("产品状态");
            GetOrCreateDataDictionaryItem(productStatusDict, "在售", "OnSale");
            GetOrCreateDataDictionaryItem(productStatusDict, "下架", "OffSale");
            GetOrCreateDataDictionaryItem(productStatusDict, "缺货", "OutOfStock");
            
            // 客户分类
            var customerCategoryDict = GetOrCreateDataDictionary("客户分类");
            GetOrCreateDataDictionaryItem(customerCategoryDict, "个人客户", "Individual");
            GetOrCreateDataDictionaryItem(customerCategoryDict, "企业客户", "Enterprise");
            GetOrCreateDataDictionaryItem(customerCategoryDict, "VIP客户", "VIP");
            
            // 客户状态
            var customerStatusDict = GetOrCreateDataDictionary("客户状态");
            GetOrCreateDataDictionaryItem(customerStatusDict, "活跃", "Active");
            GetOrCreateDataDictionaryItem(customerStatusDict, "休眠", "Inactive");
            GetOrCreateDataDictionaryItem(customerStatusDict, "流失", "Churned");
            
            // 订单状态
            var orderStatusDict = GetOrCreateDataDictionary("订单状态");
            GetOrCreateDataDictionaryItem(orderStatusDict, "新建", "New");
            GetOrCreateDataDictionaryItem(orderStatusDict, "已付款", "Paid");
            GetOrCreateDataDictionaryItem(orderStatusDict, "已发货", "Shipped");
            GetOrCreateDataDictionaryItem(orderStatusDict, "已完成", "Completed");
            GetOrCreateDataDictionaryItem(orderStatusDict, "已取消", "Cancelled");
            
            // 支付方式
            var paymentMethodDict = GetOrCreateDataDictionary("支付方式");
            GetOrCreateDataDictionaryItem(paymentMethodDict, "支付宝", "Alipay");
            GetOrCreateDataDictionaryItem(paymentMethodDict, "微信支付", "WeChatPay");
            GetOrCreateDataDictionaryItem(paymentMethodDict, "银行卡", "BankCard");
            GetOrCreateDataDictionaryItem(paymentMethodDict, "货到付款", "CashOnDelivery");
            
            // 配送方式
            var shippingMethodDict = GetOrCreateDataDictionary("配送方式");
            GetOrCreateDataDictionaryItem(shippingMethodDict, "顺丰快递", "SFExpress");
            GetOrCreateDataDictionaryItem(shippingMethodDict, "中通快递", "ZTO");
            GetOrCreateDataDictionaryItem(shippingMethodDict, "圆通快递", "YTO");
            GetOrCreateDataDictionaryItem(shippingMethodDict, "韵达快递", "Yunda");
        }
        
        /// <summary>
        /// 获取或创建数据字典
        /// </summary>
        private DataDictionary GetOrCreateDataDictionary(string name)
        {
            var dict = ObjectSpace.FirstOrDefault<DataDictionary>(d => d.Name == name);
            if (dict == null)
            {
                dict = ObjectSpace.CreateObject<DataDictionary>();
                dict.Name = name;
            }
            return dict;
        }
        
        /// <summary>
        /// 获取或创建数据字典项
        /// </summary>
        private DataDictionaryItem GetOrCreateDataDictionaryItem(DataDictionary dictionary, string name, string code)
        {
            var item = ObjectSpace.FirstOrDefault<DataDictionaryItem>(i => i.Name == name && i.DataDictionary == dictionary);
            if (item == null)
            {
                item = ObjectSpace.CreateObject<DataDictionaryItem>();
                item.Name = name;
                item.Code = code;
                item.DataDictionary = dictionary;
            }
            return item;
        }
        
        /// <summary>
        /// 创建产品模拟数据
        /// </summary>
        private List<产品> CreateProducts()
        {
            // 获取数据字典项
            var productCategoryDict = GetOrCreateDataDictionary("产品分类");
            var productCategories = ObjectSpace.GetObjects<DataDictionaryItem>(CriteriaOperator.Parse("[DataDictionary.Name] = '产品分类'")).ToList();
            var brandDict = GetOrCreateDataDictionary("品牌");
            var brands = ObjectSpace.GetObjects<DataDictionaryItem>(CriteriaOperator.Parse("[DataDictionary.Name] = '品牌'")).ToList();
            var unitDict = GetOrCreateDataDictionary("计量单位");
            var units = ObjectSpace.GetObjects<DataDictionaryItem>(CriteriaOperator.Parse("[DataDictionary.Name] = '计量单位'")).ToList();
            var productStatusDict = GetOrCreateDataDictionary("产品状态");
            var productStatuses = ObjectSpace.GetObjects<DataDictionaryItem>(CriteriaOperator.Parse("[DataDictionary.Name] = '产品状态'")).ToList();
            
            // 产品列表
            var productsData = new List<(string name, string code, string category, string brand, string unit, decimal price, decimal cost, int stock, string status)>
            {
                ("iPhone 15 Pro", "IP15PRO", "电子产品", "苹果", "台", 8999, 6999, 100, "在售"),
                ("Galaxy S24", "GS24", "电子产品", "三星", "台", 7999, 5999, 80, "在售"),
                ("Mate 60 Pro", "MT60PRO", "电子产品", "华为", "台", 6999, 4999, 120, "在售"),
                ("Redmi Note 13", "RN13", "电子产品", "小米", "台", 1999, 1499, 200, "在售"),
                ("ThinkPad X1", "TPX1", "电子产品", "联想", "台", 9999, 7999, 50, "在售"),
                ("纯棉T恤", "COTTONT", "服装", "小米", "件", 99, 59, 300, "在售"),
                ("牛仔裤", "JEANS", "服装", "苹果", "件", 199, 119, 150, "在售"),
                ("牛奶", "MILK", "食品", "华为", "件", 68, 48, 250, "在售"),
                ("面包", "BREAD", "食品", "三星", "个", 12, 8, 500, "在售"),
                ("洗衣液", "DETERGENT", "家居用品", "联想", "件", 39, 25, 180, "在售")
            };
            
            var createdProducts = new List<产品>();
            
            foreach (var (name, code, category, brand, unit, price, cost, stock, status) in productsData)
            {
                // 检查产品是否已存在
                var existingProduct = ObjectSpace.FirstOrDefault<产品>(p => p.产品编码 == code);
                if (existingProduct == null)
                {
                    var product = ObjectSpace.CreateObject<产品>();
                    product.产品名称 = name;
                    product.产品编码 = code;
                    
                    // 使用FirstOrDefault替代First，防止找不到匹配元素时抛出异常
                    product.产品分类 = productCategories.FirstOrDefault(c => c.Name == category) ?? GetOrCreateDataDictionaryItem(productCategoryDict, category, category);
                    product.品牌 = brands.FirstOrDefault(b => b.Name == brand) ?? GetOrCreateDataDictionaryItem(brandDict, brand, brand);
                    product.计量单位 = units.FirstOrDefault(u => u.Name == unit) ?? GetOrCreateDataDictionaryItem(unitDict, unit, unit);
                    product.销售价格 = price;
                    product.成本价格 = cost;
                    product.库存数量 = stock;
                    product.产品状态 = productStatuses.FirstOrDefault(s => s.Name == status) ?? GetOrCreateDataDictionaryItem(productStatusDict, status, status);
                    product.生产日期 = DateTime.Now.AddDays(-30);
                    product.产品描述 = $"这是一款优质的{name}，质量保证，欢迎购买！";
                    
                    createdProducts.Add(product);
                }
                else
                {
                    createdProducts.Add(existingProduct);
                }
            }
            
            return createdProducts;
        }
        
        /// <summary>
        /// 创建客户模拟数据
        /// </summary>
        private List<客户> CreateCustomers()
        {
            // 获取数据字典项
            var customerCategoryDict = GetOrCreateDataDictionary("客户分类");
            var customerCategories = ObjectSpace.GetObjects<DataDictionaryItem>(CriteriaOperator.Parse("[DataDictionary.Name] = '客户分类'")).ToList();
            var customerStatusDict = GetOrCreateDataDictionary("客户状态");
            var customerStatuses = ObjectSpace.GetObjects<DataDictionaryItem>(CriteriaOperator.Parse("[DataDictionary.Name] = '客户状态'")).ToList();
            
            // 客户列表
            var customersData = new List<(string name, string id, string category, string phone, string email, string status)>
            {
                ("张三", "CUST001", "个人客户", "13800138001", "zhangsan@example.com", "活跃"),
                ("李四", "CUST002", "个人客户", "13800138002", "lisi@example.com", "活跃"),
                ("王五", "CUST003", "VIP客户", "13800138003", "wangwu@example.com", "活跃"),
                ("赵六", "CUST004", "企业客户", "13800138004", "zhaoliu@example.com", "休眠"),
                ("孙七", "CUST005", "个人客户", "13800138005", "sunqi@example.com", "活跃"),
                ("周八", "CUST006", "企业客户", "13800138006", "zhouba@example.com", "活跃"),
                ("吴九", "CUST007", "VIP客户", "13800138007", "wujiu@example.com", "流失"),
                ("郑十", "CUST008", "个人客户", "13800138008", "zhengshi@example.com", "活跃")
            };
            
            var createdCustomers = new List<客户>();
            
            foreach (var (name, id, category, phone, email, status) in customersData)
            {
                // 检查客户是否已存在
                var existingCustomer = ObjectSpace.FirstOrDefault<客户>(c => c.客户ID == id);
                if (existingCustomer == null)
                {
                    var customer = ObjectSpace.CreateObject<客户>();
                    customer.客户名称 = name;
                    customer.客户ID = id;
                    customer.客户分类 = customerCategories.FirstOrDefault(c => c.Name == category) ?? GetOrCreateDataDictionaryItem(customerCategoryDict, category, category);
                    customer.电话 = phone;
                    customer.邮箱 = email;
                    customer.客户状态 = customerStatuses.FirstOrDefault(s => s.Name == status) ?? GetOrCreateDataDictionaryItem(customerStatusDict, status, status);
                    customer.注册日期 = DateTime.Now.AddDays(-new Random().Next(1, 365));
                    customer.备注 = $"这是{name}的客户记录";
                    
                    createdCustomers.Add(customer);
                }
                else
                {
                    createdCustomers.Add(existingCustomer);
                }
            }
            
            return createdCustomers;
        }
        
        /// <summary>
        /// 创建订单模拟数据
        /// </summary>
        private void CreateOrders(List<产品> products, List<客户> customers)
        {
            // 确保订单相关的数据字典项存在
            var orderStatusDict = GetOrCreateDataDictionary("订单状态");
            var paymentMethodDict = GetOrCreateDataDictionary("支付方式");
            var shippingMethodDict = GetOrCreateDataDictionary("配送方式");
            
            // 获取或创建订单状态
            var newStatus = GetOrCreateDataDictionaryItem(orderStatusDict, "新建", "New");
            var paidStatus = GetOrCreateDataDictionaryItem(orderStatusDict, "已付款", "Paid");
            var shippedStatus = GetOrCreateDataDictionaryItem(orderStatusDict, "已发货", "Shipped");
            var completedStatus = GetOrCreateDataDictionaryItem(orderStatusDict, "已完成", "Completed");
            var cancelledStatus = GetOrCreateDataDictionaryItem(orderStatusDict, "已取消", "Cancelled");
            
            // 获取或创建支付方式
            var alipay = GetOrCreateDataDictionaryItem(paymentMethodDict, "支付宝", "Alipay");
            var wechatPay = GetOrCreateDataDictionaryItem(paymentMethodDict, "微信支付", "WeChatPay");
            var bankCard = GetOrCreateDataDictionaryItem(paymentMethodDict, "银行卡", "BankCard");
            var cashOnDelivery = GetOrCreateDataDictionaryItem(paymentMethodDict, "货到付款", "CashOnDelivery");
            
            // 获取或创建配送方式
            var sfExpress = GetOrCreateDataDictionaryItem(shippingMethodDict, "顺丰快递", "SFExpress");
            var zto = GetOrCreateDataDictionaryItem(shippingMethodDict, "中通快递", "ZTO");
            var yto = GetOrCreateDataDictionaryItem(shippingMethodDict, "圆通快递", "YTO");
            var yunda = GetOrCreateDataDictionaryItem(shippingMethodDict, "韵达快递", "Yunda");
            
            // 订单状态列表
            var orderStatuses = new List<DataDictionaryItem> { newStatus, paidStatus, shippedStatus, completedStatus, cancelledStatus };
            // 支付方式列表
            var paymentMethods = new List<DataDictionaryItem> { alipay, wechatPay, bankCard, cashOnDelivery };
            // 配送方式列表
            var shippingMethods = new List<DataDictionaryItem> { sfExpress, zto, yto, yunda };
            
            // 打印传入的产品和客户列表
            Console.WriteLine($"[Updater] 开始创建订单...");
            Console.WriteLine($"[Updater] 传入的产品列表数量: {products.Count}");
            foreach (var product in products)
            {
                Console.WriteLine($"[Updater] 产品: {product.产品名称}, ID: {product.Oid}, 价格: {product.销售价格}");
            }
            
            Console.WriteLine($"[Updater] 传入的客户列表数量: {customers.Count}");
            foreach (var customer in customers)
            {
                Console.WriteLine($"[Updater] 客户: {customer.客户名称}, ID: {customer.Oid}, 客户ID: {customer.客户ID}");
            }
            
            // 生成随机数种子
            var random = new Random();
            
            // 生成10个订单
            for (int i = 1; i <= 10; i++)
            {
                // 使用唯一的订单编号
                var orderNo = $"ORD{i.ToString().PadLeft(5, '0')}";
                Console.WriteLine($"[Updater] 创建订单 {orderNo}...");
                var existingOrder = ObjectSpace.FirstOrDefault<订单>(o => o.订单编号 == orderNo);
                
                if (existingOrder == null)
                {
                    var order = ObjectSpace.CreateObject<订单>();
                    order.订单编号 = orderNo;
                    order.订单日期 = DateTime.Now.AddDays(-random.Next(1, 30));
                    order.订单状态 = orderStatuses[random.Next(orderStatuses.Count)];
                    order.支付方式 = paymentMethods[random.Next(paymentMethods.Count)];
                    order.配送方式 = shippingMethods[random.Next(shippingMethods.Count)];
                    
                    // 确保产品和客户列表非空
                    Console.WriteLine($"[Updater] 检查产品和客户列表数量 - 产品: {products.Count}, 客户: {customers.Count}");
                    if (products.Count > 0 && customers.Count > 0)
                    {
                        // 随机分配产品
                        var productIndex = random.Next(products.Count);
                        var product = products[productIndex];
                        Console.WriteLine($"[Updater] 分配产品 {product.产品名称} (索引: {productIndex}) 到订单 {orderNo}");
                        order.产品 = product;
                        order.单价 = product.销售价格;
                        Console.WriteLine($"[Updater] 订单 {orderNo} 产品设置成功: {order.产品?.产品名称}, 单价: {order.单价}");
                        
                        // 随机分配客户
                        var customerIndex = random.Next(customers.Count);
                        var customer = customers[customerIndex];
                        Console.WriteLine($"[Updater] 分配客户 {customer.客户名称} (索引: {customerIndex}) 到订单 {orderNo}");
                        order.客户 = customer;
                        Console.WriteLine($"[Updater] 订单 {orderNo} 客户设置成功: {order.客户?.客户名称}");
                        
                        // 设置数量
                        order.数量 = random.Next(1, 10);
                        Console.WriteLine($"[Updater] 订单 {orderNo} 数量设置为: {order.数量}");
                        
                        // 更新客户的购买产品列表
                        Console.WriteLine($"[Updater] 更新客户 {customer.客户名称} 的购买产品列表...");
                        if (!customer.产品.Contains(product))
                        {
                            customer.产品.Add(product);
                            Console.WriteLine($"[Updater] 将产品 {product.产品名称} 添加到客户 {customer.客户名称} 的购买产品列表中");
                        }
                        else
                        {
                            Console.WriteLine($"[Updater] 产品 {product.产品名称} 已在客户 {customer.客户名称} 的购买产品列表中");
                        }
                        
                        order.备注 = $"这是第{i}个测试订单";
                        Console.WriteLine($"[Updater] 订单 {orderNo} 创建完成");
                    }
                    else
                    {
                        Console.WriteLine($"[Updater] 产品或客户列表为空，跳过订单 {orderNo} 创建");
                    }
                }
                else
                {
                    Console.WriteLine($"[Updater] 订单 {orderNo} 已存在，跳过");
                }
            }
            Console.WriteLine($"[Updater] 订单创建完成");
        }
        
        public override void UpdateDatabaseBeforeUpdateSchema()
        {
            base.UpdateDatabaseBeforeUpdateSchema();
            //if(CurrentDBVersion < new Version("1.1.0.0") && CurrentDBVersion > new Version("0.0.0.0")) {
            //    RenameColumn("DomainObject1Table", "OldColumnName", "NewColumnName");
            //}
        }
    }
}