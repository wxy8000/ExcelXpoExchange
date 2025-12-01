using System;
using DevExpress.ExpressApp.Model;
using DevExpress.ExpressApp.Xpo;

namespace Wxy.Common
{
    /// <summary>
    /// 关联对象转换器接口，用于抽象关联对象转换逻辑
    /// </summary>
    public interface IRelatedObjectConverter
    {
        /// <summary>
        /// 判断当前转换器是否支持指定类型
        /// </summary>
        /// <param name="objectType">要转换的对象类型</param>
        /// <returns>是否支持</returns>
        bool CanConvert(Type objectType);
        
        /// <summary>
        /// 将Excel单元格值转换为关联对象
        /// </summary>
        /// <param name="cellValue">Excel单元格值</param>
        /// <param name="member">模型成员信息</param>
        /// <param name="objectSpace">对象空间</param>
        /// <returns>转换后的关联对象</returns>
        object Convert(object cellValue, IModelMember member, XPObjectSpace objectSpace);
    }
}