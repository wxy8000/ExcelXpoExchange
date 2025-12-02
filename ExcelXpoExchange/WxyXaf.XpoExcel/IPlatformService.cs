using System;
using System.Threading.Tasks;

namespace WxyXaf.XpoExcel
{
    /// <summary>
    /// 平台服务接口，定义各平台特有的操�?
    /// </summary>
    public interface IPlatformService
    {
        /// <summary>
        /// 显示保存文件对话�?
        /// </summary>
        /// <param name="defaultFileName">默认文件�?/param>
        /// <param name="filter">文件过滤�?/param>
        /// <returns>用户选择的文件路径，如果用户取消则返回null</returns>
        string ShowSaveFileDialog(string defaultFileName, string filter);
        
        /// <summary>
        /// 显示打开文件对话�?
        /// </summary>
        /// <param name="filter">文件过滤�?/param>
        /// <returns>用户选择的文件路径，如果用户取消则返回null</returns>
        string ShowOpenFileDialog(string filter);
        
        /// <summary>
        /// 显示消息
        /// </summary>
        /// <param name="message">消息内容</param>
        /// <param name="type">消息类型</param>
        void ShowMessage(string message, MessageType type);
        
        /// <summary>
        /// 显示导入模式选择对话�?
        /// </summary>
        /// <returns>导入模式</returns>
        ImportMode ShowImportModeSelectionDialog();
        
        /// <summary>
        /// 下载文件
        /// </summary>
        /// <param name="fileName">文件�?/param>
        /// <param name="fileBytes">文件字节数据</param>
        /// <param name="contentType">内容类型</param>
        void DownloadFile(string fileName, byte[] fileBytes, string contentType);
        
        /// <summary>
        /// 导航到导入页�?
        /// </summary>
        /// <param name="objectTypeName">对象类型的程序集限定�?/param>
        void NavigateToImportPage(string objectTypeName);
    }
    
    /// <summary>
    /// 消息类型
    /// </summary>
    public enum MessageType
    {
        Info,
        Success,
        Warning,
        Error
    }
}
