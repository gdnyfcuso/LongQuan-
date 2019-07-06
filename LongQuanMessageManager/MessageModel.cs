/*******************************************************
 * 
 * 作者：李行周
 * 创建日期：20161222
 * 说明：此文件只包含一个类，具体内容见类型注释。
 * 版本号：1.0.0
 * 
 * 历史记录：
 * 创建文件 李行周 20161222 13:47
 * 
*******************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LongQuanMessageManager
{
    /// <summary>
    /// 发送消息运行的实体
    /// </summary>
    public class MessageModel
    {
        /// <summary>
        /// 在sms网站上使用的用户名
        /// </summary>
        public string UserName { get; set; }

        /// <summary>
        /// 发送短信对应的密钥
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// 要发送的短信集合，key为要发送的手机号，value为要发送的内容
        /// </summary>
        public Dictionary<string, string> MessagePhone { get; set; }

        /// <summary>
        /// 初始化要发送的短信集合
        /// </summary>
        public MessageModel()
        {
            MessagePhone = new Dictionary<string, string>();
        }

    }

    public class MessageContext
    {
        /// <summary>
        /// 发送短信之后的显示框
        /// </summary>
        public RichTextBox Result { get; set; }

        /// <summary>
        /// 发送之后的状态更新
        /// </summary>
        public Label StatusLabel { get; set; }

        /// <summary>
        /// 要发送的数据集
        /// </summary>
        public MessageModel Content { get; set; }

    }
}
