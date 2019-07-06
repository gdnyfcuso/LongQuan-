using DBEN.DBI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LongQuanMessageManager
{
    public partial class MessageManager : Form
    {
        public MessageManager()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openExeclFileDialog.Filter = "execl(*.xls,*.xlsx)|*.xls;*.xlsx";
            if (openExeclFileDialog.ShowDialog() == DialogResult.OK)
            {
                tbExeclFilePath.Text = openExeclFileDialog.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (saveTxtFileDialog.ShowDialog() == DialogResult.OK)
            {
                tbSaveFile.Text = saveTxtFileDialog.FileName;
            }
        }

        private void btnExecl_Click(object sender, EventArgs e)
        {
            try
            {
                var fileName = tbExeclFilePath.Text;
                if (string.IsNullOrWhiteSpace(fileName))
                {
                    MessageBox.Show("请选择要处理的execl文件");
                    return;
                }
                var phoneDic = this.InitPhoneContext(fileName, txtPhoneColumnName.Text, richTextBox1.Text);
                StringBuilder sb = new StringBuilder();
                foreach (var item in phoneDic)
                {
                    sb.Append("\r\n" + item.Value + "\r\n");
                }
                richTextBox2.Text = sb.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void saveExeclResult_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(tbSaveFile.Text))
            {
                MessageBox.Show("请选择要保存的文件");
                return;
            }
            File.AppendAllText(tbSaveFile.Text, richTextBox2.Text);
            MessageBox.Show($"文件成功保存到{tbSaveFile.Text}");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openExeclFileDialog.Filter = "文件文件(*.txt,*.data)|*.txt;*.data";
            if (openExeclFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openExeclFileDialog.FileName;
                richTextBox1.Text = File.ReadAllText(textBox1.Text, Encoding.Default);
            }
        }

        //private string url = "http://utf8.sms.webchinese.cn/?";
        //private string strUid = "Uid=";
        //private string strKey = "&key=*******************"; //这里*代表秘钥，由于从长有点麻烦，就不在窗口上输入了
        //private string strMob = "&smsMob=";
        //private string strContent = "&smsText=";

        private void btnSendMessage_Click(object sender, EventArgs e)
        {
            var fileName = tbExeclFilePath.Text;
            if (string.IsNullOrWhiteSpace(fileName))
            {
                MessageBox.Show("请选择要处理的execl文件");
                return;
            }
            var phoneDic = this.InitPhoneContext(fileName, txtPhoneColumnName.Text, richTextBox1.Text);

            var mesModel = new MessageModel();
            mesModel.UserName = sendUserName.Text;
            mesModel.Key = MessageKey.Text;
            mesModel.MessagePhone = phoneDic;

            var mesContext = new MessageContext();
            mesContext.Content = mesModel;
            mesContext.StatusLabel = labResult;
            this.richTextBoxResult.Text = string.Empty;
            mesContext.Result = this.richTextBoxResult;
            ThreadPool.QueueUserWorkItem(this.InitSendMessage, mesContext);
        }

        /// <summary>
        /// 初始化手机数据短信数据集合
        /// key为手机号
        /// value为手机要发送的内容
        /// </summary>
        /// <param name="fileName">要用NPOI解析的execl地址</param>
        /// <returns></returns>
        public Dictionary<string, string> InitPhoneContext(string fileName, string ExeclphoneColumnName, string modelContext)
        {
            var phonedic = new Dictionary<string, string>();
            try
            {
                if (string.IsNullOrWhiteSpace(fileName))
                {
                    throw new Exception("请选择要处理的execl文件");
                }
                var dt = ExcelHelper.ImportExceltoDt(fileName, 0, "1");
                StringBuilder sb = new StringBuilder();

                Dictionary<string, string> cKey = new Dictionary<string, string>();

                string keyName = string.Empty;
                foreach (DataColumn column in dt.Columns)
                {
                    keyName = column.ColumnName;
                    if (!cKey.ContainsKey(keyName))
                    {
                        cKey.Add(keyName, $"[{keyName}]");
                    }
                }
                if (!cKey.ContainsKey(ExeclphoneColumnName))
                {
                    throw new Exception("execl里面指定的电话列不存在");
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string fileContext = modelContext;
                    foreach (DataColumn column in dt.Columns)
                    {
                        fileContext = fileContext.Replace(cKey[column.ColumnName].ToString(), dt.Rows[i][column.ColumnName].ToString());
                    }
                    if (!phonedic.ContainsKey(dt.Rows[i][ExeclphoneColumnName].ToString()))
                    {
                        phonedic.Add(dt.Rows[i][ExeclphoneColumnName].ToString(), fileContext.Trim());
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return phonedic;
        }

        /// <summary>
        /// 发送短信的结果进行转换，UTF-8 http://sms.webchinese.cn
        /// -1	没有该用户账户
        ///-2	接口密钥不正确[查看密钥] 不是账户登陆密码
        ///-21	MD5接口密钥加密不正确
        ///-3	短信数量不足
        ///-11	该用户被禁用
        ///-14	短信内容出现非法字符
        ///-4	手机号格式不正确
        ///-41	手机号码为空
        ///-42	短信内容为空
        ///-51	短信签名格式不正确 接口签名格式为：【签名内容】
        ///-6	IP限制
        ///大于0 短信发送数量
        /// </summary>
        /// <param name="resultCode"></param>
        /// <returns></returns>
        public string ChangMessageResult(string resultCode)
        {
            string result = string.Empty;
            try
            {
                switch (resultCode)
                {
                    case "-1": result = "没有该用户账户"; break;
                    case "-2": result = "接口密钥不正确[查看密钥] 不是账户登陆密码"; break;
                    case "-21": result = "MD5接口密钥加密不正确"; break;
                    case "-3": result = "短信数量不足"; break;
                    case "-11": result = "该用户被禁用"; break;
                    case "-14": result = "短信内容出现非法字符"; break;
                    case "-4": result = "手机号格式不正确"; break;
                    case "-41": result = "手机号码为空"; break;
                    case "-42": result = "短信内容为空"; break;
                    case "-51": result = "短信签名格式不正确 接口签名格式为：【签名内容】"; break;
                    case "-6": result = "IP限制"; break;

                    default:
                        if (Convert.ToInt32(resultCode) > 0)
                            result = "发送成功" + resultCode + "条";
                        else
                            result = "无法识别的返回结果";
                        break;
                }
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }
            return result;
        }

        int messageCount = 0;
        /// <summary>
        /// 初始化发送短信
        /// </summary>
        /// <param name="userName">要发送的用户名</param>
        /// <param name="recivePhone">要接收内容的手机号</param>
        /// <param name="messageContent">要发送的内容</param>
        /// <returns></returns>
        public void InitSendMessage(object messageContext)
        {
            messageCount = 0;
            var context = (MessageContext)messageContext;
            string url = "http://utf8.sms.webchinese.cn/?";
            string strUid = "Uid=";
            string strKey = "&key=" + context.Content.Key; //这里*代表秘钥，由于从长有点麻烦，就不在窗口上输入了
            string strMob = "&smsMob=";
            string strContent = "&smsText=";
            string sendUrl = string.Empty;
            foreach (var item in context.Content.MessagePhone)
            {
                Thread.Sleep(1000);
                sendUrl = url + strUid + context.Content.UserName + strKey + strMob + item.Key + strContent + item.Value;
                var result = this.ChangMessageResult(this.GetHtmlFromUrl(sendUrl));
                messageCount++;
                context.Result.Invoke(new Action(() => {
                    context.Result.Text += "\r\n" + "电话:" + item.Key + "结果:" + result + "\r\n";
                }));
                context.StatusLabel.Invoke(new Action(() => {
                    context.StatusLabel.Text = messageCount+"条";
                }));
            }
            context.StatusLabel.Invoke(new Action(() => {
                context.StatusLabel.Text = "短信发送完毕，总共发送了:" +context.Content.MessagePhone.Keys.Count + "条";
            }));
        }

        /// <summary>
        /// 发送短信核心
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public string GetHtmlFromUrl(string url)
        {
            string strRet = null;
            if (url == null || url.Trim().ToString() == "")
            {
                return strRet;
            }
            string targeturl = url.Trim().ToString();
            try
            {
                HttpWebRequest hr = (HttpWebRequest)WebRequest.Create(targeturl);
                hr.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)";
                hr.Method = "GET";
                hr.Timeout = 30 * 60 * 1000;
                WebResponse hs = hr.GetResponse();
                Stream sr = hs.GetResponseStream();
                StreamReader ser = new StreamReader(sr, Encoding.Default);
                strRet = ser.ReadToEnd();
            }
            catch (Exception ex)
            {
                strRet = "发送失败" + ex.Message;
            }
            return strRet;
        }

        /// <summary>
        /// 单次发送短信
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            var mesModel = new MessageModel();
            mesModel.UserName = sendUserName.Text;
            mesModel.Key = MessageKey.Text;
            mesModel.MessagePhone.Add(txtPhone.Text, sendContent.Text);
            this.richTextBoxResult.Text = string.Empty;
            var mesContext = new MessageContext();
            mesContext.Content = mesModel;
            mesContext.StatusLabel = labResult;
            mesContext.Result = this.richTextBoxResult;
            this.InitSendMessage(mesContext);
        }
    }
}
