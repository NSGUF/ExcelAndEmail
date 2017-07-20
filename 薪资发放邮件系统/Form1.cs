using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
namespace 薪资发放邮件系统
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //this.skinEngine1.SkinFile = "DiamondBlue.ssk"; //DiamondBlue.ssk可换用皮肤目录中你喜欢的.ssk文件
        }
        #region 全局变量
        string sendPath;//获得发送邮箱文件完整路径
        string rePath;//获得接受邮箱文件完整路径
        string salaryPath;//获得工资明细文件完整路劲

        DateTime dtStart;//开始发送的时间
        string wrong = "";//错误代码
        int allCount;//日志文件一共多少行

        string emailNum = "";//发送邮件账号
        string emailPwd = "";//发送邮件密码

        DataSet salaryDS;//获取工资表数据

        DataTable dt;
        DataSet emailDS;//邮箱数据

        int num1;//已发送成功了的信息条数
        int num2;//已发送成功了的信息条数
        int num3;//已发送成功了的信息条数
        int num4;//已发送成功了的信息条数
        int num5;//已发送成功了的信息条数

        Thread th1;//发送邮件的线程
        Thread th2;//发送邮件的线程
        Thread th3;//发送邮件的线程
        Thread th4;//发送邮件的线程
        Thread th5;//发送邮件的线程
        Thread th6;//写入未发送人的信息的线程


        List<int> list1 = new List<int>();//存储发送失败的行号
        List<int> list2 = new List<int>();//存储发送失败的行号
        List<int> list3 = new List<int>();//存储发送失败的行号
        List<int> list4 = new List<int>();//存储发送失败的行号
        List<int> list5 = new List<int>();//存储发送失败的行号

        string[] title = new string[100];  //无限制条件下，获取待发送excel文件第一行数据
        int countTitle = 0;// 无限制条件下，获取待发送excel文件的总列数
        int sendExcelCount = 0;//总共发送文件的个数
        #endregion

        #region 选择发送的邮箱账号
        private void selectSendBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "选择文件";
            ofd.InitialDirectory = Application.StartupPath + "\\发送邮箱";//打开控件后，默认目录
            ofd.Filter = "所有文件(*.*)|*.*|Excle文件(*.xls)|*.xls|Excle文件(*.xlsx)|*.xlsx";//打开文件类型
            ofd.RestoreDirectory = true;//设置对话框是否记忆之前打开的目录
            ofd.Multiselect = false;//是否可以同时打开多个文件
            if (ofd.ShowDialog() == DialogResult.OK)//是否选择了文件
            {
                sendPath = ofd.FileName.ToString();//获得用户选择的文件完整路径
                if (Path.GetExtension(sendPath) == ".xls" || Path.GetExtension(sendPath) == ".xlsx")//判断是否为excel文件
                {
                    sendPathText.Text = sendPath;
                    DataSet sendDs = null;
                    try//用来判断文件是否是发送邮箱的邮件
                    {
                        sendDs = excelToDS(sendPath, 1);
                        DataRow dr = sendDs.Tables[0].Rows[0];//选择正个表的第一行为发送邮箱
                        emailNum = dr["邮箱"].ToString().Trim();//得到邮箱和密码
                        emailPwd = dr["密码"].ToString().Trim();
                        if (emailNum == "" || emailPwd == "")
                        {
                            MessageBox.Show("发送邮箱中账号或者密码为空,请重新选择！");
                            sendPathText.Text = "";
                        }
                    }
                    catch
                    {
                        MessageBox.Show("所选的发送邮箱文件格式不对，请重新选择！");
                        sendPathText.Text = "";
                    }
                }
                else
                {
                    MessageBox.Show("请选择发送邮箱文件，且其为Excel文件！");
                }
            }
        }
        #endregion

        #region  选择接收邮箱账号
        private void selectReceBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "选择文件";
            ofd.InitialDirectory = Application.StartupPath + "\\接收邮箱";//打开控件后，默认目录
            ofd.Filter = "所有文件(*.*)|*.*|Excle文件(*.xls)|*.xls|Excle文件(*.xlsx)|*.xlsx";//打开文件类型
            ofd.RestoreDirectory = true;//设置对话框是否记忆之前打开的目录
            ofd.Multiselect = false;//是否可以同时打开多个文件
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                rePath = ofd.FileName.ToString();//获得用户选择的文件完整路径
                if (Path.GetExtension(rePath) == ".xls" || Path.GetExtension(rePath) == ".xlsx")
                {
                    rePathText.Text = rePath;
                    try//判断文件是否是接收邮箱文件
                    {
                        emailDS = excelToDS(rePath, 2);
                        DataRow dr = emailDS.Tables[0].Rows[0];
                        string s = dr["电子信箱"].ToString().Trim();
                    }
                    catch
                    {
                        MessageBox.Show("所选的接收邮箱文件格式不对，请重新选择！");
                        rePathText.Text = "";
                    }
                }
                else
                {
                    MessageBox.Show("请选择接收邮箱文件，且其为Excel文件！");
                }
            }
        }
        #endregion

        #region  选择发送的excel文件
        private void selectSendExcelBtn_Click(object sender, EventArgs e)
        {
            if (excelType.SelectedIndex != -1)
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Title = "选择文件";
                ofd.InitialDirectory = Application.StartupPath + "\\发送文件";//打开控件后，默认目录
                ofd.Filter = "所有文件(*.*)|*.*|Excle文件(*.xls)|*.xls|Excle文件(*.xlsx)|*.xlsx";//打开文件类型
                ofd.RestoreDirectory = true;//设置对话框是否记忆之前打开的目录
                ofd.Multiselect = false;//是否可以同时打开多个文件
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    salaryPath = ofd.FileName.ToString();//获得用户选择的文件完整路径
                    if (Path.GetExtension(salaryPath) == ".xls" || Path.GetExtension(salaryPath) == ".xlsx")
                    {
                        salaryText.Text = salaryPath;
                        if (excelType.SelectedIndex == 0)
                        {
                            try//判断是否为工资明细
                            {
                                salaryDS = excelToDS(salaryPath, 2);
                                DataRow dr = salaryDS.Tables[0].Rows[0];
                                sendExcelCount = salaryDS.Tables[0].Rows.Count;
                                dt = salaryDS.Tables[0];
                                string s = dr["实领金额"].ToString().Trim();//三个判断Excel中列名是否对的上，一个也行，但是增加概率
                                string s1 = dr["岗位工资"].ToString().Trim();
                                string s2 = dr["薪级工资"].ToString().Trim();
                            }
                            catch
                            {
                                MessageBox.Show("所选的工资明细文件格式不对，请重新选择！");
                                salaryText.Text = "";
                                salaryDS = null;
                            }
                        }
                        else if (excelType.SelectedIndex == 1)
                        {

                            salaryDS = excelToDS(salaryPath, 3);
                            dt = salaryDS.Tables[0];
                            getTitleArr();
                            if (countTitle == 0)
                            {
                                MessageBox.Show("所选的发送文件为空，请重新选择！");
                                salaryText.Text = "";
                                salaryDS = null;
                            }
                            else
                            {
                                DataRow dr = salaryDS.Tables[0].Rows[0];
                                string s = dr[0].ToString().Trim();
                                if (s != "人员代码")
                                {
                                    MessageBox.Show("所选的文件的第一列应为”人员代码“请重新选择！");
                                    salaryDS = null;
                                    salaryText.Text = "";
                                }
                            }


                        }
                        else
                        {
                            MessageBox.Show("请选择发送文件类型");
                        }
                    }
                    else
                    {
                        MessageBox.Show("请选择工资明细文件，且其为Excel文件！");
                    }
                }
            }
            else
            {
                MessageBox.Show("请选择文件类型!");
            }
        }
        #endregion

        #region  获取excel中的头和列数据
        /// <summary>
        ///得到数据的头的列数
        /// </summary>
        private void getTitleArr()
        {
            DataRow dr = salaryDS.Tables[0].Rows[0];
            for (countTitle = 0; countTitle < title.Length; countTitle++)
            {
                try
                {
                    title[countTitle] = dr[countTitle].ToString().Trim();
                }
                catch
                {
                    break;
                }
            }
        }
        #endregion

        #region  开始发送邮件
        private void sendBtn_Click(object sender, EventArgs e)
        {
            list1.Clear();
            list2.Clear();
            list3.Clear();
            list4.Clear();
            list5.Clear();
            if (emailNum == "")
            {
                MessageBox.Show("请选择发送邮箱文件");
            }
            else if (emailDS == null)
            {
                MessageBox.Show("请选择接收邮箱文件");
            }
            else if (excelType.SelectedIndex == -1)
            {
                MessageBox.Show("请选择发送文件类型");
            }
            else if (salaryDS == null)
            {
                MessageBox.Show("请选择发送文件");
            }
            else
            {
                if ((th1 == null && th2 == null && th3 == null && th4 == null && th5 == null && th6 == null) || (th1.IsAlive == false && th2.IsAlive == false && th3.IsAlive == false && th4.IsAlive == false && th5.IsAlive == false && th6.IsAlive == false))
                {
                    startThOneToThFive();//启动一到五的线程
                    if (excelType.SelectedIndex == 1) //无限制条件下，插入失败的数据
                        th6 = new Thread(insertFailInfoAll);
                    else if (excelType.SelectedIndex == 0) //发送文件为工资明细数据时，插入失败的数据
                        th6 = new Thread(insertFailInfoMoney);
                    th6.IsBackground = true;//将线程设置为后台线程
                    th6.Start();
                }
                else
                {
                    MessageBox.Show("请勿操作，正在发送中....");
                }
            }
        }
        #endregion

        #region  开启所有的线程
        private void startThOneToThFive()
        {
            createLog();
            dtStart = DateTime.Now;
            //创建一个线程去执行这个方法
            th1 = new Thread(sendOne);
            th1.IsBackground = true;//将线程设置为后台线程
            th1.Start();
            th2 = new Thread(sendTwo);
            th2.IsBackground = true;//将线程设置为后台线程
            th2.Start();
            th3 = new Thread(sendThree);
            th3.IsBackground = true;//将线程设置为后台线程
            th3.Start();
            th4 = new Thread(sendFour);
            th4.IsBackground = true;//将线程设置为后台线程
            th4.Start();
            th5 = new Thread(sendFive);
            th5.IsBackground = true;//将线程设置为后台线程
            th5.Start();
        }
        #endregion

        #region 取消发送
        private void cancel_Click(object sender, EventArgs e)
        {

            if (th1 != null && th1.IsAlive == true)
            {
                th1.Abort();
            } if (th2 != null && th2.IsAlive == true)
            {
                th2.Abort();
            } if (th3 != null && th3.IsAlive == true)
            {
                th3.Abort();
            } if (th4 != null && th4.IsAlive == true)
            {
                th4.Abort();
            } if (th5 != null && th5.IsAlive == true)
            {
                th5.Abort();
            }
        }
        #endregion

        #region  页面初始化
        private void restart()
        {
            sendPathText.Text = "";
            rePathText.Text = "";
            salaryText.Text = "";
            excelType.SelectedIndex = -1;
            emailNum = "";
            emailPwd = "";
            list1.Clear();
            list2.Clear();
            list3.Clear();
            list4.Clear();
            list5.Clear();
            infoOne.Text = infoTwo.Text = infoThree.Text = infoFour.Text = infoFive.Text = "总共发送了0条信息，已发送成功了0条信息";
            salaryDS = null;//获取工资表数据
            emailDS = null;//邮箱数据
            wrong = "";//错误代码
            title = new string[100];  //无限制条件下，获取待发送excel文件第一行数据
        }
        #endregion

        #region 加载页面状态
        /// <summary>
        /// //页面加载时
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {

            Control.CheckForIllegalCrossThreadCalls = false; //取消跨线程的访问
            eBody.Text = "老师您好:\n人事处给您的工资明细清单如下：";
        }
        #endregion

        #region  逐一发送邮件
        private void sendEmail(string body, string jobNum, int numTag, int sendExcelNum, List<int> list)
        {
            if (body != "")
            {
                try
                {
                    SmtpClient smtp = new SmtpClient();//实例化一个SmtpClient
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network; //将smtp的出站方式设为 Network
                    int index = emailNum.IndexOf("@");
                    if (emailNum.Substring(index) == "@jxnu.edu.cn")
                    {
                        smtp.Host = "mail.jxnu.edu.cn"; //指定 smtp 服务器地址
                    }
                    else
                    {
                        smtp.EnableSsl = true;//smtp服务器是否启用SSL加密 咱们这个不用开启这个 但是如果用其他邮箱 比如qq就要开启 否则会出错
                        smtp.Host = "smtp." + emailNum.Substring(index + 1);
                    }
                    smtp.Port = 25;             //指定smtp服务器的端口，默认是25，如果采用默认端口，可省去
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = new NetworkCredential(emailNum, emailPwd);//发送邮箱和密码
                    string emailReNum = "";//接收邮箱
                    for (int j = 0; j < emailDS.Tables[0].Rows.Count; j++)//取第一个
                    {
                        DataRow dr = emailDS.Tables[0].Rows[j];
                        if (dr["工号"].ToString().Trim() == jobNum)
                        {
                            emailReNum = dr["电子信箱"].ToString().Trim();
                            break;
                        }
                    }
                    MailMessage mm = new MailMessage(); //实例化一个邮件类
                    mm.Priority = MailPriority.High; //邮件的优先级，分为 Low, Normal, High，通常用 Normal即可
                    mm.From = new MailAddress(emailNum, "江西师范大学人事处", Encoding.GetEncoding(936));//收件方看到的邮件来源//第一个参数是发信人邮件地址//第二参数是发信人显示的名称//第三个参数是 第二个参数所使用的编码，如果指定不正确，则对方收到后显示乱码
                    //mm.CC.Add("17770426925@163.com"); //邮件的抄送者，支持群发，多个邮件地址之间用 半角逗号 分开
                    //mm.Bcc.Add("17770426925@163.com");//邮件的密送者，支持群发，多个邮件地址之间用 半角逗号 分开
                    mm.To.Add(emailReNum);
                    mm.Subject = "工资发放"; //邮件标题
                    mm.SubjectEncoding = Encoding.GetEncoding(936);//这里非常重要，如果你的邮件标题包含中文，这里一定要指定，否则对方收到的极有可能是乱码。//936是简体中文的pagecode，如果是英文标题，这句可以忽略不用
                    mm.IsBodyHtml = true; //邮件正文是否是HTML格式
                    mm.BodyEncoding = Encoding.GetEncoding(936); //邮件正文的编码， 设置不正确， 接收者会收到乱码
                    mm.Body = body;//邮件正文
                    //mm.Attachments.Add(new Attachment(@"C:\Users\Administrator\Desktop\1.txt", System.Net.Mime.MediaTypeNames.Application.Pdf));//添加附件，第二个参数，表示附件的文件类型，可以不用指定//可以添加多个附件
                    //mm.Attachments.Add(new Attachment(@"d:b.doc"));
                    smtp.Send(mm); //发送邮件，如果不返回异常， 则大功告成了。
                    switch (numTag)
                    {
                        case 1: num1++;
                            break;
                        case 2: num2++;
                            break;
                        case 3: num3++;
                            break;
                        case 4: num4++;
                            break;
                        case 5: num5++;
                            break;
                    }

                }
                catch (Exception ee)
                {
                    list.Add(sendExcelNum);
                    if (wrong == "")
                        wrong = ee.ToString();
                }
            }
            else
            {
                list.Add(sendExcelNum);
            }
        }
        #endregion

        #region  插入发送失败信息
        /// <summary>
        /// //插入发送失败的人的信息
        /// </summary>
        private void insertFailInfoMoney()
        {
            string fileName = Application.StartupPath + "\\发送文件\\工资明细发送失败的信息.xlsx";
            try
            {
                if (File.Exists(fileName))//如有，则删了去
                    File.Delete(fileName);
            }
            catch { }
            while (true)
            {
                if (th1.IsAlive == false && th2.IsAlive == false && th3.IsAlive == false && th4.IsAlive == false && th5.IsAlive == false)
                {
                    List<int> list = new List<int>();
                    list.AddRange(list1);
                    list.AddRange(list2);
                    list.AddRange(list3);
                    list.AddRange(list4);
                    list.AddRange(list5);
                    list.Sort();
                    writeLog(list.Count);
                    if (list.Count > 0)
                    {
                        object Nothing = System.Reflection.Missing.Value;
                        var app = new Microsoft.Office.Interop.Excel.Application();
                        app.Visible = false;
                        Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Add(Nothing);
                        Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets[1];
                        worksheet.Name = "发送失败信息";
                        //Microsoft.Office.Interop.Excel.Range contentRange = worksheet.get_Range(worksheet.Cells[3, i + 1], worksheet.Cells[rowCount - 1 + 3, i + 1]);
                        //contentRange.NumberFormatLocal = "@";//文本格式
                        worksheet.Cells[1, 1] = "人员代码";//标题
                        worksheet.Cells[1, 2] = "年";
                        worksheet.Cells[1, 3] = "月";//标题
                        worksheet.Cells[1, 4] = "部门名称";
                        worksheet.Cells[1, 5] = "姓名";//标题
                        worksheet.Cells[1, 6] = "岗位工资";
                        worksheet.Cells[1, 7] = "薪级工资";//标题
                        worksheet.Cells[1, 8] = "工资差额";
                        worksheet.Cells[1, 9] = "基础绩效";//标题
                        worksheet.Cells[1, 10] = "奖励绩效";
                        worksheet.Cells[1, 11] = "减少绩效";//标题
                        worksheet.Cells[1, 12] = "交通补贴";
                        worksheet.Cells[1, 13] = "离退休费";//标题
                        worksheet.Cells[1, 14] = "遗属补助";
                        worksheet.Cells[1, 15] = "退休补贴";//标题
                        worksheet.Cells[1, 16] = "教护龄";
                        worksheet.Cells[1, 17] = "提高百分之十";//标题
                        worksheet.Cells[1, 18] = "保留补贴";
                        worksheet.Cells[1, 19] = "独子女费";//标题
                        worksheet.Cells[1, 20] = "女卫生费";
                        worksheet.Cells[1, 21] = "校内补贴";
                        worksheet.Cells[1, 22] = "政策补贴";
                        worksheet.Cells[1, 23] = "住房房贴";
                        worksheet.Cells[1, 24] = "特殊补贴";
                        worksheet.Cells[1, 25] = "护理费";
                        worksheet.Cells[1, 26] = "长寿津贴";
                        worksheet.Cells[1, 27] = "厅级补贴";
                        worksheet.Cells[1, 28] = "机动收入";
                        worksheet.Cells[1, 29] = "收入合计";
                        worksheet.Cells[1, 30] = "会费";
                        worksheet.Cells[1, 31] = "公积金";
                        worksheet.Cells[1, 32] = "失业保险";
                        worksheet.Cells[1, 33] = "社会保险";
                        worksheet.Cells[1, 34] = "养老保险";
                        worksheet.Cells[1, 35] = "医疗保险";
                        worksheet.Cells[1, 36] = "家俱房租";
                        worksheet.Cells[1, 37] = "水费";
                        worksheet.Cells[1, 38] = "电费";
                        worksheet.Cells[1, 39] = "收视费";
                        worksheet.Cells[1, 40] = "互助金";
                        worksheet.Cells[1, 41] = "电汇";
                        worksheet.Cells[1, 42] = "机动扣款";
                        worksheet.Cells[1, 43] = "考勤扣款";
                        worksheet.Cells[1, 44] = "其他扣款";
                        worksheet.Cells[1, 45] = "所得税";
                        worksheet.Cells[1, 46] = "支出合计";
                        worksheet.Cells[1, 47] = "实领金额";
                        worksheet.Cells[1, 48] = "备注";
                        for (int i = 2; i < list.Count + 2; i++)
                        {
                            worksheet.Cells[i, 1] = "'" + dt.Rows[list[i - 2]][0].ToString().Trim();//标题
                            worksheet.Cells[i, 2] = dt.Rows[list[i - 2]][1].ToString().Trim();
                            worksheet.Cells[i, 3] = dt.Rows[list[i - 2]][2].ToString().Trim();
                            worksheet.Cells[i, 4] = dt.Rows[list[i - 2]][3].ToString().Trim();
                            worksheet.Cells[i, 5] = dt.Rows[list[i - 2]][4].ToString().Trim();
                            worksheet.Cells[i, 6] = dt.Rows[list[i - 2]][5].ToString().Trim();
                            worksheet.Cells[i, 7] = dt.Rows[list[i - 2]][6].ToString().Trim();
                            worksheet.Cells[i, 8] = dt.Rows[list[i - 2]][7].ToString().Trim();
                            worksheet.Cells[i, 9] = dt.Rows[list[i - 2]][8].ToString().Trim();
                            worksheet.Cells[i, 10] = dt.Rows[list[i - 2]][9].ToString().Trim();
                            worksheet.Cells[i, 11] = dt.Rows[list[i - 2]][10].ToString().Trim();
                            worksheet.Cells[i, 12] = dt.Rows[list[i - 2]][11].ToString().Trim();
                            worksheet.Cells[i, 13] = dt.Rows[list[i - 2]][12].ToString().Trim();
                            worksheet.Cells[i, 14] = dt.Rows[list[i - 2]][13].ToString().Trim();
                            worksheet.Cells[i, 15] = dt.Rows[list[i - 2]][14].ToString().Trim();
                            worksheet.Cells[i, 16] = dt.Rows[list[i - 2]][15].ToString().Trim();
                            worksheet.Cells[i, 17] = dt.Rows[list[i - 2]][16].ToString().Trim();
                            worksheet.Cells[i, 18] = dt.Rows[list[i - 2]][17].ToString().Trim();
                            worksheet.Cells[i, 19] = dt.Rows[list[i - 2]][18].ToString().Trim();
                            worksheet.Cells[i, 20] = dt.Rows[list[i - 2]][19].ToString().Trim();
                            worksheet.Cells[i, 21] = dt.Rows[list[i - 2]][20].ToString().Trim();
                            worksheet.Cells[i, 22] = dt.Rows[list[i - 2]][21].ToString().Trim();
                            worksheet.Cells[i, 23] = dt.Rows[list[i - 2]][22].ToString().Trim();
                            worksheet.Cells[i, 24] = dt.Rows[list[i - 2]][23].ToString().Trim();
                            worksheet.Cells[i, 25] = dt.Rows[list[i - 2]][24].ToString().Trim();
                            worksheet.Cells[i, 26] = dt.Rows[list[i - 2]][25].ToString().Trim();
                            worksheet.Cells[i, 27] = dt.Rows[list[i - 2]][26].ToString().Trim();
                            worksheet.Cells[i, 28] = dt.Rows[list[i - 2]][27].ToString().Trim();
                            worksheet.Cells[i, 29] = dt.Rows[list[i - 2]][28].ToString().Trim();
                            worksheet.Cells[i, 30] = dt.Rows[list[i - 2]][29].ToString().Trim();
                            worksheet.Cells[i, 31] = dt.Rows[list[i - 2]][30].ToString().Trim();
                            worksheet.Cells[i, 32] = dt.Rows[list[i - 2]][31].ToString().Trim();
                            worksheet.Cells[i, 33] = dt.Rows[list[i - 2]][32].ToString().Trim();
                            worksheet.Cells[i, 34] = dt.Rows[list[i - 2]][33].ToString().Trim();
                            worksheet.Cells[i, 35] = dt.Rows[list[i - 2]][34].ToString().Trim();
                            worksheet.Cells[i, 36] = dt.Rows[list[i - 2]][35].ToString().Trim();
                            worksheet.Cells[i, 37] = dt.Rows[list[i - 2]][36].ToString().Trim();
                            worksheet.Cells[i, 38] = dt.Rows[list[i - 2]][37].ToString().Trim();
                            worksheet.Cells[i, 39] = dt.Rows[list[i - 2]][38].ToString().Trim();
                            worksheet.Cells[i, 40] = dt.Rows[list[i - 2]][39].ToString().Trim();
                            worksheet.Cells[i, 41] = dt.Rows[list[i - 2]][40].ToString().Trim();
                            worksheet.Cells[i, 42] = dt.Rows[list[i - 2]][41].ToString().Trim();
                            worksheet.Cells[i, 43] = dt.Rows[list[i - 2]][42].ToString().Trim();
                            worksheet.Cells[i, 44] = dt.Rows[list[i - 2]][43].ToString().Trim();
                            worksheet.Cells[i, 45] = dt.Rows[list[i - 2]][44].ToString().Trim();
                            worksheet.Cells[i, 46] = dt.Rows[list[i - 2]][45].ToString().Trim();
                            worksheet.Cells[i, 47] = dt.Rows[list[i - 2]][46].ToString().Trim();
                            worksheet.Cells[i, 48] = dt.Rows[list[i - 2]][47].ToString().Trim();
                        }
                        try
                        {
                            worksheet.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
                            workBook.Close(false, Type.Missing, Type.Missing);
                            app.Quit();
                        }
                        catch
                        {
                            //MessageBox.Show("请关闭已打开的发送失败的信息.xlsx");
                        }
                    }
                    MessageBox.Show("操作结束，请核对");
                    restart();
                    try
                    {
                        th6.Abort();
                    }
                    catch { }
                }
            }
        }
        private void insertFailInfoAll()
        {
            string fileName = Application.StartupPath + "\\发送文件\\发送失败的信息.xlsx";
            try
            {
                if (File.Exists(fileName))
                    File.Delete(fileName);
            }
            catch
            {
                //MessageBox.Show("发送失败的信息.xls被打开，请关闭！");
            }
            while (true)
            {
                if (th1.IsAlive == false && th2.IsAlive == false && th3.IsAlive == false && th4.IsAlive == false && th5.IsAlive == false)
                {
                    List<int> list = new List<int>();
                    list.AddRange(list1);
                    list.AddRange(list2);
                    list.AddRange(list3);
                    list.AddRange(list4);
                    list.AddRange(list5);
                    list.Sort();
                    writeLog(list.Count);
                    if (list.Count > 0)
                    {
                        object Nothing = System.Reflection.Missing.Value;
                        var app = new Microsoft.Office.Interop.Excel.Application();
                        app.Visible = false;
                        Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Add(Nothing);
                        Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets[1];
                        worksheet.Name = "发送失败信息";
                        //Microsoft.Office.Interop.Excel.Range contentRange = worksheet.get_Range(worksheet.Cells[3, i + 1], worksheet.Cells[rowCount - 1 + 3, i + 1]);
                        //contentRange.NumberFormatLocal = "@";//文本格式
                        for (int j = 1; j <= countTitle; j++)
                        {
                            worksheet.Cells[1, j] = title[j - 1];//标题
                        }
                        for (int i = 2; i < list.Count + 2; i++)
                        {
                            for (int j = 1; j <= countTitle; j++)
                            {
                                worksheet.Cells[i, j] = "'" + dt.Rows[list[i - 2]][j - 1].ToString().Trim();//标题
                            }
                        }
                        try
                        {
                            worksheet.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
                            workBook.Close(false, Type.Missing, Type.Missing);
                            app.Quit();
                        }
                        catch
                        {
                            //MessageBox.Show("请关闭已打开的发送失败的信息.xlsx");
                        }
                    }
                    MessageBox.Show("操作结束，请核对");
                    restart();
                    th6.Abort();
                }
            }
        }
        #endregion

        #region 发送全部邮件
        /// <summary>
        /// //发送全部邮件
        /// </summary>
        private void sendOne()
        {
            string body = "";
            string jobNum = "";
            num1 = 0;//线程1成功发送的个数
            int i;
            if (excelType.SelectedIndex == 1)
                i = 1;
            else
                i = 0;
            for (; i < salaryDS.Tables[0].Rows.Count / 5; i++)
            {
                if (excelType.SelectedIndex == 0)
                    body = getInfoMoney(i);
                else if (excelType.SelectedIndex == 1)
                    body = getInfoAll(i);
                jobNum = getBaseInfo(i);
                sendEmail(body, jobNum, 1, i, list1);
                infoOne.Text = "总共发送了" + (i + 1) + "条信息，已发送成功了" + num1 + "条信息";
                Thread.Sleep(1000 * 1);
            }
        }
        private void sendTwo()
        {
            string body = "";
            string jobNum = "";
            num2 = 0;//线程2成功发送的个数
            int count = 0;//线程2总共发送的个数
            int i;
            if (excelType.SelectedIndex == 1)
                i = salaryDS.Tables[0].Rows.Count / 5 + 1;
            else
                i = salaryDS.Tables[0].Rows.Count / 5;
            for (; i < salaryDS.Tables[0].Rows.Count / 5 * 2; i++)
            {
                if (excelType.SelectedIndex == 0)
                    body = getInfoMoney(i);
                else if (excelType.SelectedIndex == 1)
                    body = getInfoAll(i);
                jobNum = getBaseInfo(i);
                sendEmail(body, jobNum, 2, i, list2);
                count++;
                infoTwo.Text = "总共发送了" + count + "条信息，已发送成功了" + num2 + "条信息";
                Thread.Sleep(1000 * 1);
            }
        }
        private void sendThree()
        {
            string body = "";
            string jobNum = "";
            num3 = 0;//线程3成功发送的个数
            int count = 0;//线程3总共发送的个数
            int i;
            if (excelType.SelectedIndex == 1)
                i = salaryDS.Tables[0].Rows.Count / 5 * 2 + 1;
            else
                i = salaryDS.Tables[0].Rows.Count / 5 * 2;
            for (; i < salaryDS.Tables[0].Rows.Count / 5 * 3; i++)
            {
                if (excelType.SelectedIndex == 0)
                    body = getInfoMoney(i);
                else if (excelType.SelectedIndex == 1)
                    body = getInfoAll(i);
                jobNum = getBaseInfo(i);
                sendEmail(body, jobNum, 3, i, list3);
                count++;
                infoThree.Text = "总共发送了" + count + "条信息，已发送成功了" + num3 + "条信息";
                Thread.Sleep(1000 * 1);
            }
        }
        private void sendFour()
        {
            string body = "";
            string jobNum = "";
            num4 = 0;//线程4成功发送的个数
            int count = 0;//线程4总共发送的个数
            int i;
            if (excelType.SelectedIndex == 1)
                i = salaryDS.Tables[0].Rows.Count / 5 * 3 + 1;
            else
                i = salaryDS.Tables[0].Rows.Count / 5 * 3;
            for (; i < salaryDS.Tables[0].Rows.Count / 5 * 4; i++)
            {
                if (excelType.SelectedIndex == 0)
                    body = getInfoMoney(i);
                else if (excelType.SelectedIndex == 1)
                    body = getInfoAll(i);
                jobNum = getBaseInfo(i);
                sendEmail(body, jobNum, 4, i, list4);
                count++;
                infoFour.Text = "总共发送了" + count + "条信息，已发送成功了" + num4 + "条信息";
                Thread.Sleep(1000 * 1);
            }
        }
        private void sendFive()
        {
            string body = "";
            string jobNum = "";
            num5 = 0;//线程5成功发送的个数
            int count = 0;//线程5总共发送的个数
            int i;
            if (excelType.SelectedIndex == 1)
                i = salaryDS.Tables[0].Rows.Count / 5 * 4 + 1;
            else
                i = salaryDS.Tables[0].Rows.Count / 5 * 4;
            for (; i < salaryDS.Tables[0].Rows.Count; i++)
            {
                if (excelType.SelectedIndex == 0)
                    body = getInfoMoney(i);
                else if (excelType.SelectedIndex == 1)
                    body = getInfoAll(i);
                jobNum = getBaseInfo(i);
                sendEmail(body, jobNum, 5, i, list5);
                count++;
                infoFive.Text = "总共发送了" + count + "条信息，已发送成功了" + num5 + "条信息";
                Thread.Sleep(1000 * 1);
            }
        }
        #endregion

        #region  获取需要导入的excel表
        private DataSet excelToDS(string path, int check)
        {
            string strConn = "";
            if (check == 3)  //表示发送文件为工资明细文件格式，得到数据集第二行数据，不包含第一行标题数据
                strConn = getStrConnThree(path);
            else
                strConn = getStrConn(path);   //对发送文件无文件格式限制
            DataSet ds = null;
            using (OleDbConnection conn = new OleDbConnection(strConn))//使用指定的连接字符串初始化 OleDbConnection 类的新实例
            {
                conn.Open();//打开连接
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                System.Data.DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString().Trim();
                strExcel = "select * from [" + sheetName + "]";
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                ds = new DataSet();
                myCommand.Fill(ds, "[Sheet1$]");
            }
            return ds;
        }
        #endregion

        #region 得到对应的excel后缀符号，如.xls和.xlsx的字符串
        private string getStrConn(string file)
        {
            string fileType = Path.GetExtension(file);
            if (fileType == ".xls")
                //HDR=Yes，这代表第一行是标题，不做为数据使用 ，如果用HDR=NO，则表示第一行不是标题，做为数据来使用。系统默认的是YES
                return "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + file + ";" + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=0;\"";//.xls
            else
                return "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + file + ";" + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=0;\"";//.xlsx

        }
        private string getStrConnThree(string file)
        {
            string fileType = Path.GetExtension(file);
            if (fileType == ".xls")
                //HDR=Yes，这代表第一行是标题，不做为数据使用 ，如果用HDR=NO，则表示第一行不是标题，做为数据来使用。系统默认的是YES
                return "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + file + ";" + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1;\"";//.xls
            else
                return "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + file + ";" + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1;\"";//.xlsx

        }
        #endregion

        #region 获取工号信息
        private string getBaseInfo(int lineNum)
        {
            string jobNum = "";
            DataRow dr = salaryDS.Tables[0].Rows[lineNum];
            if (excelType.SelectedIndex == 0)
                jobNum = dr["人员代码"].ToString().Trim();
            else
                jobNum = dr[0].ToString().Trim();
            return jobNum;
        }
        #endregion

        #region 获取工资明细数据
        /// <summary>
        /// //提取工资数据
        /// </summary>
        /// <param name="salaryDS">薪资的DataSet</param>
        /// <param name="lineNum"></param>
        /// <returns>提取工资数据</returns>
        private string getInfoMoney(int lineNum)
        {
            string body = "";
            try
            {
                DataRow dr = salaryDS.Tables[0].Rows[lineNum];//高校教职工工资明细清单
                body = dr["姓名"].ToString().Trim() + eBody.Text + "<br>";
                body = body + "<table border=1 cellspacing=0 width=100% bordercolorlight=#333333 bordercolordark=#efefef><tr><td>" + "人员代码：" + dr["人员代码"].ToString().Trim() + "</td>";
                body = body + "<td>" + "年：" + dr["年"].ToString().Trim() + "</td>";
                body = body + "<td>" + "月：" + dr["月"].ToString().Trim() + "</td>";
                body = body + "<td>" + "部门名称：" + dr["部门名称"].ToString().Trim() + "</td>";
                body = body + "<td>" + "姓名：" + dr["姓名"].ToString().Trim() + "</td>";
                body = body + "<td align='center'>" + "实领金额：" + dr["实领金额"].ToString().Trim() + "</td></tr><tr>";
                body = body + "<td>" + "岗位工资：" + dr["岗位工资"].ToString().Trim() + "</td>";
                body = body + "<td>" + "薪级工资：" + dr["薪级工资"].ToString().Trim() + "</td>";
                body = body + "<td>" + "工资差额：" + dr["工资差额"].ToString().Trim() + "</td>";
                body = body + "<td>" + "基础绩效：" + dr["基础绩效"].ToString().Trim() + "</td>";
                body = body + "<td>" + "奖励绩效：" + dr["奖励绩效"].ToString().Trim() + "</td>";
                body = body + "<td rowspan='5' align='center'>" + "收入合计：" + dr["收入合计"].ToString().Trim() + "</td></tr><tr>";
                body = body + "<td>" + "减少绩效：" + dr["减少绩效"].ToString().Trim() + "</td>";
                body = body + "<td>" + "交通补贴：" + dr["交通补贴"].ToString().Trim() + "</td>";
                body = body + "<td>" + "离退休费：" + dr["离退休费"].ToString().Trim() + "</td>";
                body = body + "<td>" + "遗属补助：" + dr["遗属补助"].ToString().Trim() + "</td>";
                body = body + "<td>" + "退休补贴：" + dr["退休补贴"].ToString().Trim() + "</td></tr><tr>";
                body = body + "<td>" + "教护龄：" + dr["教护龄"].ToString().Trim() + "</td>";
                body = body + "<td>" + "提高百分之十：" + dr["提高百分之十"].ToString().Trim() + "</td>";
                body = body + "<td>" + "保留补贴：" + dr["保留补贴"].ToString().Trim() + "</td>";
                body = body + "<td>" + "独子女费：" + dr["独子女费"].ToString().Trim() + "</td>";
                body = body + "<td>" + "女卫生费：" + dr["女卫生费"].ToString().Trim() + "</td></tr><tr>";
                body = body + "<td>" + "校内补贴：" + dr["校内补贴"].ToString().Trim() + "</td>";
                body = body + "<td>" + "政策补贴：" + dr["政策补贴"].ToString().Trim() + "</td>";
                body = body + "<td>" + "住房房贴：" + dr["住房房贴"].ToString().Trim() + "</td>";
                body = body + "<td>" + "特殊补贴：" + dr["特殊补贴"].ToString().Trim() + "</td>";
                body = body + "<td>" + "护理费：" + dr["护理费"].ToString().Trim() + "</td></tr><tr>";
                body = body + "<td>" + "长寿津贴：" + dr["长寿津贴"].ToString().Trim() + "</td>";
                body = body + "<td>" + "厅级补贴：" + dr["厅级补贴"].ToString().Trim() + "</td>";
                body = body + "<td>" + "机动收入：" + dr["机动收入"].ToString().Trim() + "</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr>";
                body = body + "<td>" + "会费：" + dr["会费"].ToString().Trim() + "</td>";
                body = body + "<td>" + "公积金：" + dr["公积金"].ToString().Trim() + "</td>";
                body = body + "<td>" + "失业保险：" + dr["失业保险"].ToString().Trim() + "</td>";
                body = body + "<td>" + "社会保险：" + dr["社会保险"].ToString().Trim() + "</td>";
                body = body + "<td>" + "养老保险：" + dr["养老保险"].ToString().Trim() + "</td>";
                body = body + "<td rowspan='4' align='center'>" + "支出合计：" + dr["支出合计"].ToString().Trim() + "</td></tr><tr>";
                body = body + "<td>" + "医疗保险：" + dr["医疗保险"].ToString().Trim() + "</td>";
                body = body + "<td>" + "家俱房租：" + dr["家俱房租"].ToString().Trim() + "</td>";
                body = body + "<td>" + "水费：" + dr["水费"].ToString().Trim() + "</td>";
                body = body + "<td>" + "电费：" + dr["电费"].ToString().Trim() + "</td>";
                body = body + "<td>" + "收视费：" + dr["收视费"].ToString().Trim() + "</td></tr><tr>";
                body = body + "<td>" + "互助金：" + dr["互助金"].ToString().Trim() + "</td>";
                body = body + "<td>" + "电汇：" + dr["电汇"].ToString().Trim() + "</td>";
                body = body + "<td>" + "机动扣款：" + dr["机动扣款"].ToString().Trim() + "</td>";
                body = body + "<td>" + "考勤扣款：" + dr["考勤扣款"].ToString().Trim() + "</td>";
                body = body + "<td>" + "其他扣款：" + dr["其他扣款"].ToString().Trim() + "</td></tr><tr>";
                body = body + "<td>" + "所得税：" + dr["所得税"].ToString().Trim() + "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr>";
                body = body + "<td colspan='6'>" + "备注：" + dr["备注"].ToString().Trim() + "</td></tr></table>";
            }
            catch
            {
                body = "";
            }
            return body;
        }
        #endregion

        #region 提取工资数据
        private string getInfoAll(int lineNum)
        {
            string body = "";
            try
            {
                DataRow dr = salaryDS.Tables[0].Rows[lineNum];//高校教职工工资明细清单
                body = eBody.Text + "<br>";
                body = body + "<table border=1 cellspacing=0 width=100%><tr>";
                for (int j = 0; j < countTitle; j++)
                {
                    body = body + "<td width='100px'>" + title[j] + ":" + dr[j].ToString().Trim() + "</td>";
                    if ((j + 1) % 5 == 0)
                    {
                        body += "</tr><tr>";
                    }
                }
                body = body + "</tr></table>";
            }
            catch
            {
                body = "";
            }
            return body;
        }
        #endregion

        #region 更新日志文件
        private void writeLog(int count)
        {
            string filepath = Application.StartupPath + "\\日志文件\\日志文件.xlsx";
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks._Open(filepath, Missing.Value, Missing.Value,
                                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets[1];
            string dateDiff = null;//间隔的时间
            DateTime dtEnd = DateTime.Now;//发送完成结束的时间
            TimeSpan ts = dtStart.Subtract(dtEnd).Duration();
            dateDiff = ts.Days.ToString() + "天" + ts.Hours.ToString() + "小时" + ts.Minutes.ToString() + "分钟" + ts.Seconds.ToString() + "秒";
            worksheet.Cells[allCount + 2, 1] = emailNum;
            worksheet.Cells[allCount + 2, 2] = salaryPath;
            worksheet.Cells[allCount + 2, 3] = dtStart.ToString();
            worksheet.Cells[allCount + 2, 4] = dtEnd.ToString();
            worksheet.Cells[allCount + 2, 5] = dateDiff;
            worksheet.Cells[allCount + 2, 6] = sendExcelCount;
            worksheet.Cells[allCount + 2, 7] = num1 + num2 + num3 + num4 + num5;
            worksheet.Cells[allCount + 2, 8] = count;
            if (sendExcelCount != 0)
            {
                double rate = (num1 + num2 + num3 + num4 + num5) * 1.0 / sendExcelCount;
                worksheet.Cells[allCount + 2, 9] = rate;
            }
            worksheet.Cells[allCount + 2, 10] = wrong;
            try
            {
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false; //设置禁止弹出保存和覆盖的询问提示框   
                workBook.SaveAs(filepath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                workBook.Close(true, Type.Missing, Type.Missing);
                app.Quit();
                app = null;
            }
            catch
            {
                MessageBox.Show("日志文件更新不成功！");
            }
        }
        #endregion

        #region 创建日志文件
        private void createLog()
        {
            string fileName = Application.StartupPath + "\\日志文件\\日志文件.xlsx";
            if (!File.Exists(fileName))
            {
                allCount = 0;
                object Nothing = Missing.Value;
                var app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = false;
                Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Add(Nothing);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets[1];
                worksheet.Name = "日志文件";
                worksheet.Cells[1, 1] = "发送邮箱";
                worksheet.Cells[1, 2] = "工资明细文件";
                worksheet.Cells[1, 3] = "发送时间点";
                worksheet.Cells[1, 4] = "完成时间点";
                worksheet.Cells[1, 5] = "总共发送时间";
                worksheet.Cells[1, 6] = "总共次数";
                worksheet.Cells[1, 7] = "成功次数";//标题
                worksheet.Cells[1, 8] = "失败次数";
                worksheet.Cells[1, 9] = "成功率";
                worksheet.Cells[1, 10] = "错误代码                                       ";//10个tab
                try
                {
                    worksheet.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
                    workBook.Close(false, Type.Missing, Type.Missing);
                    app.Quit();
                    app = null;
                }
                catch
                {
                    MessageBox.Show("日志文件创建不成功");
                }
            }
            else
                allCount = excelToDS(fileName, 1).Tables[0].Rows.Count;
        }
        #endregion

        #region 线程结束，关闭窗体
        private void Form1_FormClosing_1(object sender, FormClosingEventArgs e)
        {
            if (th1 != null) //当你点击关闭窗体的时候，判断此线程是否为null
            {
                th1.Abort();//关闭线程
            }
            if (th2 != null)
            {
                th2.Abort();
            }
            if (th3 != null)
            {
                th3.Abort();
            }
            if (th4 != null)
            {
                th4.Abort();
            }
            if (th5 != null)
            {
                th5.Abort();
            }
            if (th6 != null)
            {
                th6.Abort();
            }
        }
        #endregion

    }
}