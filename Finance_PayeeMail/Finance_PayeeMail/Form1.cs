using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Threading;

namespace Finance_PayeeMail
{
    public partial class Form1 : Form
    {
        // 宣告全域變數
        Form2 f2 = new Form2();
        private string FilePath;
        public static string  from ;
        DataTable dataTable = new DataTable();
        public bool en = false;
        public static string textbox1_olddata = "";
        public static string textbox2_olddata = "";
        public static string textbox3_olddata = "";
        public static string textbox4_olddata = "";

        public Form1()
        {
            InitializeComponent();
            getDefuleData();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //textBox3.Text = "";
            //textBox3.Enabled = true;
            //textBox4.Text = "";
            textBox4.Enabled = true;
            //InputSender.Text = "";
            InputSender.Enabled = true;
            //InputMailServer.Text = "";
            //InputMailServer.Enabled = true;
            button1.Enabled = true;
            button2.Enabled = false;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            InputSender.Enabled = false;
            InputMailServer.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = true;
            from = InputSender.Text;
            setDefuleData();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            DataTable dataBuffer = new DataTable();
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select File";
            dialog.InitialDirectory = @"./";
            if(dialog.ShowDialog() == DialogResult.OK)
            {
                string FilePath = dialog.FileName;
                string ProviderName = "Microsoft.ACE.OLEDB.12.0;";
                string ExtendedString = "'Excel 8.0;";
                string HDR = "NO;";
                string IMEX = "0';";
                string connectString = "Data Source="+FilePath+ ";Provider="+ProviderName+ ";Extended Properties="+ ExtendedString+";HDR="+HDR+";IMEX="+IMEX;
                using (OleDbConnection Connect = new OleDbConnection(connectString))
                {
                    string queryString = "SELECT * FROM ["+ textBox4.Text.ToString()+"$]";
                    Connect.Open();
                    try
                    {
                        using (OleDbDataAdapter dr = new OleDbDataAdapter(queryString, Connect))
                        {
                            dr.Fill(dataBuffer);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("異常訊息:" + ex.Message, "異常訊息");
                    }
                }
            }
            dataTable = dataBuffer;
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = dataTable;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //textBox2.Text = (DateTime.Today.DayOfWeek.ToString());
            textBox2.Text = DateTime.Now.ToString("MMMM", new System.Globalization.CultureInfo("en-us")).Substring(0, 3);
            textBox2.Text = textBox2.Text + "," + DateTime.Now.ToString("dd,yyyy");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Title = "Select File";
            if(openFile.ShowDialog() == DialogResult.OK)
            {
                FilePath = openFile.FileName;
                webBrowser1.Navigate(FilePath);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 宣告區域變數
            int index = 0;
            string oldname = "";
            int num = 0;
            string body = "";
            int totalAmount = 0;

            // 宣告鏈結加入資料
            ArrayList Payee = new ArrayList();
            ArrayList Email = new ArrayList();
            ArrayList Amount = new ArrayList();
            ArrayList Description = new ArrayList();
            ArrayList Company = new ArrayList();

            // 發送無商家版本
            if (dataTable.Columns.Count == 4)
            {
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    Payee.Add(dataTable.Rows[i][0]);
                    Email.Add(dataTable.Rows[i][1]);
                    Amount.Add(dataTable.Rows[i][2]);
                    Description.Add(dataTable.Rows[i][3]);
                }

                // 混入信件資料
                foreach (string name in Payee)
                {
                    if (index == 0)
                    {
                        oldname = name;
                    }

                    if (oldname != name)
                    {
                        body = body.Replace("{{TotalAmount}}",totalAmount.ToString("#,0"));
                        webBrowser1.DocumentText = body;
                        SendMail(InputMailServer.Text, int.Parse(textBox3.Text), InputSender.Text, (Email[Convert.ToInt32(Payee.IndexOf(oldname))]).ToString(), textBox1.Text, body);
                        //Thread.Sleep(500);
                        oldname = name;
                        num = 0;
                        totalAmount = 0;
                    }

                    if (oldname == name)
                    {
                        if (num == 0)
                        {
                            // 讀取模板
                            body = getModel(FilePath);
                            body = body.Replace("{{Payee}}", name)
                                       .Replace("</b>,", "</b>.<p>")
                                       .Replace("{{Amount}}", Convert.ToInt32(Amount[index]).ToString("#,0"))
                                       .Replace("{{SendDate}}", textBox2.Text)
                                       .Replace("{{Dscr}}", Description[index].ToString())
                                       .Replace("</table>Finance Department", "</table><p><p>Finance Department")
                                       .Replace("Thanks", "<br>Thanks");
                            totalAmount = totalAmount + Convert.ToInt32(Amount[index]);
                            num++;
                        }
                        else
                        {
                            string addString = "<tr><td class='center'>" + name + "</td><td class='right'>" + Convert.ToInt32(Amount[index]).ToString("#,0") + "</td><td class='left'>" + Description[index] + "</td></tr>";
                            body = body.Insert(body.IndexOf("<!-- END PayeeDetails -->"), addString);
                            totalAmount = totalAmount + Convert.ToInt32(Amount[index]);
                            num++;
                        }
                    }
                    index++;
                }

                // 執行最後送信
                body = body.Replace("{{TotalAmount}}", totalAmount.ToString("#,0"));
                webBrowser1.DocumentText = body;
                SendMail(InputMailServer.Text, int.Parse(textBox3.Text), InputSender.Text, (Email[Convert.ToInt32(Payee.IndexOf(oldname))]).ToString(), textBox1.Text, body);
                //Thread.Sleep(500);
                MessageBox.Show("送信完成");
            }

            // 發送商家版本
            if(dataTable.Columns.Count == 5)
            {
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    Company.Add(dataTable.Rows[i][0]);
                    Payee.Add(dataTable.Rows[i][1]);
                    Email.Add(dataTable.Rows[i][2]);
                    Amount.Add(dataTable.Rows[i][3]);
                    Description.Add(dataTable.Rows[i][4]);
                }

                // 混入信件資料
                foreach (string name in Payee)
                {
                    if (index == 0)
                    {
                        oldname = name;
                    }

                    if (oldname != name)
                    {
                        body = body.Replace("{{TotalAmount}}", totalAmount.ToString("#,0"));
                        webBrowser1.DocumentText = body;
                        SendMail(InputMailServer.Text, int.Parse(textBox3.Text), InputSender.Text, (Email[Convert.ToInt32(Payee.IndexOf(oldname))]).ToString(), textBox1.Text, body);
                        //Thread.Sleep(500);
                        oldname = name;
                        num = 0;
                        totalAmount = 0;
                    }

                    if (oldname == name)
                    {
                        if (num == 0)
                        {
                            // 讀取模板
                            body = getModel(FilePath);
                            body = body.Replace("{{Payee}}", name)
                                       .Replace("</b>,", "</b>.<p>")
                                       .Replace("{{Amount}}", Convert.ToInt32(Amount[index]).ToString("#,0"))
                                       .Replace("{{SendDate}}", textBox2.Text)
                                       .Replace("{{Company}}", Company[index].ToString())
                                       .Replace("{{Dscr}}", Description[index].ToString())
                                       .Replace("</table>Finance Department", "</table><p><p>Finance Department")
                                       .Replace(".If", ".<br>If")
                                       .Replace("Thanks", "<br>Thanks");
                            totalAmount = totalAmount + Convert.ToInt32(Amount[index]);
                            num++;
                        }
                        else
                        {
                            string addString = "<tr><td class='center'>" + Company[index].ToString() + "</td><td class='right'>" + Convert.ToInt32(Amount[index]).ToString("#,0") + "</td><td class='left'>" + Description[index] + "</td></tr>";
                            body = body.Insert(body.IndexOf("<!-- END PayeeDetails -->"), addString);
                            totalAmount = totalAmount + Convert.ToInt32(Amount[index]);
                            num++;
                        }
                    }
                    index++;
                }

                // 執行最後送信
                body = body.Replace("{{TotalAmount}}", totalAmount.ToString("#,0"));
                webBrowser1.DocumentText = body;
                SendMail(InputMailServer.Text, int.Parse(textBox3.Text), InputSender.Text, (Email[Convert.ToInt32(Payee.IndexOf(oldname))]).ToString(), textBox1.Text, body);
                //Thread.Sleep(500);
                MessageBox.Show("送信完成");
            }
        }

        

        public void SendMail(string host,int port ,string from,string to,string title,string body)
        {
            MailMessage MailServer = new MailMessage();
            MailServer.From = new MailAddress(from);
            MailServer.To.Add(to);
            MailServer.Subject = title;
            MailServer.Body = body;
            MailServer.IsBodyHtml = true;
      
            SmtpClient smtp = new SmtpClient(host, port);
            //smtp.EnableSsl = true;
            //smtp.Credentials = new NetworkCredential("test@nextfortune.tw", "abcd1234");
            //smtp.Credentials = new NetworkCredential("h57082287@gmail.com", "Wujunting19980701");
            try
            {
                smtp.Send(MailServer);
                smtp.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show("送信失敗，發生錯誤如下:\n"+ex.ToString());
            }
        }

        public string getModel(string filePath)
        {
            string lines = "";
            string line = "";
            StreamReader sr = new StreamReader(filePath);
            while((line = sr.ReadLine()) != null)
            {
                lines += line;
            }
            return lines;
        }

        public void getDefuleData()
        {
            StreamReader sr = new StreamReader("config.txt");
            String line = "";
            string [] data = new string[4];
            int i = 0;
            while ((line = sr.ReadLine()) != null)
            {
                data[i] = line;
                i++;
            }
            InputMailServer.Text = data[0];
            textBox3.Text = data[1];
            InputSender.Text = data[2];
            textBox4.Text = data[3];
            sr.Close();
        }

        public void setDefuleData()
        {
            StreamWriter wr = new StreamWriter("config.txt");

            wr.WriteLine(InputMailServer.Text);
            wr.WriteLine(textBox3.Text);
            wr.WriteLine(InputSender.Text);
            wr.WriteLine(textBox4.Text);
            wr.Close();
        }
    }
}
