using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Facebook;
using HtmlAgilityPack;
using Microsoft.Win32;

//using HtmlDocument = System.Windows.Forms.HtmlDocument;

namespace Test
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
            notifyIcon1.Icon = this.Icon;
            notifyIcon1.Text = "MVT - AutoWeather";
        }

        private string url = "http://www.nchmf.gov.vn/web/vi-VN/81/Default.aspx";

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Interval = 1000*60*10;
            timer1.Enabled = true;

            RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            if (rk.GetValue("MVT-AutoWeather") == null || rk.GetValue("MVT-AutoWeather").ToString() != Application.ExecutablePath.ToString())
            {
                rk.SetValue("MVT-AutoWeather", Application.ExecutablePath.ToString());
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            string sHTML = GetCodeHTML(url);
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(sHTML);
            string s = GetDateFromWeb(doc);
            if (s != LoadLastUpdate())
            {
                string sContent = GetContentHtml(doc);
                string sUrl = createFile(sContent);
                webBrowser1.Url = new Uri(sUrl);
                Render(sUrl, Application.StartupPath + "\\out.jpg", new Rectangle(0, 0, 0, 0));
                //
                Thread.Sleep(1000);
                PostData(Application.StartupPath + "\\out.jpg");
                //
                UpdateLastUpdate(s);
                WriteLog("Post ok!");
            }
            timer1.Enabled = true;
        }

        public static void Render(string inputUrl, string outputPath, Rectangle crop)
        {
            WebBrowser wb = new WebBrowser();
            wb.ScrollBarsEnabled = false;
            wb.ScriptErrorsSuppressed = true;
            wb.Navigate(inputUrl);
            while (wb.ReadyState != WebBrowserReadyState.Complete)
            {
                Application.DoEvents();
            }
            wb.Width = wb.Document.Body.ScrollRectangle.Width;
            wb.Height = wb.Document.Body.ScrollRectangle.Height;
            using (Bitmap bitmap = new Bitmap(wb.Width, wb.Height))
            {
                wb.DrawToBitmap(bitmap, new Rectangle(0, 0, wb.Width, wb.Height));
                wb.Dispose();
                Rectangle rect = new Rectangle(crop.Left, crop.Top, wb.Width - crop.Width - crop.Left, wb.Height - crop.Height - crop.Top);
                Bitmap cropped = bitmap.Clone(rect, bitmap.PixelFormat);
                cropped.Save(outputPath, ImageFormat.Png);
            }
        }

        private string createFile(string scontent)
        {
            string sfile = Application.StartupPath + "\\out.html";
            StreamWriter writer = new StreamWriter(sfile, false);
            scontent = "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><LINK rel=\"stylesheet\" type=\"text/css\" href=\"http://www.nchmf.gov.vn/web/Portals/_Default/Skins/style.css\"></LINK><table width=\"600\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\"><tr><td>" + scontent + "</td></tr></table>";
            writer.Write(scontent);
            writer.Close();
            writer.Dispose();
            return sfile;
        }

        private static string GetContentHtml(HtmlAgilityPack.HtmlDocument doc)
        {
            List<HtmlNode> oNodes = null;
            oNodes = (from HtmlNode node in doc.DocumentNode.SelectNodes("//td")
                      where node.Name.ToLower() == "td"
                      && node.GetAttributeValue("id", "") == "_ctl1__ctl0_ModulePane_223"
                      //&& node.Attributes["class"] != null
                      //&& node.Attributes["class"].Value.StartsWith("img_")
                      select node).ToList();
            if (oNodes != null)
                foreach (HtmlNode node in oNodes)
                {
                    return node.InnerHtml;
                }
            return "";
        }

        private static string GetDateFromWeb(HtmlAgilityPack.HtmlDocument doc)
        {
            List<HtmlNode> oNodes = null;
            oNodes = (from HtmlNode node in doc.DocumentNode.SelectNodes("//td")
                      where node.Name == "td"
                      && node.GetAttributeValue("class", "") == "SubTitleNews_Special"
                      //&& node.Attributes["class"] != null
                      //&& node.Attributes["class"].Value.StartsWith("img_")
                      select node).ToList();
            if (oNodes != null)
                foreach (HtmlNode node in oNodes)
                {
                    string sDate = node.InnerText;
                    return GetDate(sDate);
                }
            return "";
        }

        private static string GetCodeHTML(string Url)
        {
            HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(Url);
            myRequest.Method = "GET";
            WebResponse myResponse = myRequest.GetResponse();
            StreamReader sr = new StreamReader(myResponse.GetResponseStream(), System.Text.Encoding.UTF8);
            string result = sr.ReadToEnd();
            sr.Close();
            myResponse.Close();
            return result;
        }

        private static string GetDate(string s)
        {
            Regex re = new Regex(@"\d{1,2}[hH].+\d{2}/\d{2}/\d{2,4}");
            Match match = re.Match(s);
            if(match.Success)
            {
                return match.Groups[0].Value.ToString();
            }
            return "";
        }

        private static string LoadLastUpdate()
        {
            try
            {
                string sFile = Application.StartupPath + "\\LastUpdate.txt";
                StreamReader reader = new StreamReader(sFile, true);
                string s = reader.ReadLine();
                reader.Close();
                reader.Dispose();
                return s;
            }
            catch (Exception e)
            {
                WriteLog("Lỗi: " + e.Message);
                return "";
                throw;
            }
        }

        private static void UpdateLastUpdate(string s)
        {
            try
            {
                string sFile = Application.StartupPath + "\\LastUpdate.txt";
                TextWriter file = new StreamWriter(sFile, false);
                file.Write(s);
                file.Close();
                file.Dispose();
            }
            catch (Exception e)
            {
                WriteLog("Lỗi: " + e.Message);
                throw;
            }
        }

        private void PostData(string sFile)
        {
            var fb = new FacebookClient("AAAEZA95pDXoABACHuViSkXFRUw72lZCy7F0ePDW6UWtjqODm76FoWZAxCH8GkBcTUsJ7eK0ZBu5z9ejDqnXyOu4BbMj7ku9nMqL5HXaiq4dIBpGJG6gl");
            var parameters = new Dictionary<string, object>();
            parameters["message"] = "MVT: Send from auto post...";
            parameters["source"] = new FacebookMediaObject
            {
                ContentType = "image/jpeg",
                FileName = Path.GetFileName(sFile)
            }.SetValue(File.ReadAllBytes(sFile));
            fb.PostAsync("me/photos", parameters);
        }

        private static void WriteLog(string s)
        {
            string sLogFile = Application.StartupPath + "\\log\\" + DateTime.Now.ToString("ddMMyyyy") + "_log.txt";
            TextWriter file = new StreamWriter(sLogFile, true);
            file.WriteLine("-> " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + " : " + s);
            file.Close();
        }

        private void showToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = true;
            this.WindowState=FormWindowState.Normal;
        }

        private void stopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Chương trình post dự báo thời tiết","MVT");
        }

        private void startToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(timer1.Enabled==false)
            {
                timer1.Enabled = true;
            }
        }

        private void exitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void stopToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (timer1.Enabled == true)
            {
                timer1.Enabled = false;
            }
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            this.Visible = true;
            this.WindowState = FormWindowState.Normal;
        }

        private void Main_MinimumSizeChanged(object sender, EventArgs e)
        {
            this.Visible = false;
            this.WindowState = FormWindowState.Minimized;
        }

        private void Main_SizeChanged(object sender, EventArgs e)
        {
            if(WindowState == FormWindowState.Minimized)
            {
                this.Visible = false;
                ShowInTaskbar = false;
            }
            else
            {
                this.Visible = true;
                ShowInTaskbar = true;
            }
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Visible = false;
            ShowInTaskbar = false;
        }
    }
}
