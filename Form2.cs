using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormsApplication7
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string url = "https://github.com/sokudon/MLSTRINGTOOL/blob/master/bin/Release/%E3%81%90%E3%82%8A%E3%81%BE%E3%81%99.exe?raw=true";
            string versions = "https://github.com/sokudon/MLSTRINGTOOL/blob/master/bin/Release/version.txt?raw=true";
            string logs = "https://github.com/sokudon/MLSTRINGTOOL/blob/master/bin/Release/%E3%82%8C%E3%81%82%E3%81%A9%E3%82%81.txt?raw=true";
            string html = updater(versions, 0, "");
            textBox1.Text = updater(logs, 0, "").Replace("\n","\r\n");
             string builddate = label1.Text.Substring(6, label1.Text.Length-6);
             DateTime z = DateTime.Parse(html);
             DateTime y = DateTime.Parse(builddate);
            TimeSpan x= z-y;
            int zzz=(int)x.TotalDays+(int)x.TotalHours + (int)x.TotalMinutes;
            if (zzz > 0)
            {
                if (MessageBox.Show(this, "最新版がギットレポジトリに存在します、ダウンロードしますか？", "最新BUID"+html, MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    html = html.Replace("/", "");
                    html = html.Replace(":", "");
                    string fileName = "millionlive_string_tool " + html + ".exe";
                    string OK = updater(url, 1, fileName);
                    if (OK == "OK")
                    {
                        System.Media.SystemSounds.Beep.Play();
                        MessageBox.Show(this,"最新版のダウンロードが完了しました");
                       
                    }
                }

            }
            else
            {
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show(this,"最新版です、更新の必要はありません。");
            }
        }

        string updater(string url, int mode,string fn)
        {
            string ss = "";

                System.Net.HttpWebRequest webreq =
                    (System.Net.HttpWebRequest)
                        System.Net.WebRequest.Create(url);

                System.Net.HttpWebResponse webres =
                    (System.Net.HttpWebResponse)webreq.GetResponse();

                System.Text.Encoding enc = Encoding.GetEncoding("UTF-8");
                System.IO.Stream st = webres.GetResponseStream();

                if (mode == 0)
                {
                    System.IO.StreamReader sr = new System.IO.StreamReader(st, enc);
                    string html = sr.ReadToEnd();
                    sr.Close();
                    ss= html;
                }
                else
                {

                    FileStream fs = new FileStream(fn, FileMode.Create, FileAccess.Write);
                    byte[] rd = new byte[2048];
                    int rs = 0;

                    while (true)
                    {
                        rs = st.Read(rd, 0, rd.Length);
                        if (rs == 0)
                        {
                            ss= "OK";
                            goto skkp;

                        }
                        fs.Write(rd, 0, rs);
                    }
                    skkp:
                    fs.Close();

                }

                st.Close();
                return ss;

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

    }
}