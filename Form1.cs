using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;

namespace WindowsFormsApplication7
{
    public partial class がしゃふぃるたー : Form
    {
        static byte settingver = 1;

        struct BorderTBL
        {
            public DateTime timer;
            public Int32 rank1;
            public Int32 rank2;
            public Int32 rank3;
            public Int32 rank4;
            public Int32 rank5;
            public Int32 rank6;
            public BorderTBL(DateTime now, Int32 j, Int32 k, Int32 l, Int32 m, Int32 n, Int32 o)
        {
            timer =now;
            rank1 = j;
            rank2 = k;
            rank3 = l;
            rank4 = m;
            rank5 = n;
            rank6 = o;
        }
        }

        
        public がしゃふぃるたー()
        {
            InitializeComponent();

            if (File.Exists(Application.StartupPath + "\\" + "setting"))
            {
                FileStream sr = new FileStream(Application.StartupPath + "\\" + "setting", FileMode.Open);
                if (sr.Length > 10)
                {
                    byte[] bs = new byte[sr.Length];
                    sr.Read(bs, 0, bs.Length);


                    if ((bs[7]-settingver) ==0)
                    {
                        if (((bs[0]) & 2) == 0)
                        {
                                radioButton1.Checked = true;
                        }
                        else{                        
                                radioButton3.Checked = true;
                        }
                        if (((bs[0]) & 1) != 0)
                        {
                            checkBox10.Checked = true;
                        }
                        if (((bs[0] >> 3) & 1) != 0)
                        {
                            checkBox1.Checked = true;
                        }
                        if (((bs[0] >> 4) & 1) != 0)
                        {
                            checkBox2.Checked = true;
                        }
                        if (((bs[0] >> 5) & 1) != 0)
                        {
                            checkBox3.Checked = true;
                        }
                        if (((bs[0] >> 5) & 1) != 0)
                        {
                            checkBox4.Checked = true;
                        }
                        if (((bs[0] >> 7) & 1) != 0)
                        {
                            checkBox5.Checked = true;
                        }
                        if (((bs[1]) & 1) != 0)
                        {
                            checkBox6.Checked = true;
                        }
                        if (((bs[1] >> 1) & 1) != 0)
                        {
                            checkBox7.Checked = true;
                        }
                        if (((bs[1] >> 2) & 1) != 0)
                        {
                           radioButton2.Checked = true;
                        }
                        if (((bs[1] >> 3) & 1) != 0)
                        {
                            panel1.Visible = true;
                        }
                        if (((bs[1] >> 4) & 1) != 0)
                        {
                            radioButton3.Checked = true;
                        }
                        if (((bs[1] >> 5) & 1) != 0)
                        {
                            checkBox8.Checked = true;
                        }
                        if (((bs[1] >> 6) & 1) != 0)
                        {
                            comboBox5.SelectedIndex = 1;
                        }



                        textBox3.Text = BitConverter.ToInt16(bs, 8).ToString();
                        textBox5.Text = BitConverter.ToInt16(bs, 5).ToString();

                        byte[] sb = new byte[bs.Length - 10];
                        Array.Copy(bs, 10, sb, 0, bs.Length - 10);
                        string s = Encoding.GetEncoding(65001).GetString(sb);
                        string[] ss = s.Split('\t');
                        if (ss.Length >= 6)
                        {
                            comboBox1.Text = ss[0];
                            comboBox2.Text = ss[1];
                            textBox4.Text = ss[2];
                            comboBox3.Text = ss[3];
                            comboBox4.Text = ss[4];
                            comboBox5.Text = ss[5];
                        }



                    }
                }
                sr.Close();
            }
        }


        private void がしゃふぃるたー_FormClosed(object sender, FormClosedEventArgs e)
        {
            FileStream sr = new FileStream(Application.StartupPath + "\\" + "setting", FileMode.Create);
            byte[] bs = new byte[10];
            int flag = 0;
            int flag2 = 0;

            if (checkBox10.Checked)
                  flag=1;

            if (radioButton3.Checked)
                    flag+=2;


            if (checkBox1.Checked)
                flag |= 1<<3;

            if (checkBox2.Checked)
                flag |= 1 << 4;

            if (checkBox3.Checked)
                flag |= 1 << 5;

            if (checkBox4.Checked)
                flag |= 1 << 6;

            if (checkBox5.Checked)
                flag |= 1 << 7;

            if (checkBox6.Checked)                
                flag2 |= 1;

            if (checkBox7.Checked)
                flag2 |= 1<<1;

            if (radioButton2.Checked)
                flag2 |= 1<<2;

            if (panel1.Visible == true)
                flag2 |= 1<<3;

            if (radioButton3.Checked)
                flag2 |= 1 << 4;


            if (checkBox8.Checked)
                flag2 |= 1 << 5;


            if (comboBox6.SelectedIndex>0)
                flag2 |= 1 << 6;

            bs[0] = (byte)flag;
            bs[1] = (byte)flag2;

            bs[7] = settingver;
            if (textBox3.Text == "") {
                textBox3.Text = "0";
            }

            bs[5] = (byte)(Convert.ToInt32(textBox5.Text) & 0xFF);
            bs[6] = (byte)(Convert.ToInt32(textBox5.Text) >> 8);

            bs[8] = (byte)(Convert.ToInt32(textBox3.Text) & 0xFF);
            bs[9] = (byte)(Convert.ToInt32(textBox3.Text) >>8) ;


            string s = comboBox1.Text + "\t" + comboBox2.Text + "\t" + textBox4.Text + "\t" + comboBox3.Text + "\t" + comboBox4.Text +"\t" + comboBox5.Text;
            byte[] sb = Encoding.GetEncoding(65001).GetBytes(s);

            sr.Write(bs, 0, bs.Length);
            sr.Write(sb, 0, sb.Length);

            sr.Close();


        }



        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox4.SelectedIndex > 2 && comboBox4.SelectedIndex < 6)
            {

                if (checkBox11.Checked == true)
                {
                    //textBox2.Text = boder_parser(textBox1.Text) + textBox2.Text;
                    string[] s = boder_parser(textBox1.Text).Split(("|").ToCharArray());
                    textBox2.Text = boder_parser(s[0]) + textBox2.Text;
                    if (s.Length > 1)
                    {
                        textBox6.Text = boder_parser(s[1]) + textBox6.Text;
                    }
                }
                else
                {

                    string[] s = boder_parser(textBox1.Text).Split(("|").ToCharArray());
                    textBox2.Text += boder_parser(s[0]);
                    if (s.Length > 1)
                    {
                        textBox6.Text += boder_parser(s[1]);
                    }
                }
            }
            else if (comboBox4.Text.Contains("収集")) {

                if (checkBox11.Checked == true)
                {    //textBox2.Text = boder_parser(textBox1.Text) + textBox2.Text;
                    string[] s = boder_parser(textBox1.Text).Split(("|").ToCharArray());
                    textBox2.Text = boder_parser(s[0]) + textBox2.Text;
                    if (s.Length > 1)
                    {
                        textBox6.Text = boder_parser(s[1]) + textBox6.Text;
                    }
                }
                else
                {
                    string[] s = boder_parser(textBox1.Text).Split(("|").ToCharArray());

                    textBox2.Text += boder_parser(s[0]);
                    if (s.Length > 1)
                    { 
                    textBox6.Text += boder_parser(s[1]);
                    }
                }
            }
            else if (comboBox4.Text.Contains("IMC"))
            {

                if (checkBox11.Checked == true)
                {
                    textBox2.Text = boder_parser(textBox1.Text) + textBox2.Text;
                }
                else
                {
                    textBox2.Text += boder_parser(textBox1.Text);
                }
            }
            else
            {
                textBox2.Text = boder_parser(textBox1.Text);
            }
        }


        private void button5_Click(object sender, EventArgs e)
        {
            if (comboBox4.SelectedIndex > 2 && comboBox4.SelectedIndex < 6)
            {
                if (checkBox11.Checked == true)
                {
                   textBox1.Text =  boder_parser(textBox2.Text)+ textBox1.Text;
                }
                else
                {
                    textBox1.Text += boder_parser(textBox2.Text);
                }
            }
            else if (comboBox4.Text.Contains("100/500/1200/3000収集"))
            {

                if (checkBox11.Checked == true)
                {
                    textBox1.Text = boder_parser(textBox2.Text) + textBox1.Text;
                }
                else
                {
                    textBox1.Text += boder_parser(textBox2.Text);
                }
            }
            else if (comboBox4.Text.Contains("IMC"))
            {

                if (checkBox11.Checked == true)
                {
                    textBox1.Text = boder_parser(textBox2.Text) + textBox1.Text;
                }
                else
                {
                    textBox1.Text += boder_parser(textBox2.Text);
                }
            }
            else
            {
                textBox1.Text = boder_parser(textBox2.Text);
            }
        }


        string boder_parser(string sbase){

            if (sbase.Contains("<!DOCTYPE") == true)
            {

                sbase = Regex.Replace(sbase, "210096be0c8", ">SR <");
                sbase = Regex.Replace(sbase, "21006be53eb", ">HR <");
                sbase = Regex.Replace(sbase, "210083cbf51", ">R  <");
                sbase = Regex.Replace(sbase, "2100343fe11", ">HN <");
                sbase = Regex.Replace(sbase, "2100675985a", ">N  <");
                if (comboBox4.Text.Contains("HTML") == false)
                {
                    sbase = Regex.Replace(sbase, "&nbsp;", "");
                    //sbase = Regex.Replace(sbase, "<.*?>", "");
                    sbase = Regex.Replace(sbase, "<.*?\n?.*?>", "");
                }
            }

            StringBuilder sbtm = new StringBuilder();
            bool sitaraba = false;
            bool rtb = false;

            if (comboBox5.Text != "")
            {
                sbase = Regex.Replace(sbase, comboBox5.Text, "XX:XX");
            }

            if (comboBox4.SelectedIndex == 1)
            {
                sitaraba = true;
            }
            else if (comboBox4.Text.Contains("RTB変換"))
            {
                rtb = true;

            }
            else if (comboBox4.Text.Contains("RTB収集"))
            {

                string regst = "(\\d+位.+\n)+※.+集計時点のポイントです";
                if (checkBox9.Checked == true)
                {
                    regst = "\\d+位の.*?pt\\d+(,\\d{3})*pt\r\n現在の順位.+\r\n※.+集計時点";
                }


                Regex ranks = new Regex(regst);
                Match rankm = ranks.Match(sbase);
                if (rankm.Success) {

                    sbtm.AppendLine();
                    sbtm.AppendLine(rankm.Value);
                }

                DateTime nw = DateTime.Now;
                Regex ls = new Regex("総マスターズpt\r\n(\\d+(,\\d+)*) pt");
                Match lm = ls.Match(sbase);
                Regex fs = new Regex("本日の総フェス勝利数\r\n(\\d+ / \\d+)\r\n\r\n※.+集計時点");
                Match fm = fs.Match(sbase);
                if (lm.Success)
                {

                    sbtm.AppendLine("|");
                    sbtm.AppendLine(lm.Value);
                    sbtm.AppendLine(nw.ToString("※MM/dd HH:mm 集計時点　らうんじ"));
                } if (fm.Success)
                {

                    sbtm.AppendLine("|");
                    sbtm.AppendLine(fm.Value);
                    //sbtm.AppendLine(nw.ToString("※MM/dd HH:mm 集計時点 フェス"));
                }

                return sbtm.ToString();
            }
            else if (comboBox4.Text.Contains("全PT収集"))
            {


                Regex ranks = new Regex("報酬獲得者数[\\t 　]+全\\d+名\r\n獲得者数上限アップまであと[\\t 　]+\\d+(,\\d{3})* pt\r\nイベント参加者総獲得pt[\\t 　]+\\d+(,\\d{3})* pt\r\n※.+集計時点");
                Match rankm = ranks.Match(sbase);
                if (rankm.Success)
                {

                    sbtm.AppendLine();
                    sbtm.AppendLine(rankm.Value);
                }
                return sbtm.ToString();
            }
            else if (comboBox4.Text.Contains("中間 / 最終全収集"))
            {
                int start = sbase.IndexOf("前へ次へ");
                int end = sbase.LastIndexOf("前へ次へ")-start;

                if (end-start>8)
                {
                    sbase = sbase.Substring(start + 4, end - 4);
                    sbase = Regex.Replace(sbase, "^\t", "");

                   sbtm.AppendLine(DateTime.Now.ToString("//yyyyMMddHHmm"));
                   sbtm.AppendLine(sbase);
                }
                return sbtm.ToString();
            }
            else if (comboBox4.Text.Contains("100/500/1200/3000収集"))
            {
                
                Regex ranks = new Regex("(\\d+0|\t1?1)位.+\n.+\n.+pt (\\d+(,\\d{3})*)");
                Match rankm = ranks.Match(sbase);
                int interval = 20;
                int mtime = 10;
                int dlay = interval;


                if (comboBox4.Text.Contains("※")) {
                    string tmp2 =comboBox4.Text.Substring(comboBox4.Text.LastIndexOf("※"),comboBox4.Text.Length-comboBox4.Text.LastIndexOf("※"));
                    Regex dd = new Regex("\\d+");
                    Match ddm = dd.Match(tmp2);
                    if(ddm.Success){
                        interval = Convert.ToInt32(ddm.Value);
                        ddm = ddm.NextMatch();
                        mtime = Convert.ToInt32(ddm.Value);
                        ddm = ddm.NextMatch();
                        dlay = Convert.ToInt32(ddm.Value);

                        if (interval == mtime) {
                             interval = 20;
                             mtime = 10;
                             dlay = 20;
                        }
                    }
                }

                DateTime nw = DateTime.Now.AddMinutes(-interval);

                if (interval > 1)
                {
                    while (true)
                    {
                        if (Convert.ToInt32(nw.ToString("mm")) % interval == mtime)
                        {
                            goto exs;
                        }
                        nw = nw.AddMinutes(1);

                    }
                }
                else {
                    nw = DateTime.Now;
                }



            exs:
                if (rankm.Success)
                {
                    bool ranks_add = checkBox12.Checked;

                    if(ranks_add == false) { 
                        
                    sbtm.Append( DateTime.Now.ToString("yyyyMMddHHmm"));
                    sbtm.Append("(");
                    sbtm.Append(nw.ToString("yyyyMMddHHmm"));
                    if (checkBox9.Checked == true)
                    {
                        sbtm.Append("\t");
                    }
                    else
                    {
                        sbtm.AppendLine(")");
                    }

                    if (checkBox9.Checked == true)
                    {
                        Regex idol = new Regex("(秋月律子|天海春香|伊吹翼|エミリー|大神環|春日未来|音無小鳥|我那覇響|菊地真|如月千早|北上麗花|北沢志保|木下ひなた|高坂海美|佐竹美奈子|四条貴音|篠宮可憐|島原エレナ|ジュリア|周防桃子|高槻やよい|高山紗代子|田中琴葉|天空橋朋花|徳川まつり|所恵美|豊川風花|中谷育|永吉昴|七尾百合子|二階堂千鶴|野々原茜|萩原雪歩|箱崎星梨花|馬場このみ|福田のり子|双海亜美|双海真美|星井美希|舞浜歩|真壁瑞希|松田亜利沙|三浦あずさ|宮尾美也|水瀬伊織|最上静香|望月杏奈|百瀬莉緒|矢吹可奈|横山奈緒|ロコ)");
                        Match idm = idol.Match(sbase);
                        if (idm.Success)
                        {
                            sbtm.Append(idm.Value);
                            sbtm.Append("\t");
                            
                     Regex aranks = new Regex("(\\d+)位.+\n.+\n.+pt \\d+(,\\d{3})*");
                     Match arankm = aranks.Match(sbase);
                         while(arankm.Success){
                        sbtm.Append(Regex.Replace(arankm.Value, ".+\r\n.+\r\n.+pt ", "").Trim());
                        sbtm.Append("\t");
                        arankm= arankm.NextMatch();
                          }
                        }
                    }
                    else {
                    while(rankm.Success){
                        sbtm.AppendLine(Regex.Replace(rankm.Value, ".+\r\n.+\r\n.+pt ", "").Trim());
                        rankm= rankm.NextMatch();
                    }
                    }
                    }
                   else{

                       GroupCollection groups = rankm.Groups;

                       while (rankm.Success)
                       {
                           sbtm.Append(rankm.Groups[1]);
                           sbtm.Append("位\t");
                           sbtm.Append(rankm.Groups[2]);
                           sbtm.AppendLine(" pt");
                           rankm = rankm.NextMatch();
                       }
                       sbtm.AppendLine(DateTime.Now.ToString("※MM/dd HH:mm 集計時点 BYそくどん\r\n"));
                  }
                }
            sbtm.AppendLine();
            return sbtm.ToString();
            }

            string[] str = sbase.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.None);
            string[] strs = str[0].Split(new string[] { "\t" }, StringSplitOptions.None);
            StringBuilder sb = new StringBuilder();
            DateTime dt;
            DateTime utcTime;
            DateTime localTime;
            int base100=0;
            int base500=0;
            int base1200=0;
            List<BorderTBL> arr = new List<BorderTBL>();
            SorterByTimer sorter = new SorterByTimer();
            DateTime t = DateTime.Now;
            int[] jsoku = new int[6];
            string baseyear = "2015/";
            string basenyear = "2016/";

            if (comboBox4.Text.Contains("BASEYEAR"))
            {
                Regex by = new Regex("20\\d\\d");
                Match bym = by.Match(comboBox4.Text);
                if (bym.Success)
                {
                    baseyear = bym.Value+"/";
                    DateTime dd = new DateTime(Convert.ToInt32(bym.Value), 1, 1, 0, 00,0, 0);
                    basenyear = dd.AddYears(1).ToString("yyyy/");
                }
            }
            else if (comboBox4.Text.Contains("IMC"))
            {

                sbtm.AppendLine(DateTime.Now.ToString("//yyyyMMddHHmm"));
                bool add = false;
                foreach (string s in str)
                {
                    if (comboBox4.Text.Contains("ぷれみあ"))
                    {
                        if (s.Contains("ステージ詳細") == true)
                        {
                            add = false;
                        }
                        else if (s.Contains("ステージ状況") == true)
                        {
                            add = true;

                        }
                        else if (add == true)
                        {
                            sbtm.AppendLine(s);
                        }
                    }
                    else
                    {
                        if (s.Contains("イベントTOPへ") == true)
                        {
                            add = false;
                        }
                        else if (s.Contains("自ラウンジ") == true)
                        {
                            add = true;

                        }
                        else if (add == true)
                        {
                            sbtm.AppendLine(s);
                        }
                    
                    }
                }


                return sbtm.ToString();
            }

            if (comboBox4.Text.Contains("HTML"))
            {

                string ll="<a href=\"http://imas.gree-apps.net/app/index.php/lounge/profile/id/(\\d+)\" ?>(.*?)</a>";
                string ff ="<span class=\"txt-sub2\">ファン数/週&nbsp;</span>(\\d+(,\\d+)*)人";
               string uu= "<a href=\"http://imas.gree-apps.net/app/index.php/mypage/user_profile/id/(\\d+)\" ?>(.*?)</a>";
                string ss= "<span class=\"txt-sub2\">.*?pt&nbsp;</span>(\\d+(,\\d+)*)";

                Regex lounge = new Regex(ll);
                Regex fansuu = new Regex(ff);
                Regex user = new Regex(uu);
                Regex score = new Regex(ss);
                string temp = "";
                foreach (string s in str)
                {
                    Match lm = lounge.Match(s);
                    Match fm = fansuu.Match(s);
                    Match um = user.Match(s);
                    Match sm = score.Match(s);
                    if (lm.Success) {
                        sbtm.Append(Regex.Replace(lm.Value, ll, "$2\t$1\t"));
                    }
                    if (fm.Success)
                    {
                        sbtm.AppendLine(Regex.Replace(fm.Value, ff, "$1"));
                    }
                    if (um.Success)
                    {
                        sbtm.Append(Regex.Replace(um.Value, uu, "$2\t$1\t"));
                    }
                    if (sm.Success)
                    {
                        sbtm.AppendLine(Regex.Replace(sm.Value, ss, "$1"));
                    }


                }

                sbtm.AppendLine(DateTime.Now.ToString("※MM/dd HH:mm 集計時点 BYそくどん\r\n"));

                return sbtm.ToString().Replace(",","");
            }

            string tyear = baseyear;
            string tmp = "";
            string pthout = textBox4.Text;
            string datetimes = comboBox1.Text;
            
                    if (checkBox1.Checked == false)
                    {
                        datetimes = "HH:mm";
                    }

            if (comboBox4.SelectedIndex == 6) {
                pthout = ".?";
            }



            if (textBox3.Text == "") {
                textBox3.Text = "0";
            }
            int line = 0;
            int outlim = 0;
            int.TryParse(textBox3.Text, NumberStyles.Integer, NumberFormatInfo.CurrentInfo, out line);
            int.TryParse(textBox5.Text, NumberStyles.Integer, NumberFormatInfo.CurrentInfo, out outlim);
            bool cms= checkBox8.Checked;

            if (radioButton3.Checked == true && str.Length>0)
            {


                if (DateTime.TryParse(str[0], out dt))
                {
                    sb.AppendLine(dt.ToShortDateString());
                }

                StringBuilder sb10 = new StringBuilder();
                StringBuilder sb11 = new StringBuilder();
                StringBuilder sb50 = new StringBuilder();
                StringBuilder sb30 = new StringBuilder();
                StringBuilder sb100 = new StringBuilder();
                StringBuilder sb5000 = new StringBuilder();


                if (radioButton2.Checked == false)
                {
                        string[] legend = comboBox3.Text.Split(new string[] { "/" }, StringSplitOptions.None);
                        if (legend.Length > 4)
                        {
                            sb10.AppendLine(legend[0]);
                            sb50.AppendLine(legend[1]);
                            sb11.AppendLine(legend[2]);
                            sb30.AppendLine(legend[3]);
                            sb100.AppendLine(legend[4]);
                        }
                        else
                        {
                            sb10.AppendLine("100位");
                            sb50.AppendLine("500位");
                            sb30.AppendLine("3000位");
                            sb100.AppendLine("10000位");
                            if (comboBox2.SelectedIndex < comboBox2.Items.Count - 18)
                            {
                                sb11.AppendLine("1200位");
                            }
                            else if (comboBox2.SelectedIndex < comboBox2.Items.Count - 8)
                            {
                                sb11.AppendLine("1100位");
                            }
                            else
                            {
                                sb11.AppendLine("1000位");
                            }
                        }

                        if (sitaraba==true)
                        { 
                            Regex bases = new Regex("\\d+(,\\d{3})*");
                            Match bm= bases.Match(comboBox3.Text);
                        base100=Convert.ToInt32(bm.Value);
                        bm = bm.NextMatch();
                        base500 = Convert.ToInt32(bm.Value);
                        bm = bm.NextMatch();
                        base1200 = Convert.ToInt32(bm.Value);
                        
                        }
                }
                else
                {

                    if (rtb == true)
                    {
                        Regex ra = new Regex("(\\d+位[\\t 　]+\\d+(,\\d{3})*.*?\n)+※");
                        Match ram = ra.Match(sbase);
                        if (ram.Success)
                        {
                            Regex ra2 = new Regex("\\d+位");
                            Match ra2m = ra2.Match(ram.Value);
                            int p=1;
                            Array.Resize(ref strs, 16);
                            while (ra2m.Success)
                            {
                                strs[2 * p - 1] = ra2m.Value;
                                ra2m = ra2m.NextMatch();
                                p++;
                                if (p > 7)
                                {
                                    break;
                                }
                            }

                        }
                    }
                    
                    if (strs.Length > 6){
                        sb10.AppendLine(strs[1]);
                        sb50.AppendLine(strs[3]);
                        sb11.AppendLine(strs[5]);
                        if (strs.Length > 8)
                        {
                            sb30.AppendLine(strs[7]);
                            if (strs.Length > 10)
                            {
                                sb100.AppendLine(strs[9]);
                            }
                            if (strs.Length > 12)
                            {
                                sb5000.AppendLine(strs[11]);
                            }
                        }
                    }


                }

                bool ok = true;
                int f = -6;

                foreach (string s in str)
                {
                    string[] str2 = s.Split(new string[] { "\t" }, StringSplitOptions.None);
                    f++;

                    if (line > 0)
                    {
                        line--;
                    }
                    else
                    {
                        if ((comboBox4.SelectedIndex<1) || (comboBox4.SelectedIndex==6))
                        {

                            if (str2.Length > 6)
                            {
                                if (checkBox1.Checked == true)
                                {
                                    if (DateTime.TryParse(str2[0], out dt))
                                    {
                                        if (checkBox2.Checked == true)
                                        {
                                            //GMT +9:00
                                            utcTime = dt.AddHours(-9);
                                            localTime = System.TimeZoneInfo.ConvertTimeFromUtc(utcTime, System.TimeZoneInfo.Local);
                                            dt = localTime;
                                        }


                                        //2013
                                        if (comboBox2.SelectedIndex >= comboBox2.Items.Count - 42)
                                        {
                                            DateTime juuyonen = new DateTime(2014, 1, 1, 0, 0, 0, 0);
                                            while (dt > juuyonen) {
                                                dt = dt.AddYears(-1);                                           
                                            }
                                        }
                                        //2014
                                        else if (comboBox2.SelectedIndex >= comboBox2.Items.Count - 114)
                                        {
                                            DateTime juuyonen = new DateTime(2015, 1, 1, 0, 0, 0, 0);
                                            while (dt > juuyonen)
                                            {
                                                dt = dt.AddYears(-1);
                                            }
                                        }//2015
                                        //else if (comboBox2.SelectedIndex >= comboBox2.Items.Count - 50)
                                        //{
                                        //}


                                        str2[0] = dt.ToString(datetimes);
                                        ok = true;
                                    }
                                    else
                                    {
                                        ok = false;
                                    }

                                }

                                if (checkBox10.Checked == true)
                                {
                                    if (f > 1)
                                    {
                                        sb10.AppendLine(str2[0]);

                                        for (int i = 1; i < str2.Length; i++)
                                        {
                                            if (Regex.IsMatch(str2[i], pthout))
                                            {
                                                sb10.Append(strs[i]);
                                                sb10.Append(" (+");
                                                sb10.Append(str2[i]);
                                                sb10.AppendLine(")");
                                            }
                                        }
                                        sb10.AppendLine();
                                    }
                                }
                                else if (ok == true)
                                {

                                    if (comboBox4.SelectedIndex < 6)
                                    {

                                        if (Regex.IsMatch(str2[6], pthout))
                                        {
                                            sb11.Append(str2[0]);
                                            sb11.Append("\t");
                                            sb11.Append(addcmst(str2[5],cms));
                                            sb11.Append("(+");
                                            sb11.Append(addcmst(str2[6],cms));
                                            sb11.AppendLine(")");
                                        }

                                        if (Regex.IsMatch(str2[2], pthout))
                                        {
                                            sb10.Append(str2[0]);
                                            sb10.Append("\t");
                                            sb10.Append(addcmst(str2[1],cms));
                                            sb10.Append("(+");
                                            sb10.Append(addcmst(str2[2],cms));
                                            sb10.AppendLine(")");
                                        }

                                        if (Regex.IsMatch(str2[4], pthout))
                                        {
                                            sb50.Append(str2[0]);
                                            sb50.Append("\t");
                                            sb50.Append(addcmst(str2[3],cms));
                                            sb50.Append("(+");
                                            sb50.Append(addcmst(str2[4],cms));
                                            sb50.AppendLine(")");
                                        }

                                        if (str2.Length > 8 && Regex.IsMatch(str2[8], pthout))
                                        {
                                            sb30.Append(str2[0]);
                                            sb30.Append("\t");
                                            sb30.Append(addcmst(str2[7],cms));
                                            sb30.Append("(+");
                                            sb30.Append(addcmst(str2[8],cms));
                                            sb30.AppendLine(")");
                                        }

                                        if (str2.Length > 10 && Regex.IsMatch(str2[10], pthout))
                                        {
                                            sb100.Append(str2[0]);
                                            sb100.Append("\t");
                                            sb100.Append(addcmst(str2[9],cms));
                                            sb100.Append("(+");
                                            sb100.Append(addcmst(str2[10],cms));
                                            sb100.AppendLine(")");
                                        }
                                    }
                                    else {
                                        
                                        if (int.TryParse(str2[1],NumberStyles.Integer | NumberStyles.AllowThousands, NumberFormatInfo.CurrentInfo,out jsoku[0])){
                                            for (int i = 3; i < 9; i += 2)
                                            {
                                            int.TryParse(str2[i],NumberStyles.Integer | NumberStyles.AllowThousands, NumberFormatInfo.CurrentInfo,out jsoku[i>>1]);                                            
                                            }       
                                        arr.Add(new BorderTBL(dt, jsoku[0], jsoku[1], jsoku[2],jsoku[3],jsoku[4],jsoku[5]));
                                        rtb = true;
                                        }

                                    }
                                }


                            }

                        }
                        else
                        {
                            if (sitaraba)//2
                            {
                                Regex jisoku = new Regex("^\\d?\\d:\\d\\d *\\d+\\([\\+-]-?\\d+\\)?");
                                Regex jisokuh = new Regex("^(\\d{4})?(\\d{4})?\\d{3,4}\\((\\d{4})?(\\d{4})?\\d{3,4}\\)");
                                Regex borders = new Regex("(--|\\d*ry|\\d+(,\\d{3})*)");
                                Match jm = jisoku.Match(s);
                                Match hm = jisokuh.Match(s);
                                Match ptm = borders.Match(s);


                                if (jm.Success)
                                {
                                    sb10.Append(jm.Value.Substring(0, jm.Value.IndexOf(" ")));
                                    string ttmp = jm.Value.Substring(jm.Value.IndexOf(" "), jm.Value.Length - jm.Value.IndexOf(" ")).Trim();
                                    Match ptz = borders.Match(ttmp);
                                    sb10.Append("\t");

                                    if (ptz.Success)
                                    {
                                        int score = Convert.ToInt32(ptz.Value);
                                        if (score < base500)
                                        {
                                            sb10.Append("\t\t\t\t");
                                        }
                                        else if (score < base100)
                                        {
                                            sb10.Append("\t\t");
                                        }

                                        sb10.Append(ptz.Value);
                                        sb10.Append("\t");
                                        tmp = tmp.Replace(ptz.Value, "").Replace("(", "").Replace("+", "");
                                        Match pty = borders.Match(tmp);
                                        sb10.AppendLine(pty.Value);
                                    }

                                }
                                else if (hm.Success)
                                {
                                    f = -6;
                                    tmp=hm.Value.Substring(hm.Value.IndexOf("(") + 1, hm.Value.Length - hm.Value.IndexOf("(") - 1).Replace(")", "");
                                    
                                    if (checkBox1.Checked)
                                    {
                                        if (DateTime.TryParseExact(tmp, "yyyyMMddHHmm", CultureInfo.InvariantCulture,DateTimeStyles.None, out dt) ){

                                            if (checkBox2.Checked == true)
                                            {
                                                //GMT +9:00
                                                utcTime = dt.AddHours(-9);
                                                localTime = System.TimeZoneInfo.ConvertTimeFromUtc(utcTime, System.TimeZoneInfo.Local);
                                                dt = localTime;
                                            }
                                            tmp = dt.ToString(datetimes);
                                        }
                                        else if (DateTime.TryParseExact(tmp, "HHmm", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                                        {

                                            if (checkBox2.Checked == true)
                                            {
                                                //GMT +9:00
                                                utcTime = dt.AddHours(-9);
                                                localTime = System.TimeZoneInfo.ConvertTimeFromUtc(utcTime, System.TimeZoneInfo.Local);
                                                dt = localTime;
                                            }
                                            tmp = dt.ToString(datetimes);
                                        }
                                    }
                                    else {
                                    tmp=  tmp.Insert( 2,":");
                                    }
                                    sb10.Append(tmp);
                                    sb10.Append("\t");
                                    sb11.Clear();
                                    sb30.Clear();
                                    sb50.Clear();
                                    sb100.Clear();
                                    sb11.Append(hm.Value.Substring(0, hm.Value.IndexOf("(")));

                                }
                                else if (f <= 0)
                                {
                                    if (ptm.Success==true)
                                    {
                                        if (f == -5)
                                        {
                                            sb30.Append(ptm.Value);
                                            sb30.Append("\t\t");
                                        }
                                        else if (f == -4)
                                        {
                                            sb50.Append(ptm.Value);
                                            sb50.Append("\t\t");
                                        }
                                        else if (f == -3)
                                        {
                                            sb10.Append(ptm.Value);
                                            sb10.Append("\t\t");
                                        }
                                        else if (f == -2)
                                        {
                                            sb100.Append(ptm.Value);
                                            sb100.Append("\t\t");
                                        }
                                        else if (f == -1)
                                        {
                                            sb100.Append(ptm.Value);
                                            sb100.Append("\t\t");
                                        }
                                        else if(f==0){
                                            sb10.Append(sb50.ToString());
                                            sb10.Append(sb30.ToString());
                                            sb10.Append(sb100.ToString());
                                            sb10.AppendLine(sb11.ToString());
                                        }
                                    }
                                    else
                                    {
                                        sb10.Append(sb50.ToString());
                                        sb10.Append(sb30.ToString());
                                        sb10.Append(sb100.ToString());
                                        sb10.AppendLine(sb11.ToString());
                                        f = 0;
                                    }

                                }
                            }
                            else {//3

                                string nums = "(1|\\d+0||11)位[\\t 　]+(\\d+(,\\d{3})*)";
                                string reg = "※?(\\d\\d\\d\\d/)?(\\d\\d/\\d\\d).*?(\\d\\d:\\d\\d) 集計時点";
                                string regtpt = "イベント参加者総獲得pt[\\t 　]+(\\d+(,\\d{3})*)";

                                Regex ranks = new Regex(nums);
                                Match rankm = ranks.Match(s);
                                Regex adds = new Regex(reg);
                                Match addm = adds.Match(s);
                                Regex addspt = new Regex(regtpt);
                                Match addmpt = addspt.Match(s);

                                if (rankm.Success) {
                                    tmp = rankm.Value;
                                    //if(tmp.Contains("1位")){
                                    //   // f = -5;
                                    //    if (tmp.Contains("11位"))//配列番目
                                    //    {
                                    //        tmp = Regex.Replace(tmp, nums, "$2");
                                    //        int.TryParse(tmp, NumberStyles.Integer | NumberStyles.AllowThousands, NumberFormatInfo.CurrentInfo, out jsoku[f+4]);

                                    //    }
                                    //    else
                                    //    {
                                    //        tmp = Regex.Replace(tmp, nums, "$2");
                                    //        int.TryParse(tmp, NumberStyles.Integer | NumberStyles.AllowThousands, NumberFormatInfo.CurrentInfo, out jsoku[0]);
                                    //    }
                                    //}
                                    //else if (tmp.Contains("0位"))
                                    //{

                                        tmp = Regex.Replace(tmp, nums, "$2");
                                        
                                        if ((f >= -6) && (f <0))
                                        {
                                            int.TryParse(tmp,NumberStyles.Integer | NumberStyles.AllowThousands, NumberFormatInfo.CurrentInfo,out jsoku[f+6]);   
                                         }
                                    //}
                                }
                                else if (addmpt.Success)
                                {
                                    tmp = addmpt.Value;
                                        f = -7;
                                            tmp = Regex.Replace(tmp,  regtpt, "$1");
                                            int.TryParse(tmp, NumberStyles.Integer | NumberStyles.AllowThousands, NumberFormatInfo.CurrentInfo, out jsoku[4]);

                                }
                                else if (addm.Success)
                                {

                                    string[] expectedFormats = { "yyyy/MM/dd HH:mm" };
                                    tmp = Regex.Replace(addm.Value, reg, "$1$2 $3").Replace(" 集計時点", "");
                                    if (tmp.Length < 12)
                                    {
                                        if (tmp.Contains("01/01"))
                                        {
                                            tyear = basenyear;
                                        }
                                        else if (tmp.Contains("12/31"))
                                        {
                                            tyear = baseyear;
                                        }
                                        tmp = tyear + tmp;
                                    }
                                    t = DateTime.ParseExact(tmp, expectedFormats, System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None);

                                    if (checkBox2.Checked == true)
                                    {
                                        //GMT +9:00
                                        utcTime = t.AddHours(-9);
                                        localTime = System.TimeZoneInfo.ConvertTimeFromUtc(utcTime, System.TimeZoneInfo.Local);
                                        t = localTime;
                                    }
                                    arr.Add(new BorderTBL(t, jsoku[0], jsoku[1], jsoku[2], jsoku[3], jsoku[4],jsoku[5]));
                                    f = -7;
                                }
                                else {
                                    f = -7;
                                }

                            }
                        }
                    }
                }





                if (rtb)
                {
                    arr.Sort(sorter);

                        int j=outlim;
                        int l= arr.Count-1;
                        if (j < 0) {
                            j = l + j;
                        }
                        if (outlim > l  || j<0) {
                         return "配列数オーバー ST:" + j.ToString("d") + " MAX:" + l.ToString("d");
                        }


                    int tt = 0;
                    if (checkBox1.Checked == false)
                    {
                        datetimes = "HH:mm";
                    }


                    sb10.Append(arr[j].timer.ToString(datetimes));
                    sb50.Append(arr[j].timer.ToString(datetimes));
                    sb11.Append(arr[j].timer.ToString(datetimes));
                    sb30.Append(arr[j].timer.ToString(datetimes));
                    sb100.Append(arr[j].timer.ToString(datetimes));
                    sb5000.Append(arr[j].timer.ToString(datetimes));
                    sb10.Append("\t");
                    sb50.Append("\t");
                    sb11.Append("\t");
                    sb30.Append("\t");
                    sb100.Append("\t");
                    sb5000.Append("\t");
                    sb10.AppendLine(addcm(arr[j].rank1,cms));
                    sb50.AppendLine(addcm(arr[j].rank2, cms));
                    sb11.AppendLine(addcm(arr[j].rank3, cms));
                    sb30.AppendLine(addcm(arr[j].rank4, cms));
                    sb100.AppendLine(addcm(arr[j].rank5, cms));
                    sb5000.AppendLine(addcm(arr[j].rank6, cms));

                    for (int z = j; z < l; z++)
                    {
                        sb10.Append(arr[z+1].timer.ToString(datetimes));
                        sb10.Append("\t");
                        tt = arr[z + 1].rank1 - arr[z].rank1;
                        sb10.Append(addcm(arr[z+1].rank1,cms));
                        sb10.Append("(+");
                        sb10.Append(addcm(tt,cms));
                        sb10.AppendLine(")");

                    }
                    for (int z = j; z <l; z++)
                    {
                        sb50.Append(arr[z + 1].timer.ToString(datetimes));
                        sb50.Append("\t");
                        tt = arr[z + 1].rank2 - arr[z].rank2;
                        sb50.Append(addcm(arr[z + 1].rank2,cms));
                        sb50.Append("(+");
                        sb50.Append(addcm(tt,cms));
                        sb50.AppendLine(")");

                    }
                    for (int z = j; z < l; z++)
                    {
                        sb11.Append(arr[z + 1].timer.ToString(datetimes));
                        sb11.Append("\t");
                        tt = arr[z + 1].rank3 - arr[z].rank3;
                        sb11.Append(addcm(arr[z+1].rank3,cms));
                        sb11.Append("(+");
                        sb11.Append(addcm(tt,cms));
                        sb11.AppendLine(")");

                    }
                    for (int z = j; z <l; z++)
                    {
                        sb30.Append(arr[z + 1].timer.ToString(datetimes));
                        sb30.Append("\t");
                        tt = arr[z + 1].rank4 - arr[z].rank4;
                        sb30.Append(addcm(arr[z + 1].rank4,cms));
                        sb30.Append("(+");
                        sb30.Append(addcm(tt,cms));
                        sb30.AppendLine(")");

                    }
                    for (int z = j; z <l; z++)
                    {
                        sb100.Append(arr[z + 1].timer.ToString(datetimes));
                        sb100.Append("\t");
                        tt = arr[z + 1].rank5 - arr[z].rank5;
                        sb100.Append(addcm(arr[z + 1].rank5,cms));
                        sb100.Append("(+");
                        sb100.Append(addcm(tt,cms));
                        sb100.AppendLine(")");

                    }
                    for (int z = j; z < l; z++)
                    {
                        sb5000.Append(arr[z + 1].timer.ToString(datetimes));
                        sb5000.Append("\t");
                        tt = arr[z + 1].rank6 - arr[z].rank6;
                        sb5000.Append(addcm(arr[z + 1].rank6, cms));
                        sb5000.Append("(+");
                        sb5000.Append(addcm(tt, cms));
                        sb5000.AppendLine(")");

                    }
                    
                }


                    if (checkBox7.Checked)
                    {
                        sb.AppendLine(sb5000.ToString());
                        sb.AppendLine(sb100.ToString());
                    }

                    if (checkBox6.Checked)
                    {
                        sb.AppendLine(sb30.ToString());
                    }
                    if (checkBox5.Checked)
                    {
                        sb.AppendLine(sb11.ToString());
                    }
                    if (checkBox4.Checked)
                    {
                        sb.AppendLine(sb50.ToString());
                    }
                    if (checkBox3.Checked)
                    {
                        sb.AppendLine(sb10.ToString());
                    }
            }
            else
            {
                /*if (radioButton2.Checked == true)
                {
                    Regex gasha = new Regex("^ \\[(N|HN|R|HR|SR)\\]");
                    Regex prebox = new Regex("^   (N|HN|R|HR|SR)");

                    foreach (string s in str)
                    {


                        if (gasha.IsMatch(s) == true)
                        {
                            sb.AppendLine(s.Trim());
                        }
                        else if (prebox.IsMatch(s) == true)
                        {
                            sb.AppendLine(s.Trim());
                        }
                    }
                }
                else {
                */
                Regex idol = new Regex("(秋月律子|天海春香|伊吹翼|ｴﾐﾘ-|大神環|春日未来|音無小鳥|我那覇響|菊地真|如月千早|北上麗花|北沢志保|木下ひなた|高坂海美|佐竹美奈子|四条貴音|篠宮可憐|島原ｴﾚﾅ|ｼﾞｭﾘｱ|周防桃子|高槻やよい|高山紗代子|田中琴葉|天空橋朋花|徳川まつり|所恵美|豊川風花|中谷育|永吉昴|七尾百合子|二階堂千鶴|野々原茜|萩原雪歩|箱崎星梨花|馬場このみ|福田のり子|双海亜美|双海真美|星井美希|舞浜歩|真壁瑞希|松田亜利沙|三浦あずさ|宮尾美也|水瀬伊織|最上静香|望月杏奈|百瀬莉緒|矢吹可奈|横山奈緒|ﾛｺ)");
                    Regex newidol = new Regex("「.*?[」｣]");
                    Regex skl = new Regex("Lv\\.[12]?[0-9]/20");
                    Regex rare = new Regex("(SR|HR|R |HN|N ) ");
                    Regex stop = new Regex("(全てのｱｲﾄﾞﾙ|ロコ)");
                    bool parseok = true;

                    foreach (string s in str)
                    {
                        if (stop.IsMatch(s) == true){
                            parseok = false;
                            if(s.Contains("ロコ")==true) {
                            parseok = true;
                              }
                        }
                        if (parseok == true)
                        {
                            if (skl.IsMatch(s) == true )
                            {
                                if( checkBox10.Checked == true){
                                sb.Remove(sb.Length - 2, 2);
                                sb.Append(" S");
                                sb.AppendLine(s.Substring(s.IndexOf("Lv"), s.Length - s.IndexOf("Lv")).Trim());
                                }
                            }
                            else if (idol.IsMatch(s) == true)
                            {
                                if (newidol.IsMatch(s) == false)
                                {
                                    sb.AppendLine(s.Trim());
                                }
                            }
                            else if (rare.IsMatch(s.TrimStart()) == true)
                            {
                                sb.Append(s.TrimStart().Substring(0, 3));
                            }
                        }
                    }
                
                
                //}
            }


            return sb.ToString().Replace("+-","-");


        }


        private string addcm(int c, bool d)
        {
            return (d) ? c.ToString("#,0") : c.ToString();      
        
        }


        private string addcmst(string c, bool d)
        {
            return (d) ? Regex.Replace(c,"(\\d)(?=(\\d\\d\\d)+(?!\\d))","$1,"): c;

        }


        private class SorterByTimer : IComparer<BorderTBL>
        {
            public int Compare(BorderTBL x, BorderTBL y)
            {
                return x.timer.CompareTo(y.timer);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string url = comboBox2.Text;
            if (url.Contains("日速") == true)
            {
                radioButton4.Checked = false;
                radioButton2.Checked = true;


                comboBox4.SelectedIndex = 0;
            }
            Regex ur = new Regex("(h?ttps?://)?((bit.ly/[A-Za-z0-9]*|goo\\.gl/[A-Za-z0-9]*)|(docs\\.google\\.com/spreadsheet/(ccc|pub|lv)\\?key=[0-9a-zA-Z\\-]*))");
            //Regex ur = new Regex("h?ttps?://[a-z0-9/A-Z=?.-]*");
            Match urm = ur.Match(url);
            Regex gid = new Regex("&gid=\\d+");
            Match gidm = gid.Match(url);


            if (urm.Success == true)
            {
                url = url_regex(urm.Value,gidm.Value);
                Uri u = new Uri(url);
                if (u.IsWellFormedOriginalString() == true)
                {
                System.Net.HttpWebRequest webreq =
                    (System.Net.HttpWebRequest)
                        System.Net.WebRequest.Create(url);

                System.Net.HttpWebResponse webres =
                    (System.Net.HttpWebResponse)webreq.GetResponse();

                System.Text.Encoding enc = Encoding.GetEncoding("UTF-8");
                System.IO.Stream st = webres.GetResponseStream();
                System.IO.StreamReader sr = new System.IO.StreamReader(st, enc);
                string html = sr.ReadToEnd();
                sr.Close();


                //textBox1.Text = html;

                //table抽出
                Regex r = new Regex("<table.*?</table>");
                Match m = r.Match(html);
                if (m.Success)
                {
                    html = m.Value;
                    html = html.Replace("</tr>", "\r\n");
                    html = Regex.Replace(html, "<td.*?>", "\t");
                    html = Regex.Replace(html, "<.*?>", "");
                    html = Regex.Replace(html, "\\.\t", "");

                string[] str = html.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                StringBuilder sb = new StringBuilder();
                foreach (string s in str)
                {
                    sb.AppendLine(s.Trim());

                }
                textBox1.Text = sb.ToString();
                }
                else
                {
                    html = html.Replace("\n", "\r\n");
                    textBox1.Text =html;
                }

                }
                else
                {
                    MessageBox.Show(this, "正しくないURLです", "エラー");

                }
            }
            else
            {
                MessageBox.Show(this, "スプレッドシート参照用URLが登録されてません\n参照先URLを選択するかGooglelURLshortnerで登録したものかスプレッドシート正規のURLを貼り付けて下さい\nhttp://goo.gl/*****\nhttps://docs.google.com/spreadsheet/pub\\?key=***...***", "エラー");

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            panel1.Visible = !panel1.Visible;
            reset_ui();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            panel1.Visible = false;
            button3.Enabled = false;
            reset_ui();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            panel1.Visible = false;
            button3.Enabled = false;
            reset_ui();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
                        button3.Enabled = true;
                        checkBox10.Checked = false;
                        reset_ui();
        }

        void reset_ui(){
            if (panel1.Visible == false)
            {
                textBox1.SetBounds(10, 63, 200, 100, BoundsSpecified.Y);
                textBox1.SetBounds(10, 50, 200, 284, BoundsSpecified.Height);
            }
            else
            {
                textBox1.SetBounds(10, 249, 200, 100, BoundsSpecified.Y);
                textBox1.SetBounds(10, 50, 200, 98, BoundsSpecified.Height);
            }
        
        
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string url = "http://m.ip.bn765.com/1100";
            int　i =(comboBox2.SelectedIndex &　254)+1;
            string bk = comboBox2.Text;
            Regex gazou = new Regex(",[0-9A-Za-z]+$");


            if (comboBox2.SelectedIndex > -1)
            {
                string selecter = comboBox2.Items[i].ToString();
                Match gidm = gazou.Match(selecter);
                if (gidm.Success)
                {
                    selecter = gidm.Value.Replace(",", "");
                    pictureBox1.ImageLocation = url + selecter;
                }
                else if (pictureBox1.Image != null)
                {
                    pictureBox1.Image = null;
                }
            }
            else if (pictureBox1.Image != null)
            {
                pictureBox1.Image = null;
            }

            comboBox2.Text = bk;
            


        }

        private void バージョンToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Form2 f = new Form2();
            f.ShowDialog();
            f.Dispose();
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b' &&  e.KeyChar != '-')
    {
        e.Handled = true;
    }
        }

        private void がしゃふぃるたー_Load(object sender, EventArgs e)
        {

            reset_ui();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            string url = comboBox2.Text;
            //string eventid = "http://imas.gree-apps.net/app/index.php/event/ranking/segment/1/page/1/event_id/";
            string eventid = " http://imas.gree-apps.net/app/index.php/event/ZZ/ranking/general?page=1";
            string chrome = GetDefaultBrowserExePath();
            Regex ur = new Regex("(h?ttps?://)?((goo\\.gl/[A-Za-z0-9]*)|(docs\\.google\\.com/spreadsheet/(pub|lv)\\?key=[0-9a-zA-Z\\-]*))");
            Match urm = ur.Match(url);
            Regex gid = new Regex("&gid=\\d+");
            Match gidm = gid.Match(url);

            if (urm.Success == true)
            {
                url = url_regex(urm.Value,gidm.Value);
            Uri u = new Uri(url);
              if (u.IsWellFormedOriginalString() == true)
              {
                    Process.Start(chrome, url);
                }
                else
                {
                    MessageBox.Show(this, "正しくないURLです", "エラー");                

                }
            Regex eid = new Regex("[0-9]{1,2}$");
            Match eidm = eid.Match(url);
               if (chrome.Contains("chrome"))
               {
                   if (eidm.Success==true) {
                       Process.Start(chrome,eventid.Replace("ZZ","eidm.Value"));
               }
             }           
            }
            else {
                MessageBox.Show(this, "スプレッドシート参照用URLが登録されてません\n参照先URLを選択するかGooglelURLshortnerで登録したものかスプレッドシート(旧)正規のURLを貼り付けて下さい\nhttp://goo.gl/*****\nhttps://docs.google.com/spreadsheet/pub\\?key=***...***", "エラー");
            }
        }

        private string url_regex(string url,string gid) {
            Regex ur2 = new Regex("(goo.gl/[A-Za-z0-9]*)|(docs.google.com/spreadsheet/(pub|lv)\\?key=[0-9a-zA-Z\\-]*)");
            Match urm2 = ur2.Match(url);
            if (urm2.Success == true)
            {
                url = "http://" + urm2.Value;
                if (gid != "" && url.Contains("docs")==true) {
                    url = url + gid;
                }
            }

                return url;      
        }


         private static string GetDefaultBrowserExePath()
 {
    return _GetDefaultExePath(@"http\shell\open\command");
  }


  private static string _GetDefaultExePath(string keyPath)
  {
    string path = "";

    // レジストリ・キーを開く
    // 「HKEY_CLASSES_ROOT\xxxxx\shell\open\command」
    RegistryKey rKey = Registry.ClassesRoot.OpenSubKey(keyPath);
    if (rKey != null)
    {
      // レジストリの値を取得する
      string command = (string)rKey.GetValue(String.Empty);
      if (command == null)
      {
        return path;
      }

      // 前後の余白を削る
      command = command.Trim();
      if (command.Length == 0)
      {
        return path;
      }

      // 「"」で始まる長いパス形式かどうかで処理を分ける
      if (command[0] == '"')
      {
        // 「"～"」間の文字列を抽出
        int endIndex = command.IndexOf('"', 1);
        if (endIndex != -1)
        {
          // 抽出開始を「1」ずらす分、長さも「1」引く
          path = command.Substring(1, endIndex - 1);
        }
      }
      else
      {
        // 「（先頭）～（スペース）」間の文字列を抽出
        int endIndex = command.IndexOf(' ');
        if (endIndex != -1)
        {
          path = command.Substring(0, endIndex);
        }
        else
        {
          path = command;
        }
      }
    }

    return path;
  }

  private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
  {

  }

  private void textBox5_TextChanged(object sender, EventArgs e)
  {

  }

  private void textBox1_KeyDown(object sender, KeyEventArgs e)
  {
      //正しい
      if (e.KeyData == (Keys.Control | Keys.A))
      {
              textBox1.SelectionStart = 0;
              textBox1.SelectionLength = textBox1.TextLength;
      }

  }

  private void textBox2_KeyDown(object sender, KeyEventArgs e)
  {

      if (e.KeyData == (Keys.Control | Keys.A))
      {
          textBox2.SelectionStart = 0;
          textBox2.SelectionLength = textBox2.TextLength;
      }
  }

  private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
  {

  }

  private void checkBox9_CheckedChanged(object sender, EventArgs e)
  {

  }

  private void button6_Click(object sender, EventArgs e)
  {
      System.IO.DirectoryInfo di =System.IO.Directory.CreateDirectory("rollback");

      string datestxt = (DateTime.Now).ToString("yyyy-MM-dd HH-mm-ss");
      string datestxt2 = "rollback\\" + datestxt + "-sub.txt";
      datestxt = "rollback\\" + datestxt + ".txt";
      System.IO.StreamWriter sw = new System.IO.StreamWriter(datestxt,
    false,
    System.Text.Encoding.GetEncoding("utf-8"));
      //TextBox1.Textの内容を書き込む
      sw.Write(textBox2.Text);
      sw.Close();
      if (textBox6.Text != "") {
      System.IO.StreamWriter sww = new System.IO.StreamWriter(datestxt2,
    false,
    System.Text.Encoding.GetEncoding("utf-8"));
      //TextBox1.Textの内容を書き込む
      sww.Write(textBox6.Text);
      sww.Close(); 
          System.IO.StreamWriter sw2 = new System.IO.StreamWriter("f.txt",
 false,
 System.Text.Encoding.GetEncoding("utf-8"));
      //TextBox1.Textの内容を書き込む
      sw2.Write(textBox6.Text);
      sw2.Close();
  }
  }




}
}