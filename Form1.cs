using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using System.Diagnostics;
using System.Threading;
using System.Drawing;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using Spire.Pdf.Lists;
//using Spire.Pdf.Exporting.XPS.Schema;
using iTextSharp.text;
using iTextSharp.text.pdf;
using SautinSoft;






//using System.Threading;

namespace hardware
{
    public partial class Form1 : Form
    {
        System.Diagnostics.Process process = new System.Diagnostics.Process();
        string bat = "key.bat";//@"C:\Users\sueda\Desktop\key.bat";

        #region panelin harekti için gerekli kodlar
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr one, int two, int three, int four);
        #endregion

        private PerformanceCounter pfc = new PerformanceCounter("Processor", "% Processor Time", "_Total");
        private PerformanceCounter ram = new PerformanceCounter("Memory", "% Committed Bytes In Use");
        private PerformanceCounter availableRam = new PerformanceCounter("Memory", "Available MBytes");
        private PerformanceCounter disk = new PerformanceCounter("PhysicalDisk", "% Disk Time", "_Total"); //???????? 


        //private PerformanceCounter temp_ = new PerformanceCounter("Thermal Zone Information", "Temperature",  "\\_TZ.THRM");
       

        private void temp()
        {
            PerformanceCounterCategory pcc = new PerformanceCounterCategory("Thermal Zone Information");
            string[] instance = pcc.GetInstanceNames();
            foreach (string instances in instance)
            {
                using (PerformanceCounter cnt = new PerformanceCounter("Thermal Zone Information", "Temperature", instance[0],true))
                {
                    int temmp_ = (int)cnt.NextValue();
                    circularProgressBar4.Value = (int)temmp_- 273;
                    circularProgressBar4.Text = (temmp_ - 273).ToString() + "°C";
                 
                }

            }
            
        }
      

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {   // batarya yüzdesi
            PowerStatus p = SystemInformation.PowerStatus;
            int a = (int)(p.BatteryLifePercent * 100);
            label3.Text ="Pil Yüzdesi"+" "+"%"+ a.ToString();

            DateTime localDate = DateTime.Now;
            label9.Text = localDate.ToString();
            panel2.BackColor = Color.FromArgb(0, 127, 127, 127);
            cpuSpeed();
            ramSpeed();

            //kullanılabilir ram
            int  aRam = (int)availableRam.NextValue();
            double aRam_ = aRam;
            aRam_ = Math.Round((aRam_ / (1024)), 2);  
            label12.Text =aRam_.ToString()+"  GB";

            //int w = Screen.PrimaryScreen.Bounds.Width;
            //int h = Screen.PrimaryScreen.Bounds.Height;

            //this.Size = new Size(w,h);

            int ourScreenWidth = Screen.FromControl(this).WorkingArea.Width;
            int ourScreenHeight = Screen.FromControl(this).WorkingArea.Height;
            float scaleFactorWidth = (float)ourScreenWidth / 1600f;
            float scaleFactorHeigth = (float)ourScreenHeight / 900f;
            SizeF scaleFactor = new SizeF(scaleFactorWidth, scaleFactorHeigth);
            Scale(scaleFactor);
            CenterToScreen();

            radioButton1.Visible = false;
            radioButton2.Visible = false;
            radioButton3.Visible = false;
            radioButton4.Visible = false;
            radioButton5.Visible = false;
            radioButton6.Visible = false;


        }

        private void cpuSpeed()
        {
            execute();
            timer1.Start();
            string cpuSpeed = "wmic cpu get CurrentClockSpeed";
            string text = "";
            File.WriteAllText(bat, cpuSpeed);
            text.Replace(bat, cpuSpeed);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();

            string[] sep = { "CurrentClockSpeed", " " };
            Int32 count = 5;
            string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in str)
            {
                label4.Text = str[4];
                Match match = Regex.Match(str[4], @"(\d+)");
                double parse = double.Parse(match.Groups[0].Value);
                label4.Text =  parse /1000+ " GHz";
            }
        }
        private void ramSpeed()
        {
            execute();
            string ramSpeed = "wmic memorychip get speed";
            string text = "";
            File.WriteAllText(bat, ramSpeed);
            text.Replace(bat, ramSpeed);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            string[] sep = { "speed", " " };
            Int32 count = 5;
            string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in str)
            {
                label5.Text = str[4];
                //MessageBox.Show(str[4]);

                Match match = Regex.Match(str[4], @"(\d+)");
                double parse = double.Parse(match.Groups[0].Value);
                double ss = double.Parse(match.Groups[1].Value);
                label5.Text = ss  +" MHz"+"\n" +parse + " MHz";
            }
        }
        

        private void execute()
        {
            process.StartInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + bat;// bat
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = true;
            process.Start();
        }
       
        private void timer1_Tick(object sender, EventArgs e)
        {
            //progressbar 100 ü  geçince hata text ile uyuşmuyor.


            //float t = temp_.NextValue();
            //circularProgressBar4.Value = (int)t - 273;
            //circularProgressBar4.Text = (t - 273).ToString() + "°C";
;
            temp();

            float dsk = disk.NextValue();

            if (dsk > 100)
            {
                dsk = 100;
                circularProgressBar2.Value = 100;
                circularProgressBar2.Text = "100%";

            }
            else
            {
                circularProgressBar2.Value = (int)dsk;
                circularProgressBar2.Text = string.Format("{0:0}%", dsk);
            }
            
            
           

            float cpu = pfc.NextValue();
                circularProgressBar1.Value = (int)cpu;
                circularProgressBar1.Text = string.Format("{0:0}%", cpu);
            //if (cpu>100/*circularProgressBar1.Value > 100*/)
            //{
            //    circularProgressBar1.Value = 100;
            //}


            
            float pRam = ram.NextValue();
            circularProgressBar3.Value = (int)pRam;
            circularProgressBar3.Text= string.Format("{0:0}%", pRam);
            //progressBar2.Value = (int)pRam;
            //label7.Text = string.Format("{0:0}%", pRam);



            //label3.Text = (t - 273.15).ToString();

        }
  
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //execute();
            
            string readText = File.ReadAllText(bat);
            if (String.IsNullOrEmpty(readText))
            {
                string txt = "1";
               // string text = "";
                File.WriteAllText(bat, txt);
                //text.Replace(bat, txt);

                //StreamWriter asd = new StreamWriter("1");//????
            }

        }
     
        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.ScrollBars = ScrollBars.None;
            button1.BackColor = Color.FromArgb(74, 92, 117);
            button2.BackColor = Color.FromArgb(49,61,77);
            button3.BackColor = Color.FromArgb(49, 61, 77);
            button4.BackColor = Color.FromArgb(49, 61, 77);
            button5.BackColor = Color.FromArgb(49, 61, 77);
            button6.BackColor = Color.FromArgb(49, 61, 77);
            button7.BackColor = Color.FromArgb(49, 61, 77);
            button8.BackColor = Color.FromArgb(49, 61, 77);
            button9.BackColor = Color.FromArgb(49, 61, 77);
            button10.BackColor = Color.FromArgb(49, 61, 77);
            button11.BackColor = Color.FromArgb(49, 61, 77);
            execute();
            string winkey = "wmic path softwarelicensingservice get OA3xOriginalProductKey";
            string text = "";
            File.WriteAllText(bat, winkey);
            text.Replace(bat, winkey);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            string[] sep = { "OA3xOriginalProductKey", " " };
            Int32 count = 6;
            string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in str)
            { textBox1.Text = "ÜRÜN ANAHATARI" + s; }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.BackColor = Color.FromArgb(74,92,117); 
            button1.BackColor = Color.FromArgb(49, 61, 77);
            button3.BackColor = Color.FromArgb(49, 61, 77);
            button4.BackColor = Color.FromArgb(49, 61, 77);
            button5.BackColor = Color.FromArgb(49, 61, 77);
            button6.BackColor = Color.FromArgb(49, 61, 77);
            button7.BackColor = Color.FromArgb(49, 61, 77);
            button8.BackColor = Color.FromArgb(49, 61, 77);
            button9.BackColor = Color.FromArgb(49, 61, 77);
            button10.BackColor = Color.FromArgb(49, 61, 77);
            button11.BackColor = Color.FromArgb(49, 61, 77);
            textBox1.ScrollBars = ScrollBars.None;
            execute();
            string mbSerial = "wmic baseboard get product,Manufacturer & wmic csproduct get identifyingnumber ";//"wmic csproduct get identifyingnumber";
            string text = "";
            File.WriteAllText(bat, mbSerial);
            text.Replace(bat, mbSerial);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            string[] sep = { "identifyingnumber", " " };
            Int32 count = 11;
            string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in str)
            { textBox1.Text = "ANAKART MODELİ" + s; }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(49, 61, 77);
            button2.BackColor = Color.FromArgb(49, 61, 77);
            button3.BackColor = Color.FromArgb(74, 92, 117);
            button4.BackColor = Color.FromArgb(49, 61, 77);
            button5.BackColor = Color.FromArgb(49, 61, 77);
            button6.BackColor = Color.FromArgb(49, 61, 77);
            button7.BackColor = Color.FromArgb(49, 61, 77);
            button8.BackColor = Color.FromArgb(49, 61, 77);
            button9.BackColor = Color.FromArgb(49, 61, 77);
            button10.BackColor = Color.FromArgb(49, 61, 77);
            button11.BackColor = Color.FromArgb(49, 61, 77);
            textBox1.ScrollBars = ScrollBars.None;
            execute();
            string mbModel = "wmic csproduct get name";

            string text = "";
            File.WriteAllText(bat, mbModel);
            text.Replace(bat, mbModel);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            string[] sep = { "name", " " };
            Int32 count = 5;
            string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in str)
            { textBox1.Text = "SİSTEM MODELİ" + s; }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(49, 61, 77);
            button2.BackColor = Color.FromArgb(49, 61, 77);
            button4.BackColor = Color.FromArgb(74, 92, 117);
            button3.BackColor = Color.FromArgb(49, 61, 77);
            button5.BackColor = Color.FromArgb(49, 61, 77);
            button6.BackColor = Color.FromArgb(49, 61, 77);
            button7.BackColor = Color.FromArgb(49, 61, 77);
            button8.BackColor = Color.FromArgb(49, 61, 77);
            button9.BackColor = Color.FromArgb(49, 61, 77);
            button10.BackColor = Color.FromArgb(49, 61, 77);
            button11.BackColor = Color.FromArgb(49, 61, 77);
            textBox1.ScrollBars = ScrollBars.None;
            execute();
            string cpu = " wmic cpu get name";
            string text = "";
            File.WriteAllText(bat, cpu);
            text.Replace(bat, cpu);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            string[] sep = { "name", " " };
            Int32 count = 5;
            string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in str)
            { textBox1.Text = "İŞLEMCİ ADI" + s; }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(49, 61, 77);
            button2.BackColor = Color.FromArgb(49, 61, 77);
            button5.BackColor = Color.FromArgb(74, 92, 117);
            button4.BackColor = Color.FromArgb(49, 61, 77);
            button3.BackColor = Color.FromArgb(49, 61, 77);
            button6.BackColor = Color.FromArgb(49, 61, 77);
            button7.BackColor = Color.FromArgb(49, 61, 77);
            button8.BackColor = Color.FromArgb(49, 61, 77);
            button9.BackColor = Color.FromArgb(49, 61, 77);
            button10.BackColor = Color.FromArgb(49, 61, 77);
            button11.BackColor = Color.FromArgb(49, 61, 77);
            textBox1.ScrollBars = ScrollBars.None;
            execute();
            string ramModel = "wmic memorychip get devicelocator";

            string text = "";
            File.WriteAllText(bat, ramModel);
            text.Replace(bat, ramModel);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            string[] sep = { "devicelocator", " " };
            Int32 count = 5;
            string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in str)
            { textBox1.Text = "RAM TÜRÜ" + s; }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(49, 61, 77);
            button2.BackColor = Color.FromArgb(49, 61, 77);
            button6.BackColor = Color.FromArgb(74, 92, 117);
            button4.BackColor = Color.FromArgb(49, 61, 77);
            button5.BackColor = Color.FromArgb(49, 61, 77);
            button3.BackColor = Color.FromArgb(49, 61, 77);
            button7.BackColor = Color.FromArgb(49, 61, 77);
            button8.BackColor = Color.FromArgb(49, 61, 77);
            button9.BackColor = Color.FromArgb(49, 61, 77);
            button10.BackColor = Color.FromArgb(49, 61, 77);
            button11.BackColor = Color.FromArgb(49, 61, 77);
            textBox1.ScrollBars = ScrollBars.None;
            execute();
            string ramCapacity = "wmic MEMORYCHIP get Capacity";

            string text = "";
            File.WriteAllText(bat, ramCapacity);
            text.Replace(bat, ramCapacity);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            string[] sep = { "Capacity", " " };
            Int32 count = 5;
            string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in str)
            {
                textBox1.Text = str[4];
                //MessageBox.Show(str[4]);

                Match match = Regex.Match(str[4], @"(\d+)");

                double parse = double.Parse(match.Groups[0].Value);
                double ss = double.Parse(match.Groups[1].Value);
                ss = (ss / (1024 * 1024 * 1024));
                parse = (parse / (1024 * 1024 * 1024));
                textBox1.Text = (parse + ss).ToString() + " " + "GB";
                //textBox1.Text ="RAM KAPASİTESİ"+s;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(49, 61, 77);
            button2.BackColor = Color.FromArgb(49, 61, 77);
            button7.BackColor = Color.FromArgb(74, 92, 117);
            button4.BackColor = Color.FromArgb(49, 61, 77);
            button5.BackColor = Color.FromArgb(49, 61, 77);
            button6.BackColor = Color.FromArgb(49, 61, 77);
            button3.BackColor = Color.FromArgb(49, 61, 77);
            button8.BackColor = Color.FromArgb(49, 61, 77);
            button9.BackColor = Color.FromArgb(49, 61, 77);
            button10.BackColor = Color.FromArgb(49, 61, 77);
            button11.BackColor = Color.FromArgb(49, 61, 77);
            textBox1.ScrollBars = ScrollBars.None;
            execute();
            string ramManufacturer = "wmic MemoryChip get manufacturer";

            string text = "";
            File.WriteAllText(bat, ramManufacturer);
            text.Replace(bat, ramManufacturer);
            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            string[] sep = { "Manufacturer", " " };
            Int32 count = 6;
            string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in str)
            { textBox1.Text = "RAM MARKASI" + s; }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(49, 61, 77);
            button2.BackColor = Color.FromArgb(49, 61, 77);
            button8.BackColor = Color.FromArgb(74, 92, 117);
            button4.BackColor = Color.FromArgb(49, 61, 77);
            button5.BackColor = Color.FromArgb(49, 61, 77);
            button6.BackColor = Color.FromArgb(49, 61, 77);
            button7.BackColor = Color.FromArgb(49, 61, 77);
            button3.BackColor = Color.FromArgb(49, 61, 77);
            button9.BackColor = Color.FromArgb(49, 61, 77);
            button10.BackColor = Color.FromArgb(49, 61, 77);
            button11.BackColor = Color.FromArgb(49, 61, 77);
            textBox1.ScrollBars = ScrollBars.None;
            execute();
            string hddModel = "wmic diskdrive get Model,mediaType"; //???

            string text = "";
            File.WriteAllText(bat, hddModel);
            text.Replace(bat, hddModel);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            string[] sep = { "Model", " " };
            Int32 count = 6;
            string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in str)
            { textBox1.Text = "DISK MODEL" + s; }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(49, 61, 77);
            button2.BackColor = Color.FromArgb(49, 61, 77);
            button9.BackColor = Color.FromArgb(74, 92, 117);
            button4.BackColor = Color.FromArgb(49, 61, 77);
            button5.BackColor = Color.FromArgb(49, 61, 77);
            button6.BackColor = Color.FromArgb(49, 61, 77);
            button7.BackColor = Color.FromArgb(49, 61, 77);
            button8.BackColor = Color.FromArgb(49, 61, 77);
            button3.BackColor = Color.FromArgb(49, 61, 77);
            button10.BackColor = Color.FromArgb(49, 61, 77);
            button11.BackColor = Color.FromArgb(49, 61, 77);
            textBox1.ScrollBars = ScrollBars.None;
            execute();
            string hddCapacity = "wmic diskdrive get size";

            string text = "";
            File.WriteAllText(bat, hddCapacity);
            text.Replace(bat, hddCapacity);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            string[] sep = { "size", " " };
            int count = 5;
            string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);



            foreach (var s in str)
            {
                textBox1.Text = str[4];

                Match match = Regex.Match(str[4], @"(\d+)");
                double parse = double.Parse(match.Groups[1].Value);
                parse = Math.Round((parse / (1024 * 1024 * 1024)), 2);

                textBox1.Text = parse.ToString() + " " + "GB";

            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
        //    textBox1.ScrollBars = ScrollBars.Vertical;
            button1.BackColor = Color.FromArgb(49, 61, 77);
            button2.BackColor = Color.FromArgb(49, 61, 77);
            button10.BackColor = Color.FromArgb(74, 92, 117);
            button4.BackColor = Color.FromArgb(49, 61, 77);
            button5.BackColor = Color.FromArgb(49, 61, 77);
            button6.BackColor = Color.FromArgb(49, 61, 77);
            button7.BackColor = Color.FromArgb(49, 61, 77);
            button8.BackColor = Color.FromArgb(49, 61, 77);
            button9.BackColor = Color.FromArgb(49, 61, 77);
            button3.BackColor = Color.FromArgb(49, 61, 77);
            button11.BackColor = Color.FromArgb(49, 61, 77);


            radioButton1.Visible = true;
            radioButton2.Visible = true;
            radioButton3.Visible = true;
            radioButton4.Visible = true;
            radioButton5.Visible = true;
            radioButton6.Visible = true;


            //execute();

            //string officekey = "cd /Program Files/Microsoft Office/Office16 & cscript ospp.vbs /dstatus ";
            //string text = "";
            //File.WriteAllText(bat, officekey);
            //text.Replace(bat, officekey);

            //StreamReader reader = process.StandardOutput;
            //string output = reader.ReadToEnd();

            //textBox1.Text = output;

            //if (output.Contains("bulunamyor."))
            //{

            //    execute();
            //    officekey = "cd C:/Program Files (x86)/Microsoft Office/Office16 & cscript ospp.vbs /dstatus ";


            //    text = "";
            //    File.WriteAllText(bat, officekey);
            //    text.Replace(bat, officekey);
            //    StreamReader readerr = process.StandardOutput;

            //    output = readerr.ReadToEnd();

            //    textBox1.Text = output;
            //}
            //else if (output.Contains("PRODUCT ID:"))
            //{

            //    execute();
            //    text = "";
            //    File.WriteAllText(bat, officekey);
            //    text.Replace(bat, officekey);
            //    StreamReader readerrr = process.StandardOutput;

            //    output = readerrr.ReadToEnd();
            //    string[] sep = { "Processing", "-" };
            //    Int32 count = 3;
            //    string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);
            //    foreach (string s in str)
            //    { textBox1.Text = s; }

            //    //textBox1.Text = output;

            //}
            //else
            //{
            //    MessageBox.Show("Bilgisayarınızda office yoktur");
            //}
        }

        private void panel2_MouseDown_1(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(Handle, 0x112, 0xf012, 0);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(49, 61, 77);
            button2.BackColor = Color.FromArgb(49, 61, 77);
            button11.BackColor = Color.FromArgb(74, 92, 117);
            button4.BackColor = Color.FromArgb(49, 61, 77);
            button5.BackColor = Color.FromArgb(49, 61, 77);
            button6.BackColor = Color.FromArgb(49, 61, 77);
            button7.BackColor = Color.FromArgb(49, 61, 77);
            button8.BackColor = Color.FromArgb(49, 61, 77);
            button9.BackColor = Color.FromArgb(49, 61, 77);
            button10.BackColor = Color.FromArgb(49, 61, 77);
            button3.BackColor = Color.FromArgb(49, 61, 77);
            textBox1.ScrollBars = ScrollBars.None;
            execute();
            string guı = "wmic path win32_VideoController get name ";

            string text = "";
            File.WriteAllText(bat, guı);
            text.Replace(bat, guı);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            string[] sep = { "name", " " };
            Int32 count = 6;
            string[] str = output.Split(sep, count, StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in str)
            { textBox1.Text = "EKRAN KARTI" + s; }
        }

        private void exitLbl_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
                
            textBox1.ScrollBars = ScrollBars.Vertical;
            execute();

            string officekey = "cd /Program Files/Microsoft Office/Office16 & cscript ospp.vbs /dstatus ";
            string text = "";
            File.WriteAllText(bat, officekey);
            text.Replace(bat, officekey);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            textBox1.Text = output;
        }
       
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            textBox1.ScrollBars = ScrollBars.Vertical;
            execute();

            string officekey = "cd C:/Program Files (x86)/Microsoft Office/Office16 & cscript ospp.vbs /dstatus";
            string text = "";
            File.WriteAllText(bat, officekey);
            text.Replace(bat, officekey);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            textBox1.Text = output;
            
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            
            textBox1.ScrollBars = ScrollBars.Vertical;
            execute();

            string officekey = "cd C:/Program Files/Microsoft Office/Office15 & cscript ospp.vbs /dstatus";
         

            string text = "";
            File.WriteAllText(bat, officekey);
            text.Replace(bat, officekey);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            textBox1.Text = output;
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            textBox1.ScrollBars = ScrollBars.Vertical;
            execute();

            string officekey = "cd C:/Program Files (x86)/Microsoft Office/Office15 & cscript ospp.vbs /dstatus";
            string text = "";
            File.WriteAllText(bat, officekey);
            text.Replace(bat, officekey);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            textBox1.Text = output;
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            textBox1.ScrollBars = ScrollBars.Vertical;
            execute();

            string officekey = "cd C:/Program Files/Microsoft Office/Office14 & cscript ospp.vbs /dstatus";
            string text = "";
            File.WriteAllText(bat, officekey);
            text.Replace(bat, officekey);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            textBox1.Text = output;
           
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            textBox1.ScrollBars = ScrollBars.Vertical;
            execute();

            string officekey = "cd C:/Program Files (x86)/Microsoft Office/Office14 & cscript ospp.vbs /dstatus";
            string text = "";
            File.WriteAllText(bat, officekey);
            text.Replace(bat, officekey);

            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            textBox1.Text = output;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            execute();
            string all = "wmic MemoryChip get manufacturer & wmic baseboard get product,Manufacturer  & wmic csproduct get identifyingnumber & wmic memorychip get devicelocator & wmic MEMORYCHIP get Capacity & wmic diskdrive get Model & wmic diskdrive get size & wmic diskdrive get size & wmic csproduct get name & wmic cpu get name & wmic path win32_VideoController get name & wmic path softwarelicensingservice get OA3xOriginalProductKey";
            string text = "";
            File.WriteAllText(bat, all);
            text.Replace(bat, all);
            StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            string fix = Regex.Replace(output, @"^\s*$\n", string.Empty, RegexOptions.Multiline);
            int s = fix.LastIndexOf("get OA3xOriginalProductKey");
            string res = fix.Remove(0, s + "get OA3xOriginalProductKey".Length);
            progressBar1.Value = 20;
            //int s = output.LastIndexOf("OA3xOriginalProductKey");
            //string res = output.Remove(0, s);
            //string fix = Regex.Replace(res, @"^\s*$\n", string.Empty, RegexOptions.Multiline);
            //çıktı split edidi
            string[] arr = res.Split(new string[] { "Manufacturer" }, StringSplitOptions.RemoveEmptyEntries);
            string[] arr2 = res.Split(new string[] { "IdentifyingNumber" }, StringSplitOptions.RemoveEmptyEntries);
            string[] arr3 = res.Split(new string[] { "DeviceLocator" }, StringSplitOptions.RemoveEmptyEntries);
            string[] arr4 = res.Split(new string[] { "Capacity" }, StringSplitOptions.RemoveEmptyEntries);
            string[] arr5 = res.Split(new string[] { "Model" }, StringSplitOptions.RemoveEmptyEntries);
            string[] arr6 = res.Split(new string[] { "Size" }, StringSplitOptions.RemoveEmptyEntries);
            string[] arr7 = res.Split(new string[] { "Name" }, StringSplitOptions.RemoveEmptyEntries);
            string[] arr9 = res.Split(new string[] { "OA3xOriginalProductKey" }, StringSplitOptions.RemoveEmptyEntries);

            string ramManu = arr[0];
         
            int b1 = arr2[0].IndexOf("Product") + "Product".Length;
            string mbModel = arr2[0].Substring(b1);
        
            int b2 = arr3[0].IndexOf("IdentifyingNumber") + "IdentifyingNumber".Length;
            string serial = arr3[0].Substring(b2);

            int b3 = arr4[0].IndexOf("DeviceLocator") + "DeviceLocator".Length;
            string ram = arr4[0].Substring(b3);

            int b4 = arr5[0].IndexOf("Capacity") + "Capacity".Length;         
            string ramCapacity = arr5[0].Substring(b4);
            Match match = Regex.Match(ramCapacity, @"(\d+)");
            double parse = double.Parse(match.Groups[0].Value);
            double ss = double.Parse(match.Groups[1].Value);
            ss = (ss / (1024 * 1024 * 1024));
            parse = (parse / (1024 * 1024 * 1024));
            string rCapacity = (ss + parse).ToString();

            string system = arr7[1];

            string cpu = arr7[2];

            int last = arr7[3].LastIndexOf("OA3xOriginalProductKey");
            string display = arr7[3].Substring(0, last);
 
            string diskSize = arr6[1];
            Match match2 = Regex.Match(diskSize, @"(\d+)");
            double dSize = double.Parse(match2.Groups[1].Value);
            dSize = Math.Round((dSize / (1024 * 1024 * 1024)), 2);
         


            int b5 = arr6[0].IndexOf("Model") + "Model".Length;
            string diskModel = arr6[0].Substring(b5);
            

            string win = arr9[1];
            progressBar1.Value = 45;

            //int b = fix.IndexOf("OA3xOriginalProductKey");
            //int end = fix.LastIndexOf("IdentifiyngNumber");
            //string win = fix.Remove(b,s);
            //MessageBox.Show(win);


            string date = DateTime.Now.ToShortDateString();
            string time = DateTime.Now.ToShortTimeString();

          

            //veriler pdf'e aktarıldı
            string file = "donanım.pdf";
            FileStream fs = new FileStream(Path.Combine(AppDomain.CurrentDomain.BaseDirectory + file), FileMode.Create);

            iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4);

            PdfWriter write = PdfWriter.GetInstance(doc, fs);
            BaseFont bF = BaseFont.CreateFont("C:\\windows\\fonts\\arial.ttf", "windows-1254", true);
            iTextSharp.text.Font font = new iTextSharp.text.Font(bF, 12f, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font font2 = new iTextSharp.text.Font(bF, 11f, iTextSharp.text.Font.BOLD);
            doc.Open();

           

            PdfPTable tbl = new PdfPTable(1);
            tbl.DefaultCell.BorderWidth = 0;
            tbl.WidthPercentage = 100;
            tbl.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
            Chunk tittle = new Chunk("BİLGiSAYARIN TEKNİK ÖZELLİKLERİ", FontFactory.GetFont("Arial"));
            tittle.Font = font;
            tittle.Font.Color = new iTextSharp.text.BaseColor(0, 0, 0);
            tittle.Font.SetStyle(3);
            tittle.Font.Size = 14;
            Phrase p1 = new Phrase();
            p1.Add(tittle);
            tbl.AddCell(p1);
            doc.Add(tbl);


            PdfPTable tbl4 = new PdfPTable(1);
            tbl4.DefaultCell.BorderWidth = 0;
            tbl4.WidthPercentage = 100;
            PdfPCell cell16 = new PdfPCell(new Phrase(date + "  " + time));
            cell16.HorizontalAlignment = Element.ALIGN_RIGHT;
            cell16.VerticalAlignment = Element.ALIGN_TOP;
            cell16.MinimumHeight = 20;
            cell16.Border = 0;
            tbl4.AddCell(cell16);
            doc.Add(tbl4);

            PdfPTable tbl1 = new PdfPTable(2);
            tbl1.DefaultCell.Padding = 5;
            tbl1.WidthPercentage = 100;
            tbl1.DefaultCell.BorderWidth = 0.5f;
            tbl1.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;
            PdfPCell cell = new PdfPCell(new Phrase("ÜRÜNLER", font2));
            cell.BackgroundColor = new iTextSharp.text.BaseColor(22, 85, 160);
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            tbl1.AddCell(cell);
            PdfPCell cell2 = new PdfPCell(new Phrase("TEKNİK ÖZELLİKLER", font2));
            cell2.BackgroundColor = new iTextSharp.text.BaseColor(22, 85, 160);
            cell2.HorizontalAlignment = Element.ALIGN_CENTER;
            tbl1.AddCell(cell2);
            tbl1.AddCell(new Phrase("WINDOWS ÜRÜN ANAHTARI", font2));
            PdfPCell cell3 = new PdfPCell(new Phrase(win.Trim()));
            cell3.VerticalAlignment = Element.ALIGN_CENTER;
            tbl1.AddCell(cell3);
            tbl1.AddCell(new Phrase("OFFICE ÜRÜN ANAHTARI", font2));
            PdfPCell cell4 = new PdfPCell(new Phrase(""));
            tbl1.AddCell(cell4);
            tbl1.AddCell(new Phrase("ANAKART MODELİ", font2));
            PdfPCell cell5 = new PdfPCell(new Phrase(mbModel.Trim()));
            tbl1.AddCell(cell5);
            tbl1.AddCell(new Phrase("SİSTEM MODELİ", font2));
            PdfPCell cell6 = new PdfPCell(new Phrase(system.Trim()));
            tbl1.AddCell(cell6);
            tbl1.AddCell(new Phrase("İŞLEMCİ ADI", font2));
            PdfPCell cell7 = new PdfPCell(new Phrase(cpu.Trim()));
            tbl1.AddCell(cell7);
            tbl1.AddCell(new Phrase("RAM TÜRÜ", font2));
            PdfPCell cell8 = new PdfPCell(new Phrase(ram.Trim()));
            tbl1.AddCell(cell8);
            tbl1.AddCell(new Phrase("RAM KAPASİTESİ", font2));
            PdfPCell cell9 = new PdfPCell(new Phrase(rCapacity.Replace("\r", string.Empty).Replace("\n", string.Empty)+" GB"));
            tbl1.AddCell(cell9);
            tbl1.AddCell(new Phrase("RAM MARKASI", font2));
            PdfPCell cell10 = new PdfPCell(new Phrase(ramManu.Trim()));
            tbl1.AddCell(cell10);
            tbl1.AddCell(new Phrase("DİSK", font2));
            PdfPCell cell11 = new PdfPCell(new Phrase(dSize+" GB"));
            tbl1.AddCell(cell11);
            tbl1.AddCell(new Phrase("DİSK MODEL", font2));
            PdfPCell cell14 = new PdfPCell(new Phrase(diskModel.Trim()));
            tbl1.AddCell(cell14);
            tbl1.AddCell(new Phrase("EKRAN KARTI", font2));
            PdfPCell cell12 = new PdfPCell(new Phrase(display.Trim()));
            tbl1.AddCell(cell12);
            tbl1.AddCell(new Phrase("ANAKART SERİ NUMARASI", font2));
            PdfPCell cell13 = new PdfPCell(new Phrase(serial.Trim()));
            cell13.VerticalAlignment = Element.ALIGN_CENTER;
            doc.Add(tbl1);

     
            doc.Close();
            fs.Close();
            progressBar1.Value = 75;
            //pdf excel'e dönüştürüldü
            string pdfFile = file;
            string excelFile = System.IO.Path.ChangeExtension(pdfFile, ".xls");

            PdfFocus f = new PdfFocus();
            f.ExcelOptions.ConvertNonTabularDataToSpreadsheet = true;

            f.ExcelOptions.PreservePageLayout = true;

            f.OpenPdf(pdfFile);

            if (f.PageCount > 0)
            {
                f.ToExcel(excelFile);
            }
            progressBar1.Value = 100;
            //label3.Text = "Pdf ve Excel Formatında Kaydedildi.";



            //string file = "donanım.pdf";            
            //FileStream fs = new FileStream (file, FileMode.Create);
            ////PdfPTable pdf1 = new PdfPTable(2);
            ////pdf1.AddCell("assd");
            //Document doc = new Document(PageSize.A4);
            //PdfWriter write = PdfWriter.GetInstance(doc,fs);
            //doc.Open();
            //Paragraph p1 = new Paragraph(fix);
            //doc.Add(p1);
            //MessageBox.Show("PDF formatında kaydedildi");
            ////doc.Add(pdf1);
            //doc.Close();








        //    string ss = "dkflkggıkggokb";
        //    PdfDocument doc = new PdfDocument();
        //    PdfPageBase page = doc.Pages.Add(PdfPageSize.A4);
        //    //???????
        //    PdfList asd = new PdfList(ss);

            

            //Table table = new Table(8);
            //for (int i = 0; i < 16; i++)
            //{
            //    table.addCell("hi");
            //}
            //doc.add(table);

           

            //doc.SaveToFile("Donanım.pdf");
            //doc.Close();
            //System.Diagnostics.Process.Start("Donanım.pdf");


            

        }

    }
}

