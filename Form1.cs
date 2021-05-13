using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using SkiaSharp.QrCode.Image;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Zack.ComObjectHelpers;

namespace PhoneAsPrompter
{
    public partial class Form1 : Form
    {
        private const int port = 7999;

        private IWebHost webHost;

        private dynamic presentation;

        private string ip = "";

        private COMReferenceTracker comReference = new COMReferenceTracker();

        public Form1()
        {
            InitializeComponent();
            ShowUrl();
            // 配置服务器
            this.webHost = new WebHostBuilder()
                .UseKestrel()
                .Configure(ConfigureWebApp)
                .UseUrls("http://*:" + port)
                .Build();

            // 异步运行服务器
            this.webHost.RunAsync();
            

            // 关闭窗口处理
            this.FormClosed += Form1_FormClosed;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            // 关闭所有 COM　对象，以及当前打开的PPT
            ClearComRefs();
            // 停止运行服务器
            this.webHost.StopAsync();
            this.webHost.WaitForShutdown();
            Process.GetCurrentProcess().Kill();
        }

        private void ShowUrl()
        {
            this.ip = "http://";
            string name = Dns.GetHostName();
            IPAddress[] ipadrlist = Dns.GetHostAddresses(name);
            foreach (IPAddress ipa in ipadrlist)
            {
                if (ipa.AddressFamily == AddressFamily.InterNetwork)
                {
                    ip += ipa.ToString() + ":" + port;
                    break;
                }
            }
           

            this.urlLable.Text = "请扫描二维码或者用浏览器打开：" + ip;
            this.urlLable.Links.Add(15, ip.Length, ip);
            var qrcode = new QrCode(ip, new Vector2Slim(256, 256), SkiaSharp.SKEncodedImageFormat.Png);
            using (MemoryStream stream = new MemoryStream())
            {
                qrcode.GenerateImage(stream);
                stream.Position = 0;
                imgQRCode.SizeMode = PictureBoxSizeMode.Zoom;
                imgQRCode.Image = Image.FromStream(stream);
            }

        }

        private void ConfigureWebApp(IApplicationBuilder app)
        {
            app.UseDefaultFiles();
            app.UseStaticFiles();
            app.Run(async (context) =>
            {
                // 处理非静态请求 
                var request = context.Request;
                var response = context.Response;
                string path = request.Path.Value;
                response.ContentType = "application/json; charset=UTF-8";
                if (path == "/report")
                {
                    string value = request.Query["value"];
                    this.BeginInvoke(new Action(() => {
                        this.PageLabel.Text = value;
                    }));
                    response.StatusCode = 200;
                    await response.WriteAsync("ok");
                }
                else if (path == "/getNote")
                {
                    string notesText = null;
                    this.Invoke(new Action(() => {
                        if (this.presentation == null)
                        {
                            return;
                        }
                        try
                        {
                            dynamic notesPage = T(T(T(T(presentation.SlideShowWindow).View).Slide).NotesPage);
                            notesText = GetInnerText(notesPage);
                        }
                        catch (COMException ex)
                        {
                            notesText = "";
                        }
                    }));
                    await response.WriteAsync(notesText);
                }
                else if (path == "/next")
                {
                    response.StatusCode = 200;
                    this.Invoke(new Action(() => {
                        if (this.presentation == null)
                        {
                            return;
                        }
                        T(T(this.presentation.SlideShowWindow).View).Next();
                    }));
                    await response.WriteAsync("OK");
                }
                else if (path == "/previous")
                {
                    response.StatusCode = 200;
                    this.Invoke(new Action(() => {
                        if (this.presentation == null)
                        {
                            return;
                        }
                        T(T(this.presentation.SlideShowWindow).View).Previous();
                    }));
                    await response.WriteAsync("OK");
                }
                else
                {
                    response.StatusCode = 404;
                }
            });
            
        }


        private string GetInnerText(dynamic part)
        {
            StringBuilder sb = new StringBuilder();
            dynamic shapes = T(T(part).Shapes);
            int shapesCount = shapes.Count;
            for (int i = 0; i < shapesCount; i++)
            {
                dynamic shape = T(shapes[i + 1]);
                var textFrame = T(shape.TextFrame);
                // MsoTriState.msoTrue==-1
                if (textFrame.HasText == -1)
                {
                    string text = T(textFrame.TextRange).Text;
                    sb.AppendLine(text);
                }
                sb.AppendLine();
            }
            return sb.ToString();
        }

        private void ClearComRefs()
        {
            try
            {
                if (this.presentation != null)
                {
                    T(this.presentation.Application).Quit();
                    this.presentation = null;
                }
            }
            catch (COMException ex)
            {
                Debug.WriteLine(ex);
            }
            this.comReference.Dispose();
            this.comReference = new COMReferenceTracker();
        }

        private dynamic T(dynamic comObj)
        {
            return this.comReference.T(comObj);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //// 创建 PPT 对象
            //dynamic ppt = T(PowerPointHelper.CreatePowerPointApplication());
            //// 显示 PPT
            //ppt.Visible = true;

            //dynamic presentations = T(ppt.Presentations);
            //// 打开 PPT
            //this.presentation = T(presentations.Open(@"E:\test.pptx"));
            //// 全屏显示
            //T(this.presentation.SlideShowSettings).Run();
            openFileDialog.Filter = "ppt文件|*.ppt;*.pptx;*.pptm";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }
          
            string filename = openFileDialog.FileName;
            this.ClearComRefs();
            dynamic pptApp = T(PowerPointHelper.CreatePowerPointApplication());
            pptApp.Visible = true;
            dynamic presentations = T(pptApp.Presentations);
            this.presentation = T(presentations.Open(filename));
            T(this.presentation.SlideShowSettings).Run();
        }

        /**
         * 获取当前页备注
         */
        private void button2_Click(object sender, EventArgs e)
        {
            if (this.presentation == null)
            {
                MessageBox.Show("请先选择打开一个PPT文件");
                return;
            }
            
            dynamic notesPage = T(T(T(T(presentation.SlideShowWindow).View).Slide).NotesPage);
            string notesText = GetInnerText(notesPage);
            MessageBox.Show(notesText);
        }

        /**
         * 上一个 
         */
        private void button3_Click(object sender, EventArgs e)
        {
            if (this.presentation == null)
            {
                MessageBox.Show("请先选择打开一个PPT文件");
                return;
            }
            T(T(presentation.SlideShowWindow).View).Previous();
        }

        /**
         * 下一个 
         */
        private void button4_Click(object sender, EventArgs e)
        {
            if (this.presentation == null)
            {
                MessageBox.Show("请先选择打开一个PPT文件");
                return;
            }
            T(T(presentation.SlideShowWindow).View).Next();
        }

        private void urlLable_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("explorer.exe", this.ip);
        }
    }
}
