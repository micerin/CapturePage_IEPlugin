using System;
using System.ComponentModel;
using System.Windows.Forms;

using BandObjectLib;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using SHDocVw;
using System.Text;

namespace CaptureWebPage
{

    [Guid("AE07101B-46D4-4a98-AF68-0333EA26E113")]
    [BandObject("CapturePage", BandObjectStyle.Horizontal | BandObjectStyle.ExplorerToolbar | BandObjectStyle.TaskbarToolBar, HelpText = "Capture web page.")]
    public class CapturePage : BandObject
    {
        private System.Windows.Forms.Button button1;
        private System.ComponentModel.Container components = null;

        [DllImport("user32")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32")]
        public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        static extern IntPtr GetFocus();

        [DllImport("user32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(IntPtr hwnd, int msg, int wParam, StringBuilder sb);

        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string _ClassName, string _WindowName);

        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        private static extern IntPtr FindWindowEx(IntPtr _Parent, IntPtr childAfter, string _ClassName, string _WindowName);

        public const int WM_GETTEXTLENGTH = 0x000E;
        public const int WM_GETTEXT = 0x000D;

        public CapturePage()
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                    components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.Highlight;
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.Lime;
            this.button1.FlatAppearance.BorderSize = 3;
            this.button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Lime;
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.button1.Font = new System.Drawing.Font("Microsoft YaHei UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.SystemColors.Menu;
            this.button1.Location = new System.Drawing.Point(19, 0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(95, 24);
            this.button1.TabIndex = 0;
            this.button1.Text = "CapturePage";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // HelloWorldBar
            // 
            this.Controls.Add(this.button1);
            this.MinSize = new System.Drawing.Size(150, 24);
            this.Name = "HelloWorldBar";
            this.Size = new System.Drawing.Size(150, 24);
            this.Title = "Hello Bar";
            this.ResumeLayout(false);

        }
        #endregion

        private void button1_Click(object sender, System.EventArgs e)
        {
            //var searchContent = this.textBox1.Text;
            //if (searchContent.Equals(null) || searchContent.Equals(string.Empty))
            //{
            //    Process.Start("www.bing.com");
            //}
            //else
            //{
            //    Process.Start(string.Format("http://www.bing.com/search?q={0}", searchContent));
            //}

            //MessageBox.Show(searchContent);

            string strURL = GetActiveUrl();
            CaptureWebpageToPic(strURL);
        }

        public static void CaptureWebpageToPic(string strURL)
        {
            string picName = DateTime.UtcNow.ToFileTime().ToString();
            string t_strSaveFolder = "c:\\tests";
            string t_strLargeImage = t_strSaveFolder + "\\" + picName + ".png";
            WebsiteToImage websiteToImage = new WebsiteToImage(strURL, t_strLargeImage);
            websiteToImage.Generate();
        }

        //private static void SaveWebPageByACAWebThumb(string strURL)
        //{
        //    string picName = DateTime.UtcNow.ToFileTime().ToString();
        //    string t_strSaveFolder = "c:\\tests";
        //    string t_strLargeImage = t_strSaveFolder + "\\" + picName + ".png";

        //    // Create instance
        //    ThumbMakerClass t_xThumbMaker = new ACAWebThumbLib.ThumbMakerClass();

        //    t_xThumbMaker.SetURL("strURL");
        //    t_xThumbMaker.StartSnap();

        //    // Save the captured web page to image with full size in C#
        //    t_xThumbMaker.SaveImage(t_strLargeImage);

        //    MessageBox.Show(t_strLargeImage + " done!");
        //}

        private static void SaveWebPageByWebSiteScreenShot(string strURL)
        {
            WebsitesScreenshot.WebsitesScreenshot _Obj;
            _Obj = new WebsitesScreenshot.WebsitesScreenshot();
            WebsitesScreenshot.WebsitesScreenshot.Result _Result;
            _Result = _Obj.CaptureWebpage(strURL);
            string picName = DateTime.UtcNow.ToFileTime().ToString();
            if (_Result == WebsitesScreenshot.WebsitesScreenshot.Result.Captured)
            {
                _Obj.ImageFormat = WebsitesScreenshot.
                    WebsitesScreenshot.ImageFormats.JPG;
                _Obj.BrowserHeight = 1500;
                _Obj.BrowserWidth = 800;
                _Obj.SaveImage(string.Format("c:\\tests\\{0}.jpg", picName));
            }
            _Obj.Dispose();
        }

        public static string GetActiveUrl()
        {
            string strURL = string.Empty;
            var childHandle = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "IEFrame", null);
            if (childHandle != IntPtr.Zero)
            {
                //get the handle to the address bar on IE
                childHandle = FindWindowEx(childHandle, IntPtr.Zero, "WorkerW", null);
                if (childHandle != IntPtr.Zero)
                {
                    //get handle to edit
                    childHandle = FindWindowEx(childHandle, IntPtr.Zero, "ReBarWindow32", null);
                    if (childHandle != IntPtr.Zero)
                    {
                        // get a handle to comboBoxEx32
                        childHandle = FindWindowEx(childHandle, IntPtr.Zero, "Address Band Root", null);
                        if (childHandle != IntPtr.Zero)
                        {
                            // get a handle to combo box
                            childHandle = FindWindowEx(childHandle, IntPtr.Zero, "Edit", null);
                            if (childHandle != IntPtr.Zero)
                            {
                                // now to get the URL we need to get the Text - but first get the length of the URL
                                StringBuilder sb = new StringBuilder();
                                int length = SendMessage(childHandle, WM_GETTEXTLENGTH, 0, sb);
                                length += 1;    // because the length returned above included 0
                                StringBuilder text = new StringBuilder(length); // need stringbuilder - not string
                                int hr = SendMessage(childHandle, WM_GETTEXT, length, text); // get the URL
                                strURL = text.ToString();
                            }
                        }
                    }
                }
            }
            return strURL;
        }
    }
}
