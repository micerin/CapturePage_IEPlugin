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
        #region DllImports
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
        #endregion DllImports

        public const int WM_GETTEXTLENGTH = 0x000E;
        public const int WM_GETTEXT = 0x000D;
        private System.Windows.Forms.Button CaptureButton;
        private System.ComponentModel.Container components = null;

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
            this.CaptureButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // CaptureButton
            // 
            this.CaptureButton.BackColor = System.Drawing.SystemColors.Highlight;
            this.CaptureButton.FlatAppearance.BorderColor = System.Drawing.Color.Lime;
            this.CaptureButton.FlatAppearance.BorderSize = 3;
            this.CaptureButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Lime;
            this.CaptureButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.CaptureButton.Font = new System.Drawing.Font("Microsoft YaHei UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CaptureButton.ForeColor = System.Drawing.SystemColors.Menu;
            this.CaptureButton.Location = new System.Drawing.Point(19, 0);
            this.CaptureButton.Name = "CaptureButton";
            this.CaptureButton.Size = new System.Drawing.Size(95, 24);
            this.CaptureButton.TabIndex = 0;
            this.CaptureButton.Text = "CapturePage";
            this.CaptureButton.UseVisualStyleBackColor = false;
            this.CaptureButton.Click += new System.EventHandler(this.CaptureButton_Click);
            // 
            // CapturePage
            // 
            this.Controls.Add(this.CaptureButton);
            this.MinSize = new System.Drawing.Size(150, 24);
            this.Name = "CapturePage";
            this.Size = new System.Drawing.Size(150, 24);
            this.Title = "Hello Bar";
            this.ResumeLayout(false);

        }
        #endregion

        /// <summary>
        /// Button Clieck
        /// </summary>
        private void CaptureButton_Click(object sender, System.EventArgs e)
        {
            string strURL = GetActiveUrl();
            CaptureWebpageToPic(strURL);
        }

        /// <summary>
        /// Capture Webpage with given URL to png
        /// </summary>
        public static void CaptureWebpageToPic(string strURL)
        {
            string picName = DateTime.UtcNow.ToFileTime().ToString();
            string t_strSaveFolder = "c:\\tests";
            string t_strLargeImage = t_strSaveFolder + "\\" + picName + ".png";
            WebsiteToImage websiteToImage = new WebsiteToImage(strURL, t_strLargeImage);
            websiteToImage.Generate();
        }

        /// <summary>
        /// Get current active URL for IE
        /// </summary>
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
