using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Dynamics.Framework.UI.Extensibility;
using Microsoft.Dynamics.Framework.UI.Extensibility.WinForms;

namespace TTTViewer
{
    // Public key token: 764896b391714106

    [ControlAddInExport("TTTViewer")]

    public class TTTViewerControlAddIn : StringControlAddInBase
    {
        String urlString;
        WebBrowser webBrowser1;
        System.Uri webUri;

        [ApplicationVisible]
        public event MethodInvoker ControlAddInReady;

        [ApplicationVisible]
        public void ActivateBrowser()
        {
            webBrowser1.Focus();
        }

        [ApplicationVisible]
        public bool VisibleControl
        {
            get { return webBrowser1.Visible; }
            set { webBrowser1.Visible = value; }
        }

        [ApplicationVisible]
        public void SendCtrlL()
        {
            SendKeys.Send("^l");
            SendKeys.Flush();
        }

        protected override Control CreateControl()
        {
            urlString = "http://www.telosteam.dk/";

            try
            {
                webBrowser1 = new WebBrowser();
                webBrowser1.TabIndex = 0;
                // webBrowser1.WebBrowserShortcutsEnabled = false;
                webBrowser1.WebBrowserShortcutsEnabled = true;
                webBrowser1.Dock = DockStyle.Fill;
                webBrowser1.MaximumSize = new Size(int.MaxValue, int.MaxValue);
                webBrowser1.MinimumSize = new Size(640, 960);

                webUri = new Uri(urlString);
                webBrowser1.Url = webUri;
            }
            catch (Exception Exc)
            {
                throw new Exception(Exc.Message);
            }

            webBrowser1.HandleCreated += (sender, args) =>
            {
                if (ControlAddInReady != null)
                {
                    ControlAddInReady();
                }
            };
            return webBrowser1;
        }

        public override string Value
        {
            get
            {
                return base.Value;
            }
            set
            {
                if (value.ToUpper().StartsWith("FILE://"))
                {
                    if (System.IO.File.Exists(value.Remove(0, 7)))
                    {
                        webBrowser1.Navigate(value);
                    }
                    else
                    {
                        webBrowser1.Navigate("");
                    }
                }
                else
                {
                    if (value.Contains("\\"))
                    {
                        if (System.IO.File.Exists(value))
                        {
                            webBrowser1.Navigate(value);
                            // SendCtrlL();
                        }
                        else
                        {
                            webBrowser1.Navigate("");
                        }
                    }
                    else
                    {
                        webBrowser1.Navigate(value);
                    }
                }
            }
        }
    }
}
