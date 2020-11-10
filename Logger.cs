using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ShellBeeHelper
{
    public class Logger
    {
        #region Constructors

        public Logger()
        {

        }

        #endregion

        #region Variables

        private TextBlock _LogBox = null;
        public TextBlock LogBox
        {
            get { return _LogBox; }
            set { _LogBox = value; }
        }

        public string LogPath = "C:\\Users\\" + Environment.UserName + "\\Desktop\\ShellBeeHelper_Logs.txt";

        #endregion

        #region Methods and Functions

        public void Error(Exception ex = null, string extraLogs = "")
        {
            LogBox.Text += DateTime.Now.ToString() + ((ex != null) ? ":\t" + ex.ToString() + "\n" : "");
            LogBox.Text += String.IsNullOrWhiteSpace(extraLogs) ? "" : "\t" + extraLogs + "\n";

            using (StreamWriter sw = File.Exists(LogPath) ? File.AppendText(LogPath) : File.CreateText(LogPath))
            {
                sw.WriteLine(DateTime.Now.ToString() + ((ex != null) ? ":\t" + ex.ToString() + "\n" : ""));
                sw.WriteLine(String.IsNullOrWhiteSpace(extraLogs) ? "" : "\t" + extraLogs + "\n");
            }
        }

        public void Log(string logs = "")
        {
            LogBox.Text += DateTime.Now.ToString() + ":\t" + logs + "\n";
        }

        #endregion
    }
}
