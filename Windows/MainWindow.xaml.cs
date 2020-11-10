using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Xsl;

namespace ShellBeeHelper.Windows
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        #region Constructors

        public MainWindow()
        {
            InitializeComponent();
        }

        #endregion

        #region Variables

        private Logger _Log = new Logger();
        public Logger Log
        {
            get { return _Log; }
            set { _Log = value; }
        }

        #endregion

        #region Events

        /*private void LogBox_Initialized(object sender, EventArgs e)
        {
            Logger.LogBox = LogBox;
        }*/

        
        #endregion

        #region Methods and Functions

        #endregion
    }
}
