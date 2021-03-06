﻿using ShellBeeHelper.Tabs;
using System;
using System.Windows.Controls;

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

            Log = new Logger("C:\\Users\\" + Environment.UserName + "\\Desktop\\ShellBeeHelper.log");
            Log.LogBox = LogBox;

            TabsList.Items.Add(new TabItem() { Header = "Contracts from Outlook", Content = new ContractsFromOutlookTab(Log) });
            TabsList.Items.Add(new TabItem() { Header = "Video Sheet Converter", Content = new VideoSheetConverterTab(Log) });
        }

        #endregion

        #region Variables

        private Logger _Log = null;
        public Logger Log
        {
            get { return _Log; }
            set { _Log = value; }
        }

        #endregion

        #region Events
        
        #endregion

        #region Methods and Functions

        #endregion
    }
}
