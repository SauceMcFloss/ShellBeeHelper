using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace ShellBeeHelper.Tabs
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class VideoSheetConverterTab : UserControl
    {
        #region Constructor

        public VideoSheetConverterTab(Logger log)
        {
            InitializeComponent();

            Log = log;
        }

        #endregion

        #region Variables

        private Logger _Log = null;
        public Logger Log
        {
            get { return _Log; }
            set { _Log = value; }
        }

        OpenFileDialog OFD = new OpenFileDialog()
        {
            InitialDirectory = "C:\\Users\\" + Environment.UserName + "\\Desktop",
            Filter = "Excel |*.xlsx",
            RestoreDirectory = true,
            Multiselect = false,
        };

        string ConvertedPath = "";

        #endregion

        #region Events

        private void FindButton_Click(object sender, RoutedEventArgs e)
        {
            if(OFD.ShowDialog() == true)
            {
                VideoSheetTextBox.Text = OFD.FileName;
                ConvertButton.IsEnabled = true;
                ConvertedPath = Path.GetDirectoryName(OFD.FileName);
            }
            else
            {
                VideoSheetTextBox.Text = "*missing*";
                ConvertButton.IsEnabled = false;
                ConvertedPath = "";
            }
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application eApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                eApp = new Excel.Application();
                workbook = eApp.Workbooks.Open(OFD.FileName);
                worksheet = workbook.ActiveSheet;

                #region CSVs

                using (StreamWriter Products = new StreamWriter(ConvertedPath.ToString() + "\\1 - Products.csv", true))
                {
                    using (StreamWriter Assets = new StreamWriter(ConvertedPath + "\\2 - Assets.csv"))
                    {
                        using (StreamWriter Products_Assets_Assigner = new StreamWriter(ConvertedPath + "\\3 - Products_Assets_Assigner.csv", true))
                        {
                            using (StreamWriter Contract_Assigner = new StreamWriter(ConvertedPath + "\\4 - Contract_Assigner.csv", true))
                            {
                                using (StreamWriter Automatch_to_contract_SEND_TO_FUGA = new StreamWriter(ConvertedPath + "\\5 - Automatch_to_contract_SEND_TO_FUGA.csv", true))
                                {
                                    Products.WriteLine("product_reference,barcode,title,artist,version,catalog_number,subgenre,genre,label_name,configuration_name,accounting_group_name", true);
                                    Assets.WriteLine("asset_reference,title,artist,version,isrc,duration,genre,subgenre,accounting_group_name");
                                    Products_Assets_Assigner.WriteLine("asset_reference,product_reference,share");
                                    Contract_Assigner.WriteLine("asset_isrc,contract_reference,share");
                                    Automatch_to_contract_SEND_TO_FUGA.WriteLine("barcode,label,automatch_id");

                                    int row = 2;
                                    while (!String.IsNullOrWhiteSpace((string)worksheet.Cells[row, 1].Value))
                                    {
                                        Products.WriteLine("\"{0}\",\"{0}\",\"{1}\",\"{2}\",,\"{0}\",,,\"{3}\",,",
                                            (string)worksheet.Cells[row, 4].Value, // catalog-number
                                            (string)worksheet.Cells[row, 1].Value, // album-title
                                            (string)worksheet.Cells[row, 2].Value, // primary-album-artist
                                            (string)worksheet.Cells[row, 3].Value // label
                                            );
                                        Assets.WriteLine("\"{0}\",\"{1}\",\"{2}\",,\"{0}\",,,,",
                                            (string)worksheet.Cells[row, 7].Value, // isrc
                                            (string)worksheet.Cells[row, 1].Value, // album-title
                                            (string)worksheet.Cells[row, 2].Value // primary-album-artist
                                            );
                                        Products_Assets_Assigner.WriteLine("\"{0}\",\"{1}\",1",
                                            (string)worksheet.Cells[row, 7].Value, // isrc
                                            (string)worksheet.Cells[row, 4].Value // catalog-number
                                            );
                                        Contract_Assigner.WriteLine("\"{0}\",\"{1}\",1",
                                            (string)worksheet.Cells[row, 7].Value, // isrc
                                            (string)worksheet.Cells[row, 3].Value // label
                                            );
                                        Automatch_to_contract_SEND_TO_FUGA.WriteLine("\"{0}\",\"{1}\",\"{1}\"",
                                            (string)worksheet.Cells[row, 4].Value, // catalog-number
                                            (string)worksheet.Cells[row, 3].Value // label
                                            );

                                        row++;
                                    }
                                }
                            }
                        }
                    }
                }
                
                #endregion
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error Setting up or leaving Excel. Idk, maybe it's a CSV issue?");
                return;
            }

            // Release Excel
            try
            {
                eApp.Quit();

                worksheet = null;
                workbook = null;
                eApp = null;
            }
            catch
            {

            }
        }

        #endregion

        #region Methods and Functions

        #endregion
    }
}
