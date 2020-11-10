using System;
using System.Collections.Generic;
using System.Linq;
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
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ShellBeeHelper.Tabs
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class ContractsFromOutlookTab : UserControl
    {
        public ContractsFromOutlookTab()
        {
            InitializeComponent();
        }

        Logger Log = null;

        private void ScanButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                
                List<string> icList = new List<string>();

                #region Outlook

                try
                {
                    Outlook.Application oApp = new Outlook.Application();
                    Outlook.NameSpace oNameSpace = oApp.GetNamespace("mapi");
                    oNameSpace.Logon(Missing.Value, Missing.Value, false, true);

                    Outlook.MAPIFolder contractsFolder = oNameSpace.Folders["shelby@zojakworldwide.com"].Folders["Contracts"];
                    Outlook.Items items = contractsFolder.Items;

                    try
                    {
                        foreach (Outlook.MailItem msg in items)
                        {
                            if (msg.SenderName.Contains("DocuSign"))
                            {
                                Log.Log("Encountered DocuSign.");
                                continue;
                            }
                            try
                            {
                                string importantContent = "";
                                if (msg.SenderName == "HelloSign" || msg.Subject.Contains("You've been copied on Zojak"))
                                {
                                    importantContent = msg.Body.Substring(msg.Body.IndexOf("@zojakworldwide.com") + 31);
                                    importantContent = importantContent.Substring(0, importantContent.IndexOf("@zojakworldwide.com"));
                                    importantContent = importantContent.Substring(0, importantContent.IndexOf("\r\n"));
                                }
                                else if (msg.Body.Contains("The document is being sent to:"))
                                {
                                    importantContent = msg.Body.Substring(msg.Body.IndexOf("The document is being sent to:") + 35);
                                    importantContent = importantContent.Substring(0, importantContent.IndexOf("@zojakworldwide.com"));
                                    importantContent = importantContent.Substring(0, importantContent.IndexOf("\r\n"));
                                }
                                else if (msg.Body.Contains("The document is being sent in this order:"))
                                {
                                    importantContent = msg.Body.Substring(msg.Body.IndexOf("The document is being sent in this order:") + 49);
                                    importantContent = importantContent.Substring(0, importantContent.IndexOf("@zojakworldwide.com"));
                                    importantContent = importantContent.Substring(0, importantContent.IndexOf("\r\n"));
                                }
                                else
                                {
                                    Log.Error(null, "Encountered unknown email format.\n\tFrom: " + msg.SenderName + "\n\tSubject: " + msg.Subject + "\n\tReceived at: " + msg.ReceivedTime);
                                    continue;
                                }

                                importantContent = importantContent.Trim(' ', '\r', '\n');
                                icList.Add(importantContent);
                            }
                            catch (Exception ex)
                            {
                                Log.Error(ex, "Error reading message.");
                            }
                        }

                        Log.Log("Completed scan. Found " + icList.Count + " contracts.");
                        Log.Log("Checked for duplicates. Found " + icList.Distinct().ToList().Count + " unique contracts.");
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "Error walking through MailItems.");
                    }
                    oNameSpace.Logoff();
                    oApp.Quit();

                    //Explicitly release objects.
                    items = null;
                    contractsFolder = null;
                    oNameSpace = null;
                    oApp = null;
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "Error Setting up or leaving Outlook.");
                }

                #endregion

                #region Excel

                Excel.Application eApp = new Excel.Application();
                Excel.Workbook workbook = eApp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                int i = 1;
                foreach (string contract in icList.Distinct().ToList())
                {
                    string tempcontract = contract.Trim(')');

                    string email = "";
                    foreach (char letter in contract.Substring(0, contract.Length - 1).Reverse())
                    {
                        tempcontract = tempcontract.Remove(tempcontract.Length - 1);
                        if (letter == '(')
                        {
                            break;
                        }
                        email = email.Insert(0, letter.ToString());
                    }
                    worksheet.Cells[i, 7] = email;

                    if (tempcontract.Contains("("))
                    {
                        tempcontract = tempcontract.Trim(')', ' ');

                        // label
                        string name = "";
                        foreach (char letter in tempcontract.Reverse())
                        {
                            tempcontract = tempcontract.Remove(tempcontract.Length - 1);
                            if (letter == '(')
                            {
                                break;
                            }
                            name = name.Insert(0, letter.ToString());
                        }
                        worksheet.Cells[i, 1] = name.Trim(' ');

                        // legal
                        worksheet.Cells[i, 4] = tempcontract.Trim(' ');
                    }
                    else if (tempcontract.Contains("/"))
                    {
                        tempcontract = tempcontract.Trim(')', ' ');

                        // label
                        string name = "";
                        foreach (char letter in tempcontract.Reverse())
                        {
                            tempcontract = tempcontract.Remove(tempcontract.Length - 1);
                            if (letter == '/')
                            {
                                break;
                            }
                            name = name.Insert(0, letter.ToString());
                        }
                        worksheet.Cells[i, 1] = name.Trim(' ');

                        // legal
                        worksheet.Cells[i, 4] = tempcontract.Trim(' ');
                    }
                    else
                    {
                        worksheet.Cells[i, 4] = contract.Substring(0, contract.IndexOf("(")).Trim(' ');
                        worksheet.Cells[i, 1] = contract.Substring(0, contract.IndexOf("(")).Trim(' ');
                    }
                    i++;
                }

                eApp.Visible = true;
                eApp.Quit();

                #endregion
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error thrown to higher level to avoid crashing.");
            }
        }

        private void LogBox_Initialized(object sender, EventArgs e)
        {
            Log = new Logger();
        }
    }
}
