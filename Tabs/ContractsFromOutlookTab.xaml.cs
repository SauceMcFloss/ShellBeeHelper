using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Globalization;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using ShellBeeHelper.Properties;

namespace ShellBeeHelper.Tabs
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class ContractsFromOutlookTab : UserControl
    {
        #region Constructor

        public ContractsFromOutlookTab(Logger log)
        {
            InitializeComponent();

            Log = log;

            EmailAddressTextBox.Text = Settings.Default.EmailAddress;
            SourceFolderTextBox.Text = Settings.Default.SourceFolder;
            DestFolderTextBox.Text = Settings.Default.DestFolder;
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

        private void ScanButton_Click(object sender, RoutedEventArgs e)
        {
            #region Validation

            if (!IsValidEmail(EmailAddressTextBox.Text))
            {
                Log.Log("No valid email address entered.");
                EmailAddressTextBox.Background = System.Windows.Media.Brushes.Red;
                return;
            }
            if (String.IsNullOrWhiteSpace(SourceFolderTextBox.Text))
            {
                Log.Log("No source folder name entered.");
                SourceFolderTextBox.Background = System.Windows.Media.Brushes.Red;
                return;
            }
            if (String.IsNullOrWhiteSpace(DestFolderTextBox.Text))
            {
                Log.Log("No destination folder name entered.");
                DestFolderTextBox.Background = System.Windows.Media.Brushes.Red;
                return;
            }

            #endregion

            try
            {
                List<string> icList = new List<string>();

                #region Outlook

                Outlook.Application oApp = null;
                Outlook.NameSpace oNameSpace = null;
                Outlook.MAPIFolder contractsSourceFolder = null;
                Outlook.MAPIFolder contractsDestFolder = null;
                Outlook.Items items = null;
                List<Outlook.MailItem> messages = null;

                try
                {
                    oApp = new Outlook.Application();
                    oNameSpace = oApp.GetNamespace("mapi");
                    oNameSpace.Logon(Missing.Value, Missing.Value, false, true);

                    contractsSourceFolder = oNameSpace.Folders[EmailAddressTextBox.Text].Folders[SourceFolderTextBox.Text];
                    contractsDestFolder = oNameSpace.Folders[EmailAddressTextBox.Text].Folders[DestFolderTextBox.Text];
                    items = contractsSourceFolder.Items;
                    messages = new List<Outlook.MailItem>();

                    try
                    {
                        foreach(Outlook.MailItem msg in items)
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
                                    Log.Log("Encountered unknown email format.\n\tFrom: " + msg.SenderName + "\n\tSubject: " + msg.Subject + "\n\tReceived at: " + msg.ReceivedTime);
                                    continue;
                                }

                                importantContent = importantContent.Trim(' ', '\r', '\n');
                                icList.Add(importantContent);
                                messages.Add(msg);

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
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "Error Setting up or leaving Outlook.");
                    return;
                }

                #endregion

                #region Excel

                Excel.Application eApp = null;
                Excel.Workbook workbook = null;
                Excel.Worksheet worksheet = null;

                try
                {
                    eApp = new Excel.Application();
                    workbook = eApp.Workbooks.Add();
                    worksheet = workbook.ActiveSheet;

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
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "Error Setting up or leaving Excel.");
                    return;
                }

                eApp.Visible = true;

                #endregion

                #region Move emails

                foreach(Outlook.MailItem msg in messages)
                {
                    msg.Move(contractsDestFolder);
                    msg.UnRead = false;
                }

                #endregion

                #region Global

                // Release Outlook
                try
                {
                    oNameSpace.Logoff();
                    oApp.Quit();

                    messages = null;
                    items = null;
                    contractsDestFolder = null;
                    contractsSourceFolder = null;
                    oNameSpace = null;
                    oApp = null;
                }
                catch
                {

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

                Settings.Default.EmailAddress = EmailAddressTextBox.Text;
                Settings.Default.SourceFolder = SourceFolderTextBox.Text;
                Settings.Default.DestFolder = DestFolderTextBox.Text;
                Settings.Default.Save();

                #endregion
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error thrown to higher level to avoid crashing.");
            }
        }

        private void EmailAddressTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (IsValidEmail(EmailAddressTextBox.Text))
            {
                EmailAddressTextBox.Background = System.Windows.Media.Brushes.White;
            }
            else
            {
                EmailAddressTextBox.Background = System.Windows.Media.Brushes.Red;
            }
        }

        #endregion

        #region Methods and Functions

        public static bool IsValidEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;

            try
            {
                // Normalize the domain
                email = Regex.Replace(email, @"(@)(.+)$", DomainMapper,
                                      RegexOptions.None, TimeSpan.FromMilliseconds(200));

                // Examines the domain part of the email and normalizes it.
                string DomainMapper(Match match)
                {
                    // Use IdnMapping class to convert Unicode domain names.
                    var idn = new IdnMapping();

                    // Pull out and process domain name (throws ArgumentException on invalid)
                    string domainName = idn.GetAscii(match.Groups[2].Value);

                    return match.Groups[1].Value + domainName;
                }
            }
            catch (RegexMatchTimeoutException e)
            {
                return false;
            }
            catch (ArgumentException e)
            {
                return false;
            }

            try
            {
                return Regex.IsMatch(email,
                    @"^[^@\s]+@[^@\s]+\.[^@\s]+$",
                    RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
        }

        #endregion
    }
}
