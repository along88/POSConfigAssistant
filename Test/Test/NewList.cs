using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Test
{
    public partial class NewList : Form
    {
        public string ExcelFilePath;
        public int Rownumber;
        int Increase = 100;
        CheckedListBox[] Registers;
        int[] cells = { 14,32,50,62,65,72,
                          79,95,109,116,125,
                          137,142,148,161,174,
                          200,206,224,228,330,
                          241,272,318,295,353,
                          360,372,387,393,400,506};

        public NewList()
        {
            InitializeComponent();
        }

        private void NewListNextButton_Click(object sender, EventArgs e)
        {
            //ExcelFilePath = "ONE-CLICK.COAL 6x .POS CONFIG-GO LIVE CHECKLIST. v2.0.xlsx";
            //openExcel();
            //if (NewListConfigorGoLivecomboBox.Text.ToUpper() == "GO LIVE")
            //{
            //    addDataToExcel(NewListClientNametextBox.Text,
            //          NewListStoreNumberTextBox.Text,
            //          NewListPhoneNumbertextBox.Text,
            //          NewListStoreTypecomboBox.Text,
            //          NewListInstallTickettextBox.Text,
            //          NewListInstallercomboBox.Text,
            //          NewListSalesTaxtextBox.Text,
            //          NewListTypecomboBox.Text,
            //          NewListCoalitoncomboBox.Text,
            //          NewListBoardTypecomboBox.Text,
            //          NewListYourNameTextBox.Text);

            //    closeExcel();
            //    NextPanel.Visible = true;
            //}
            NextPanel.Visible = true;
        }
        Excel.Application myExcelApplication;
        Excel.Workbook myExcelWorkbook;
        Excel.Worksheet myExcelWorkSheet;
        //differeniate between go live or a config
        Excel.Worksheet GoLiveSheet;
        Excel.Worksheet ConfigSheet;
        ExcelClass myClass = new ExcelClass();

        
        public void openExcel()
        {
            myExcelApplication = null;

            myExcelApplication = new Excel.Application(); // create Excell App
            myExcelApplication.DisplayAlerts = false; // turn off alerts
            myExcelWorkbook = myExcelApplication.Workbooks.Open("ONE-CLICK.COAL 6x .POS CONFIG-GO LIVE CHECKLIST. v2.0.xlsx");
            myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[2]; // define in which worksheet, do you want to add data
            GoLiveSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[10];
        }

        public void ExcelCheckList(string name)
        {
            if (NewListConfigorGoLivecomboBox.Text.ToUpper() == "GO LIVE")
            {


                for (int i = 0; i < Registers[0].CheckedItems.Count; i++)
                {
                    if (Registers[0].Items[i] == Registers[0].CheckedItems[i])
                    {
                        
                        
                        //if (GoLiveSheet.Cells[x, "A"].Value % 1 == 0)
                            GoLiveSheet.Cells[cells[i], "C"] = name;
                        
                    }

                }

            }
        }
        public void addDataToExcel(string client,
            string storeNumber,
            string phoneNumber,
            string storetype,
            string ticket,
            string installer,
            string tax,
            string installType,
            string coal,
            string board,
            string name)
        {

            myExcelWorkSheet.Cells[7, "B"] = client;
            myExcelWorkSheet.Cells[8, "B"] = storeNumber;
            myExcelWorkSheet.Cells[9, "B"] = phoneNumber;
            myExcelWorkSheet.Cells[10, "B"] = storetype;
            myExcelWorkSheet.Cells[13, "B"] = ticket;
            myExcelWorkSheet.Cells[15, "B"] = installer;
            myExcelWorkSheet.Cells[19, "B"] = tax;
            myExcelWorkSheet.Cells[14, "B"] = installType;
            myExcelWorkSheet.Cells[33, "B"] = coal;
            myExcelWorkSheet.Cells[32, "B"] = board;
            myExcelWorkSheet.Cells[4, "B"] = name;
            myExcelWorkSheet.Cells[3, "B"] = DateTime.Now.ToString();
            myExcelWorkSheet.Cells[16, "B"] = "YES";
            myExcelWorkSheet.Cells[17, "B"] = "YES";
            myExcelWorkSheet.Cells[18, "B"] = "YES";
            myExcelWorkSheet.Cells[11, "B"] = NewListRegistertextBox.Text;
            myExcelWorkSheet.Cells[12, "B"] = NewListLineIPtextBox.Text;
            myExcelWorkSheet.Cells[20, "B"] = NewListPublicIPtextBox1.Text;
            myExcelWorkSheet.Cells[21, "B"] = NewListPublicIPtextBox2.Text;
            myExcelWorkSheet.Cells[22, "B"] = NewListInternalIPtextBox1.Text;
            myExcelWorkSheet.Cells[23, "B"] = NewListInternalIPtextBox2.Text;
            myExcelWorkSheet.Cells[24, "B"] = NewListSubnettextBox.Text;
            myExcelWorkSheet.Cells[25, "B"] = NewListGatewaytextBox.Text;
            myExcelWorkSheet.Cells[26, "B"] = NewListDnstextBox1.Text;
            myExcelWorkSheet.Cells[27, "B"] = NewListDnstextBox2.Text;
            myExcelWorkSheet.Cells[28, "B"] = NewListPrinterIPtextBox.Text;
            myExcelWorkSheet.Cells[29, "B"] = NewListRegKeytextBox1.Text;
            myExcelWorkSheet.Cells[30, "B"] = NewListRegKeytextBox2.Text;
            myExcelWorkSheet.Cells[32, "B"] = NewListBoardTypecomboBox.Text;
            myExcelWorkSheet.Cells[33, "B"] = NewListCoalitoncomboBox.Text;
            myExcelWorkSheet.Cells[34, "B"] = NewListOScomboBox.Text;
            myExcelWorkSheet.Cells[35, "B"] = NewListConnectivitycomboBox.Text;
            myExcelWorkSheet.Cells[36, "B"] = NewListMerchantcomboBox.Text;
            myExcelWorkSheet.Cells[37, "B"] = NewListCreditcomboBox.Text;
            myExcelWorkSheet.Cells[38, "B"] = NewListPinpadscomboBox.Text;
            myExcelWorkSheet.Cells[39, "B"] = NewListGiftcomboBox.Text;
            myExcelWorkSheet.Cells[40, "B"] = NewListCheckscomboBox.Text;
            myExcelWorkSheet.Cells[41, "B"] = NewListCustDisplaycomboBox.Text;
            myExcelWorkSheet.Cells[42, "B"] = NewListPrintercomboBox.Text;
            myExcelWorkSheet.Cells[43, "B"] = NewListScannerModelcomboBox.Text;
            myExcelWorkSheet.Cells[44, "B"] = NewListSymbolcomboBox.Text;
            myExcelWorkSheet.Cells[45, "B"] = NewListMonarchcomboBox.Text;
            myExcelWorkSheet.Cells[46, "B"] = NewListNetworkPrintercomboBox.Text;
            myExcelWorkSheet.Cells[47, "B"] = NewListMposcomboBox.Text;

            if (coal == "COAL6")
            {
                myExcelWorkSheet.Cells[50, "B"] = "T:\\POS CONFIG TEAM\\Installation Documents and Electronic Tracking-12-23-10\\Coalition 6\\Installation by Client-Store\\" + client + "\\" + storeNumber;
            }
            else if (coal == "COAL5")
            {
                myExcelWorkSheet.Cells[50, "B"] = "T:\\POS CONFIG TEAM\\Installation Documents and Electronic Tracking-12-23-10\\Coalition 5\\Installation by Client-Store\\" + client + "\\" + storeNumber;
            }


        }

        public void closeExcel()
        {
            try
            {


                myExcelWorkbook.SaveCopyAs("\\NewExcel.xlsx");
                myExcelWorkbook.Close(false, Type.Missing, Type.Missing);
                myExcelApplication.Application.Quit();
                myExcelApplication.Quit();

            }
            finally
            {
                if (myExcelApplication != null)
                {
                    myExcelApplication.Quit(); // close the excel application
                }
            }

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            NextPanel2.Visible = false;

        }

        private void NewListNextbutton2_Click(object sender, EventArgs e)
        {
            NextPanel2.Visible = true;


        }

        private void button2_Click(object sender, EventArgs e)
        {
            NextPanel.Visible = false;

        }

        private void button5_Click(object sender, EventArgs e)
        {
            Panel ChecklistPanel = new Panel();
            Button SubmitList = new Button();
            SubmitList.Parent = ChecklistPanel;
            SubmitList.Text = "Submit Check List";
            SubmitList.Dock = DockStyle.Right;
            SubmitList.Click += new EventHandler(SubmitList_Click);

            //SubmitList.Height = 184;
            //SubmitList.Anchor = AnchorStyles.Right;
            //SubmitList.Anchor = AnchorStyles.Bottom;
            ChecklistPanel.Parent = NextPanel;
            ChecklistPanel.BringToFront();
            NextPanel2.Visible = false;
            ChecklistPanel.Dock = DockStyle.Fill;
            ChecklistPanel.Enabled = true;
            ChecklistPanel.Visible = true;



            string[] items = { "Assign Ticket","POS Printer", "Micr reader", "datetime", "pinpad", "check RC", "Customer#", "fipay store#", "Update Nav",
                                 "update next trans", "keyboard", "check master", "Incrementals", "Symbol3000", "report printer", "replication",
                                 "scanner", "touchscreen", "gcfix", "Testing", "Cash", "Credit", "debit", "check", "Gift Card", "clean up", "EJ BackUp", 
                                 "Set Live","remove Fipay", "archive IP", "Updates/post installs", "Polling", };

            Registers = new CheckedListBox[int.Parse(NewListRegistertextBox.Text)];


            for (int i = 0; i < int.Parse(NewListRegistertextBox.Text); i++)
            {
                if (Registers[i] == Registers[0])
                {
                    Registers[i] = new CheckedListBox();
                    Registers[i].Parent = ChecklistPanel;
                    Registers[i].IntegralHeight = true;
                    Registers[i].Dock = DockStyle.Left;
                    for (int x = 0; x < items.Length; x++)
                    {
                        Registers[i].Items.Add(items[x].ToString());

                    }
                    Registers[i].Name = "Controller";
                    ChecklistPanel.Controls.Add(Registers[i]);
                    Registers[i].Visible = true;
                }
                else if (Registers[i] != Registers[0])
                {
                    Registers[i] = new CheckedListBox();
                    Registers[i].Name = "Primary";
                    for (int x = 0; x < items.Length; x++)
                    {
                        Registers[i].Items.Add(items[x].ToString());
                    }
                    Registers[i].FormattingEnabled = true;
                    Registers[i].IntegralHeight = true;
                    Registers[i].Dock = DockStyle.Left;
                    Registers[i].Location = new Point(Increase + Registers[0].Location.X);
                    ChecklistPanel.Controls.Add(Registers[i]);
                    Registers[i].Visible = true;
                    Increase++;
                    if (i > 1)
                    {
                        NewList.ActiveForm.Width += 100;
                    }


                }
            }
        }

        private void SubmitList_Click(object sender, EventArgs e)
        {
            ExcelFilePath = "ONE-CLICK.COAL 6x .POS CONFIG-GO LIVE CHECKLIST. v2.0.xlsx";
            //openExcel();
            myClass.openExcel();
            if (NewListConfigorGoLivecomboBox.Text.ToUpper() == "GO LIVE")
            {
                addDataToExcel(NewListClientNametextBox.Text,
                      NewListStoreNumberTextBox.Text,
                      NewListPhoneNumbertextBox.Text,
                      NewListStoreTypecomboBox.Text,
                      NewListInstallTickettextBox.Text,
                      NewListInstallercomboBox.Text,
                      NewListSalesTaxtextBox.Text,
                      NewListTypecomboBox.Text,
                      NewListCoalitoncomboBox.Text,
                      NewListBoardTypecomboBox.Text,
                      NewListYourNameTextBox.Text);

                ExcelCheckList(NewListYourNameTextBox.Text);
                //closeExcel();
                myClass.closeExcel();
            }
        }

        private void CheckListPanel_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
