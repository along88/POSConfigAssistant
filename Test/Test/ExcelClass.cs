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
    class ExcelClass
    {
        Excel.Application myExcelApplication;
        Excel.Workbook myExcelWorkbook;
        Excel.Worksheet myExcelWorkSheet;
        Excel.Worksheet GoLiveSheet;
        NewList thislist;
        

        public void openExcel()
        {
            myExcelApplication = null;

            myExcelApplication = new Excel.Application(); // create Excell App
            myExcelApplication.DisplayAlerts = false; // turn off alerts
            myExcelWorkbook = myExcelApplication.Workbooks.Open("ONE-CLICK.COAL 6x .POS CONFIG-GO LIVE CHECKLIST. v2.0.xlsx");
            myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[2]; // define in which worksheet, do you want to add data
            GoLiveSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[10];
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
    }
}
