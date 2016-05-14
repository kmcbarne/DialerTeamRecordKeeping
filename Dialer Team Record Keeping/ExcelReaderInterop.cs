using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Runtime.InteropServices;
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
using TextBox = System.Windows.Controls.TextBox;
using Exception = System.Exception;
using System.Collections;

namespace Dialer_Team_Record_Keeping
{
    class ExcelReaderInterop
    {
        Microsoft.Office.Interop.Excel.Application excelApp;

        public ExcelReaderInterop()
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
        }

        public List<WrapReport> ExcelOpenSpreadsheet(string fileName)
        {
            List<WrapReport> report = new List<WrapReport>();
            try
            {
                //Opens Excel workbook and assigns it to Workbook book
                Workbook book = excelApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //Scans worksheets in Workbook book
                report = ExcelScanWorkbook(book);

                //Closes Workbook book and purges COM object from memory
                book.Close(false, fileName);
                Marshal.ReleaseComObject(book);

                return report;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
        }

        private List<WrapReport> ExcelScanWorkbook(Workbook book)
        {
            List<WrapReport> report = new List<WrapReport>();

            //Gets sheet count and stores the number of sheets
            int sheetCount = book.Sheets.Count;

            //Iterate through sheets in Workbook book.  Index begins at 1 instead of 0.
            for (int sheetNumber = 1; sheetNumber < sheetCount + 1; sheetNumber++)
            {
                Worksheet sheet = (Worksheet)book.Sheets[sheetNumber];

                //Takes the used range of the sheet.  Gets object array of all cells in the sheet by value.
                Range excelRange = sheet.UsedRange;
                object[,] valueArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                //Process data in array
                report = ProcessObjects(valueArray);
            }

            return report;
        }

        public List<ListProgressData> ExcelReadSpreadsheet(string fileName)
        {
            List<ListProgressData> import = new List<ListProgressData>();
            try
            {
                //Opens Excel workbook and assigns it to Workbook book
                Workbook book = excelApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //Scans worksheets in Workbook book
                import = ImportEarlyStageData(book);

                //Closes Workbook book and purges COM object from memory
                book.Close(false, fileName);
                Marshal.ReleaseComObject(book);

                return import;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
        }

        public List<ListProgressData> ImportEarlyStageData(Workbook book)
        {
            List<Worksheet> sheets = new List<Worksheet>();
            List<ListProgressData> allLists = new List<ListProgressData>();

            int sheetCount = book.Sheets.Count;

            for (int sheetNumber = 1; sheetNumber < sheetCount + 1; sheetNumber++)
            {
                Worksheet sheet = (Worksheet)book.Sheets[sheetNumber];

                ListProgressData massCell = new ListProgressData();
                ListProgressData miCell = new ListProgressData();
                ListProgressData mass = new ListProgressData();
                ListProgressData nh = new ListProgressData();
                ListProgressData nonCont = new ListProgressData();
                ListProgressData or = new ListProgressData();
                ListProgressData all = new ListProgressData();

                Range excelRange = sheet.UsedRange;
                object[,] valueArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
                switch (sheetNumber)
                {
                    case 1:
                        break;
                    case 2:
                        break;
                    case 3:
                        massCell.CallsDialed = sheet.Range["D3"].Value2;
                        massCell.TransferredToAgent = sheet.Range["D4"].Value2;
                        massCell.PTP = sheet.Range["D5"].Value2;
                        massCell.RPC = sheet.Range["D6"].Value2;
                        allLists.Add(massCell);
                        miCell.CallsDialed = sheet.Range["E3"].Value2;
                        miCell.TransferredToAgent = sheet.Range["E4"].Value2;
                        miCell.PTP = sheet.Range["E5"].Value2;
                        miCell.RPC = sheet.Range["E6"].Value2;
                        allLists.Add(miCell);
                        mass.CallsDialed = sheet.Range["F3"].Value2;
                        mass.TransferredToAgent = sheet.Range["F4"].Value2;
                        mass.PTP = sheet.Range["F5"].Value2;
                        mass.RPC = sheet.Range["F6"].Value2;
                        allLists.Add(mass);
                        nh.CallsDialed = sheet.Range["G3"].Value2;
                        nh.TransferredToAgent = sheet.Range["G4"].Value2;
                        nh.PTP = sheet.Range["G5"].Value2;
                        nh.RPC = sheet.Range["G6"].Value2;
                        allLists.Add(nh);
                        nonCont.CallsDialed = sheet.Range["H3"].Value2;
                        nonCont.TransferredToAgent = sheet.Range["H4"].Value2;
                        nonCont.PTP = sheet.Range["H5"].Value2;
                        nonCont.RPC = sheet.Range["H6"].Value2;
                        allLists.Add(nonCont);
                        or.CallsDialed = sheet.Range["I3"].Value2;
                        or.TransferredToAgent = sheet.Range["I4"].Value2;
                        or.PTP = sheet.Range["I5"].Value2;
                        or.RPC = sheet.Range["I6"].Value2;
                        allLists.Add(or);
                        all.CallsDialed = sheet.Range["J3"].Value2;
                        all.TransferredToAgent = sheet.Range["J4"].Value2;
                        all.PTP = sheet.Range["J5"].Value2;
                        all.RPC = sheet.Range["J6"].Value2;
                        allLists.Add(all);
                        break;
                    case 4:
                        break;
                    case 5:
                        allLists[0].Records = sheet.Range["B5"].Value2;
                        allLists[0].Penetration = sheet.Range["C5"].Value2;
                        allLists[1].Records = sheet.Range["B6"].Value2;
                        allLists[1].Penetration = sheet.Range["C6"].Value2;
                        allLists[2].Records = sheet.Range["B7"].Value2;
                        allLists[2].Penetration = sheet.Range["C7"].Value2;
                        allLists[3].Records = sheet.Range["B8"].Value2;
                        allLists[3].Penetration = sheet.Range["C8"].Value2;
                        allLists[4].Records = sheet.Range["B9"].Value2;
                        allLists[4].Penetration = sheet.Range["C9"].Value2;
                        allLists[5].Records = sheet.Range["B10"].Value2;
                        allLists[5].Penetration = sheet.Range["C10"].Value2;
                        allLists[6].Records = sheet.Range["B11"].Value2;
                        allLists[6].Penetration = sheet.Range["C11"].Value2;
                        break;
                }
            }
            return allLists;
        }

        private List<WrapReport> ProcessObjects(object[,] valueArray)
        {
            List<WrapReport> report = new List<WrapReport>();

            //valueArray.Length avoid exception being thrown from final row in source sheet having a null value in column 1			
            for (int i = 1; i < valueArray.GetLength(0); i++)
            {
                WrapReport agentWrapReport = new WrapReport();

                //Create string identifier to check whether line contains agent data
                string identifier = ((string)valueArray[i, 1]).Substring(0, 1);

                //Checks to see if row contains agent data
                if (identifier == "T")
                {
                    agentWrapReport.Tid = ((string)valueArray[i, 1]).Substring(0, 7);
                    agentWrapReport.Name = ((string)valueArray[i, 1]).Substring(10);
                    agentWrapReport.Calls = Convert.ToInt32(valueArray[i, 2]);
                    agentWrapReport.Login = TimeSpan.Parse((string)valueArray[i, 3]);
                    agentWrapReport.Active = TimeSpan.Parse((string)valueArray[i, 4]);
                    agentWrapReport.NotReady = TimeSpan.Parse((string)valueArray[i, 5]);
                    agentWrapReport.Idle = TimeSpan.Parse((string)valueArray[i, 6]);
                    agentWrapReport.Wrap = TimeSpan.Parse((string)valueArray[i, 7]);
                    agentWrapReport.Hold = TimeSpan.Parse((string)valueArray[i, 8]);
                    agentWrapReport.HoldCount = Convert.ToInt32(valueArray[i, 9]);
                    agentWrapReport.AvgWrap = agentWrapReport.CalculateAverageWrap(agentWrapReport.Wrap, agentWrapReport.Calls);

                    report.Add(agentWrapReport);
                }

            }

            //Reorders the list in descending order by TimeSpan AvgWrap
            report.Sort(delegate (WrapReport x, WrapReport y)
            {
                return y.AvgWrap.CompareTo(x.AvgWrap);
            });

            return report;
        }
    }
}
