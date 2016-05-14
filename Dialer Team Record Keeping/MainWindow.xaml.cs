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
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private string backendFilePath;

        public MainWindow()
        {
            InitializeComponent();
            InitialStartup();
        }

        public void InitialStartup()
        {
            SetGlobalVariables();
            CheckDailyRecords();
            InitializeDataGrids();
        }

        private void SetGlobalVariables()
        {
            backendFilePath = "C:\\Dialer Team Back-End Database\\DialerTeam_be.accdb";
        }

        public void CheckDailyRecords()
        {
            if (!HasCurrentDayRecords())
            {
                AddDailyRecords();
            }
        }

        public void InitializeDataGrids()
        {
            InitializeInstrumentIdDataGrid();
            InitializeDialerRecordsDataGrid();
        }

        public void InitializeInstrumentIdDataGrid()
        {
            string queryString = "SELECT * FROM InstrumentIDs";
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source=" + backendFilePath + ";" +
                @"User Id=;Password=;";

            FillDataGrid(instrumentIdDataGrid, queryString, connectionString);
        }

        public void InitializeDialerRecordsDataGrid()
        {
            string queryString = "SELECT * FROM DialerRecords";
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source=" + backendFilePath + ";" +
                @"User Id=;Password=;";

            FillDataGrid(dialerRecordsDataGrid, queryString, connectionString);
        }

        public bool HasCurrentDayRecords()
        {
            string query = "SELECT * FROM DialerRecords WHERE DialerDate = #" + DateTime.Today.ToShortDateString() + "#";
            bool hasRecords = false;

            bool recordCheck = (ReadAccessDb(query, false));

            if (recordCheck)
            {
                hasRecords = true;
            }

            return hasRecords;
        }

        public void AddDailyRecords()
        {
            List<string> dailyDialerRecords = new List<string>();

            int currentDay = (int)DateTime.Today.DayOfWeek;

            switch (currentDay)
            {
                //Sunday
                case 0:
                    break;
                //Monday
                case 1:
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','7:15:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','8:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','8:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','10:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','10:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','14:00:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','14:00:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','18:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','19:00:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','18:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','21:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO LSPassRecords(LateStageDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO LMDialerRecords(LossMitDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO LMTimeAndAgents(LossMitDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO ESDialsPerList(CallDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    break;
                //Tuesday
                case 2:
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','7:15:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','8:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','8:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','10:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','10:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','13:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','13:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','18:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','18:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','19:00:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','18:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','21:05:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','21:15:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO LSPassRecords(LateStageDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO LMDialerRecords(LossMitDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO LMTimeAndAgents(LossMitDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO ESDialsPerList(CallDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    break;
                //Wednesday
                case 3:
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','7:15:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','8:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','8:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','10:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','10:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','10:15:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','14:00:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','14:00:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','18:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','18:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','19:00:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','21:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO LSPassRecords(LateStageDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO LMDialerRecords(LossMitDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO LMTimeAndAgents(LossMitDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO ESDialsPerList(CallDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    break;
                //Thursday
                case 4:
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','7:15:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','8:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','8:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','10:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','10:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','10:15:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','14:00:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','10:15:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','18:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','18:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','19:00:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','20:45:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','21:15:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO LSPassRecords(LateStageDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO LMDialerRecords(LossMitDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO LMTimeAndAgents(LossMitDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO ESDialsPerList(CallDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    break;
                //Friday
                case 5:
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','7:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','7:15:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','8:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','8:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','10:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','10:00:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','10:15:00 AM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','14:00:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','14:00:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Early Stage','16:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Loss Mit','19:00:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','18:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO DialerRecords (DialerDate, Department, StartTime) VALUES (#" + DateTime.Today.ToShortDateString() + "#,'Late Stage','21:30:00 PM')");
                    dailyDialerRecords.Add("INSERT INTO LSPassRecords(LateStageDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO LMDialerRecords(LossMitDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO LMTimeAndAgents(LossMitDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    dailyDialerRecords.Add("INSERT INTO ESDialsPerList(CallDate) VALUES(#" + DateTime.Today.ToShortDateString() + "#)");
                    break;
                //Saturday
                case 6:
                    break;


            }

            foreach (string dailyRecord in dailyDialerRecords)
            {
                WriteAccessDb(dailyRecord);
            }


            //WriteAccessDb(query);
        }

        #region Report Methods
        public List<WrapReport> ImportWrap(string fileName)
        {
            List<WrapReport> report = new List<WrapReport>();

            ExcelReaderInterop excelReader = new ExcelReaderInterop();
            report = excelReader.ExcelOpenSpreadsheet(fileName);

            return report;
        }

        public void ExportWrap(List<WrapReport> report)
        {
            OutlookWriterInterop email = new OutlookWriterInterop();
            email.CreateWrapEmail(report);
        }

        public void EmailUpdate()
        {
            OutlookWriterInterop email = new OutlookWriterInterop();
            email.CreateIntradayUpdate(ListProgressTableCreation());
        }

        public string ListProgressTableCreation()
        {
            double currentGoal = CalculateCurrentGoal();
            string rowColor = "<td bgcolor''baa777'>";

            string table = "<TABLE BORDER='4' CELLSPACING='0' CELLPADDING='3' style='border-collapse: collapse; text-align:center; width: 225px; bordercolor='#111111'>" +
                           "<TR HEIGHT=19><TD colspan='2' BGCOLOR='#4d4d4d'><FONT COLOR='FFFFFF'>List Progress Summary</TD></TR><tr><FONT COLOR='000000'>" +
                           rowColor + "Intensity</td>" +
                           rowColor + combinedSaturation.ToString() + "</td></FONT></tr>";

            table += "<tr><FONT COLOR='000000'>" +
                rowColor + "PID Goal</td>" +
                rowColor + (currentGoal * 100).ToString() + "%</td>" +
                "</FONT></tr>";

            table += "<tr><FONT COLOR='000000'>" +
                rowColor + "Pass Progression I</td>" +
                rowColor + (double.Parse(combinedPassPercent.ToString()) * 100) + "%</td>" +
                "</FONT></tr>";

            //table += "<tr><FONT COLOR='000000'>" +
            //    rowColor + "Pass Progression II</td>" +
            //    rowColor + (double.Parse(combinedPercentWorkedTwo.ToString()) * 100) + "%</td>" +
            //    "</FONT></tr>";

            //table += "<tr><FONT COLOR='000000'>" +
            //    rowColor + "Pass Progression III</td>" +
            //    rowColor + (double.Parse(combinedPercentWorkedThree.ToString()) * 100) + "%</td>" +
            //    "</FONT></tr></table><br>";

            return table;
        }
        #endregion

        #region Database Methods
        public void FillDataGrid(DataGrid dataGridName, string queryString, string connectionString)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                using (OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection))
                {
                    System.Data.DataTable table = new System.Data.DataTable();
                    adapter.Fill(table);

                    if (dataGridName.Name == "dialerRecordsDataGrid")
                    {
                        foreach (DataRow row in table.Rows)
                        {
                            if (row["DialerDate"].ToString() != DateTime.Today.ToString())
                            {
                                row.Delete();
                            }
                        }

                        table.Columns.Remove("ID");
                        table.Columns.Remove("Target");
                        table.Columns.Remove("BusinessCenter");
                        table.Columns.Remove("DialerDate");
                    }

                    dataGridName.ItemsSource = table.DefaultView;


                }
            }
        }

        public void FillDataGrid(DataGrid dataGridName, string queryString, string connectionString, DateTime filter)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                using (OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection))
                {
                    System.Data.DataTable table = new System.Data.DataTable();
                    adapter.Fill(table);

                    if (dataGridName.Name == "dialerRecordsDataGrid")
                    {
                        foreach (DataRow row in table.Rows)
                        {
                            if (row["DialerDate"].ToString() != filter.ToString())
                            {
                                row.Delete();
                            }
                        }

                        //table.Columns.Remove
                    }

                    dataGridName.ItemsSource = table.DefaultView;


                }
            }
        }

        public void InsertData(string tableName)
        {

        }

        /// <summary>
        /// Creates an array used to lookup Early Stage penetration goal based on current hour and current day of the week.
        /// </summary>
        /// <returns>Returns the projected penetration goal for the current day and hour.</returns>
        public double CalculateCurrentGoal()
        {
            int currentDay = (int)DateTime.Now.DayOfWeek;
            int currentHour = DateTime.Now.Hour;

            double goal = 0.0;
            double[,] goals = new double[5, 24];

            try
            {
                for (int i = 0; i < 5; i++)
                {
                    for (int j = 0; j < 8; j++)
                    {
                        goals[i, j] = 0;
                    }
                }

                goals[0, 8] = 0.14;
                goals[0, 9] = 0.19;
                goals[0, 10] = 0.23;
                goals[0, 11] = 0.37;
                goals[0, 12] = 0.42;
                goals[0, 13] = 0.45;
                goals[0, 14] = 0.54;
                goals[0, 15] = 0.67;
                goals[0, 16] = 0.80;
                goals[0, 17] = 0.85;
                goals[0, 18] = 0.88;
                goals[0, 19] = 0.90;
                goals[0, 20] = 0.93;
                goals[0, 21] = 0.97;
                goals[0, 22] = 1.00;
                goals[0, 23] = 0.0;

                goals[1, 8] = 0.11;
                goals[1, 9] = 0.19;
                goals[1, 10] = 0.23;
                goals[1, 11] = 0.29;
                goals[1, 12] = 0.35;
                goals[1, 13] = 0.50;
                goals[1, 14] = 0.69;
                goals[1, 15] = 0.80;
                goals[1, 16] = 0.87;
                goals[1, 17] = 0.90;
                goals[1, 18] = 0.92;
                goals[1, 19] = 0.94;
                goals[1, 20] = 0.97;
                goals[1, 21] = 0.99;
                goals[1, 22] = 1.00;
                goals[1, 23] = 0.0;

                goals[2, 8] = 0.15;
                goals[2, 9] = 0.27;
                goals[2, 10] = 0.37;
                goals[2, 11] = 0.51;
                goals[2, 12] = 0.61;
                goals[2, 13] = 0.67;
                goals[2, 14] = 0.74;
                goals[2, 15] = 0.81;
                goals[2, 16] = 0.89;
                goals[2, 17] = 0.92;
                goals[2, 18] = 0.95;
                goals[2, 19] = 0.96;
                goals[2, 20] = 0.98;
                goals[2, 21] = 0.99;
                goals[2, 22] = 1.00;
                goals[2, 23] = 0.0;

                goals[3, 8] = 0.7;
                goals[3, 9] = 0.13;
                goals[3, 10] = 0.16;
                goals[3, 11] = 0.21;
                goals[3, 12] = 0.28;
                goals[3, 13] = 0.43;
                goals[3, 14] = 0.56;
                goals[3, 15] = 0.74;
                goals[3, 16] = 0.89;
                goals[3, 17] = 0.93;
                goals[3, 18] = 0.95;
                goals[3, 19] = 0.96;
                goals[3, 20] = 0.98;
                goals[3, 21] = 0.99;
                goals[3, 22] = 1.00;
                goals[3, 23] = 0.0;

                goals[4, 8] = 0.9;
                goals[4, 9] = 0.18;
                goals[4, 10] = 0.22;
                goals[4, 11] = 0.30;
                goals[4, 12] = 0.36;
                goals[4, 13] = 0.43;
                goals[4, 14] = 0.53;
                goals[4, 15] = 0.64;
                goals[4, 16] = 0.77;
                goals[4, 17] = 0.83;
                goals[4, 18] = 0.85;
                goals[4, 19] = 0.89;
                goals[4, 20] = 0.95;
                goals[4, 21] = 0.98;
                goals[4, 22] = 1.00;
                goals[4, 23] = 0.0;

                goal = goals[currentDay - 1, currentHour];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return goal;
        }
        #endregion
        
        public void ImportDataFromExcel()
        {
            ExcelReaderInterop interop = new ExcelReaderInterop();
            List<ListProgressData> esImport = new List<ListProgressData>();

            esImport = interop.ExcelReadSpreadsheet("C:\\Dialer Team Back-End Database\\List Progress - All Departments.xlsm");

            massCellRecords.Text = esImport[0].Records.ToString();
            massCellDials.Text = esImport[0].CallsDialed.ToString();
            massCellPercentWorked.Text = esImport[0].Penetration.ToString();

            miCellRecords.Text = esImport[1].Records.ToString();
            miCellDials.Text = esImport[1].CallsDialed.ToString();
            miCellPercentWorked.Text = esImport[1].Penetration.ToString();

            massRecords.Text = esImport[2].Records.ToString();
            massDials.Text = esImport[2].CallsDialed.ToString();
            massPercentWorked.Text = esImport[2].Penetration.ToString();

            nhRecords.Text = esImport[3].Records.ToString();
            nhDials.Text = esImport[3].CallsDialed.ToString();
            nhPercentWorked.Text = esImport[3].Penetration.ToString();

            nonContRecords.Text = esImport[4].Records.ToString();
            nonContDials.Text = esImport[4].CallsDialed.ToString();
            nonContPercentWorked.Text = esImport[4].Penetration.ToString();

            orRecords.Text = esImport[5].Records.ToString();
            orDials.Text = esImport[5].CallsDialed.ToString();
            orPercentWorked.Text = esImport[5].Penetration.ToString();

            //Inteded for combined, probably not needed
            allRecords.Text = esImport[6].Records.ToString();
            //miCellDials.Text = esImport[0].CallsDialed.ToString();

            //If Pass Two is complete, enter data into Pass Three
            if ((allPercentWorkedThree.Text != "0") || (allPercentWorkedThree.Text == "0" && esImport[6].Penetration < double.Parse(allPercentWorkedTwo.Text)))
            {
                allPercentWorkedThree.Text = esImport[6].Penetration.ToString();
                allDialsThree.Text = (esImport[6].CallsDialed - (double.Parse(allDialsTwo.Text) + double.Parse(allDialsOne.Text))).ToString();
            }
            //If Pass One is complete, enter data into Pass Two
            else if ((allPercentWorkedTwo.Text != "0") || (allPercentWorkedTwo.Text == "0" && esImport[6].Penetration < double.Parse(allPercentWorkedOne.Text)))
            {
                allPercentWorkedTwo.Text = esImport[6].Penetration.ToString();
                allDialsTwo.Text = (esImport[6].CallsDialed - double.Parse(allDialsOne.Text)).ToString();
            }
            //Enter data into Pass One
            else
            {
                allPercentWorkedOne.Text = esImport[6].Penetration.ToString();
                allDialsOne.Text = esImport[6].CallsDialed.ToString();
            }




            #region Set up and calculate List Percent Total boxes
            
            List<TextBox> percentTotalBoxes = new List<TextBox>();            
            List<double> recordBoxes = new List<double>();

            //Get the set of Percent Total boxes
            percentTotalBoxes = CreatePercentVolumeList();

            //Get the values of the Record boxes
            recordBoxes = CreateRecordsList();

            //Iterate through the List of Percent Total boxes and set Percent Total value in them
            for (int i = 0; i < percentTotalBoxes.Count; i++)
            {
                double pctVol = CalculatePercentVolumeRaw(recordBoxes[i], double.Parse(combinedRecords.Text));
                percentTotalBoxes[i].Text = Math.Round(pctVol * 100, 2).ToString("0.00");
            }
            #endregion

            #region Set up and calculate List Saturation boxes
            List<double> recordBoxesSaturation = new List<double>();
            List<int> dialsBoxes = new List<int>();
            List<TextBox> saturationBoxes = new List<TextBox>();
            recordBoxesSaturation = CreateRecordsList(true);
            dialsBoxes = CreateDialsList();
            saturationBoxes = CreateSaturationList();

            for (int i = 0; i < saturationBoxes.Count; i++)
            {
                saturationBoxes[i].Text = (Math.Round((CalculateSaturation(recordBoxesSaturation[i], (double)dialsBoxes[i])) * 100)).ToString();
            }
            #endregion

            #region Set up and calculate Percent Worked Boxes
            List<TextBox> passPercentBoxes = new List<TextBox>();
            List<double> percentVolumeRaw = new List<double>();
            List<double> passPenetrationBoxes = new List<double>();
            List<TextBox> percentWorkedBoxes = new List<TextBox>();

            passPercentBoxes = CreatePassPenetrationList();
            percentWorkedBoxes = CreatePercentWorkedList();

            for (int i = 0; i < passPercentBoxes.Count; i++)
            {
                percentVolumeRaw.Add(CalculatePercentVolumeRaw(recordBoxesSaturation[i], double.Parse(combinedRecords.Text)));
                passPenetrationBoxes.Add(double.Parse(percentWorkedBoxes[i].Text));
            }

            for (int i = 0; i < passPercentBoxes.Count; i++)
            {
                DisplayPassPercent(percentVolumeRaw[i], passPenetrationBoxes[i], passPercentBoxes[i]);
            }

            #endregion
        }

        public List<double> CalculatePercentVolumeValues(List<TextBox> boxes)
        {
            List<double> results = new List<double>();            
            List<double> recordBoxes = new List<double>();

            //Get the values of the Record boxes
            recordBoxes = CreateRecordsList();

            //Iterate through the List of Percent Total boxes and set Percent Total value in them
            for (int i = 0; i < boxes.Count; i++)
            {
                results.Add(Math.Round(CalculatePercentVolumeRaw(recordBoxes[i], double.Parse(combinedRecords.Text)) * 100,2));
            }

            return results;
        }







        public void CreateEarlyStageReadQuery(DateTime recordDate)
        {
            if (recordDate != DateTime.Today)
            {
                recordDate = DateTime.Parse(recordDateES.ToString());
            }

            recordDateES.Text = recordDate.ToShortDateString();

            string queryString = "SELECT * FROM ESDialsPerList WHERE CallDate = #" + recordDate.ToString() + "#";

            ReadEarlyStageData(queryString);
        }

        




        

        public class ListProgressImport
        {
            List<ListProgressData> individualLists;
        }
    }    
}
