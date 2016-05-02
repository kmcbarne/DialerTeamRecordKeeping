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

namespace Dialer_Team_Record_Keeping
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            
        }

        /// <summary>
        /// Opens Access database filePath, runs SQL queryString, and closes database.
        /// </summary>
        /// <param name="filePath">File path and name of database</param>
        /// <param name="queryString">SQL query string to run against database</param>
        public void ReadAccessDb(string filePath, string queryString)
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source=" + filePath + ";" +
                @"User Id=;Password=;";
            
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (OleDbCommand command = new OleDbCommand(queryString, connection))
            {
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();

                    while(reader.Read())
                    {

                    }
                    reader.Close();
                }
                catch (System.Exception ex)
                {

                }
            }
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
            string table = "";
            string rowColor = "<td bgcolor''baa777'>";

            string sqlString = "SELECT CombinedRecords, CombinedDials, CombinedSaturation, CombinedPassPenetration, CombinedPassPenetrationTwo " +
                "FROM " + "ESDialsPerList " +
                "WHERE CallDate = #" + DateTime.Today + "#";

            table = "<TABLE BORDER='4' CELLSPACING='0' CELLPADDING='3' style='border-collapse: collapse; text-align:center; width: 225px; bordercolor='#111111'>" +
                "<TR HEIGHT=19><TD colspan='2' BGCOLOR='#4d4d4d'><FONT COLOR='FFFFFF'>List Progress Summary</TD></TR><tr><FONT COLOR='000000'>" +
                rowColor + "Intensity</td>" +
                rowColor + combinedSaturation.ToString() + "</td></FONT></tr>";

            table += "<tr><FONT COLOR='000000'>" +
                rowColor + "PID Goal</td>" +
                rowColor + (currentGoal * 100).ToString() + "%</td>" +
                "</FONT></tr>";

            table += "<tr><FONT COLOR='000000'>" +
                rowColor + "Pass Progression I</td>" +
                rowColor + (double.Parse(combinedPercentWorkedOne.ToString()) * 100) + "%</td>" +
                "</FONT></tr>";

            table += "<tr><FONT COLOR='000000'>" +
                rowColor + "Pass Progression II</td>" +
                rowColor + (double.Parse(combinedPercentWorkedTwo.ToString()) * 100) + "%</td>" +
                "</FONT></tr>";

            table += "<tr><FONT COLOR='000000'>" +
                rowColor + "Pass Progression III</td>" +
                rowColor + (double.Parse(combinedPercentWorkedThree.ToString()) * 100) + "%</td>" +
                "</FONT></tr></table><br>";

            return table;
        }
        #endregion

        #region Database Methods
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
            int currentHour = (int)DateTime.Now.Hour;

            double goal = 0.0;
            double[,] goals = new double[5,24];

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
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return goal;
        }
        #endregion


        #region Early Stage Tab Events
        private void sendWrapReportES_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void massCellRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            combinedRecords.Text = CalculateTotalRecords(CreateTotalRecordsList()).ToString();

            int records, dials = 0;

            records = Int32.Parse(massCellRecords.Text);
            dials = Int32.Parse(massCellDials.Text);

            if (records > 0 && dials > 0)
            {
                massCellSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void massCellDials_LostFocus(object sender, RoutedEventArgs e)
        {
            if (Int32.Parse(massCellDials.Text) > 0 && Int32.Parse(massCellRecords.Text) > 0)
            {
                double records, dials = 0;

                records = Double.Parse(massCellRecords.Text);
                dials = Double.Parse(massCellDials.Text);

                massCellSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";
            }

            combinedDials.Text = CalculateTotalDials(CreateTotalDialsList()).ToString();
        }

        private void SelectAllOnFocus_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBox box = (TextBox)sender;

            box.SelectAll();
        }

        private void massCellPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {

        }

        private void massCellPercentWorked_LostFocus_1(object sender, RoutedEventArgs e)
        {

        }

        private void miCellRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            combinedRecords.Text = CalculateTotalRecords(CreateTotalRecordsList()).ToString();

            int records, dials = 0;

            records = Int32.Parse(miCellRecords.Text);
            dials = Int32.Parse(miCellDials.Text);

            if (records > 0 && dials > 0)
            {
                miCellSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void miCellDials_LostFocus(object sender, RoutedEventArgs e)
        {
            combinedDials.Text = CalculateTotalDials(CreateTotalDialsList()).ToString();

            double records, dials = 0;

            records = Double.Parse(miCellRecords.Text);
            dials = Double.Parse(miCellDials.Text);

            if (records > 0 && dials > 0)
            {
                miCellSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";
            }

            if (records > 0)
            {
                int totalRecords = 0;
                totalRecords = Int32.Parse(combinedRecords.Text);

                miCellPercentTotal.Text = (Math.Round((CalculatePercentVolume(records, totalRecords)) * 100)).ToString() + "%";
            }

        }

        private void miCellPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {

        }

        private void allRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            combinedRecords.Text = CalculateTotalRecords(CreateTotalRecordsList()).ToString();

            int records, dials = 0;

            records = Int32.Parse(allRecords.Text);
            dials = Int32.Parse(allDials.Text);

            if (records > 0 && dials > 0)
            {
                allSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void allDials_LostFocus(object sender, RoutedEventArgs e)
        {
            if (Int32.Parse(allDials.Text) > 0 && Int32.Parse(allRecords.Text) > 0)
            {
                double records, dials = 0;

                records = Double.Parse(allRecords.Text);
                dials = Double.Parse(allDials.Text);

                allSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";
            }

            combinedDials.Text = CalculateTotalDials(CreateTotalDialsList()).ToString();
        }

        private void allPercentWorkedOne_LostFocus(object sender, RoutedEventArgs e)
        {

        }

        private void allPercentWorkedTwo_LostFocus(object sender, RoutedEventArgs e)
        {

        }

        private void allPercentWorkedThree_LostFocus(object sender, RoutedEventArgs e)
        {

        }

        private void massRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            combinedRecords.Text = CalculateTotalRecords(CreateTotalRecordsList()).ToString();

            int records, dials = 0;

            records = Int32.Parse(massRecords.Text);
            dials = Int32.Parse(massDials.Text);

            if (records > 0 && dials > 0)
            {
                massSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void massDials_LostFocus(object sender, RoutedEventArgs e)
        {
            if (Int32.Parse(massDials.Text) > 0 && Int32.Parse(massRecords.Text) > 0)
            {
                double records, dials = 0;

                records = Double.Parse(massRecords.Text);
                dials = Double.Parse(massDials.Text);

                massSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";
            }

            combinedDials.Text = CalculateTotalDials(CreateTotalDialsList()).ToString();
        }

        private void massPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {

        }

        private void nhRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            combinedRecords.Text = CalculateTotalRecords(CreateTotalRecordsList()).ToString();

            int records, dials = 0;

            records = Int32.Parse(nhRecords.Text);
            dials = Int32.Parse(nhDials.Text);

            if (records > 0 && dials > 0)
            {
                nhSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void nhDials_LostFocus(object sender, RoutedEventArgs e)
        {
            if (Int32.Parse(nhDials.Text) > 0 && Int32.Parse(nhRecords.Text) > 0)
            {
                double records, dials = 0;

                records = Double.Parse(nhRecords.Text);
                dials = Double.Parse(nhDials.Text);

                nhSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";
            }

            combinedDials.Text = CalculateTotalDials(CreateTotalDialsList()).ToString();
        }

        private void nhPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {

        }

        private void nonContRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            combinedRecords.Text = CalculateTotalRecords(CreateTotalRecordsList()).ToString();

            int records, dials = 0;

            records = Int32.Parse(nonContRecords.Text);
            dials = Int32.Parse(nonContDials.Text);

            if (records > 0 && dials > 0)
            {
                nonContSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void nonContDials_LostFocus(object sender, RoutedEventArgs e)
        {
            if (Int32.Parse(nonContDials.Text) > 0 && Int32.Parse(nonContRecords.Text) > 0)
            {
                double records, dials = 0;

                records = Double.Parse(nonContRecords.Text);
                dials = Double.Parse(nonContDials.Text);

                nonContSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";
            }

            combinedDials.Text = CalculateTotalDials(CreateTotalDialsList()).ToString();
        }

        private void nonContPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {

        }

        private void orRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            combinedRecords.Text = CalculateTotalRecords(CreateTotalRecordsList()).ToString();

            int records, dials = 0;

            records = Int32.Parse(orRecords.Text);
            dials = Int32.Parse(orDials.Text);

            if (records > 0 && dials > 0)
            {
                orSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void orDials_LostFocus(object sender, RoutedEventArgs e)
        {
            if (Int32.Parse(orDials.Text) > 0 && Int32.Parse(orRecords.Text) > 0)
            {
                double records, dials = 0;

                records = Double.Parse(orRecords.Text);
                dials = Double.Parse(orDials.Text);

                orSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString() + "%";
            }

            combinedDials.Text = CalculateTotalDials(CreateTotalDialsList()).ToString();
        }

        private void orPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {

        }
        #endregion

        #region Early Stage Calculations
        /// <summary>
        /// Calculates dialer saturation through calculation dials / records.
        /// </summary>
        /// <param name="records">Total number of records in list</param>
        /// <param name="dials">Number of dials attempted on accounts in list</param>
        /// <returns>Returns dials / records.</returns>
        public double CalculateSaturation(double records, double dials)
        {
            double saturation = 0.00;

            saturation = dials / records;

            return saturation;
        }

        private double CalculatePercentVolume(double records, double totalRecords)
        {
            double percentVolume = 0.0;

            percentVolume = records / totalRecords;

            return Math.Round(percentVolume, 2);
        }

        public List<double> CreateTotalRecordsList()
        {
            List<double> recordsList = new List<double>();

            recordsList.Add(Int32.Parse(massCellRecords.Text));
            recordsList.Add(Int32.Parse(miCellRecords.Text));
            recordsList.Add(Int32.Parse(allRecords.Text));
            recordsList.Add(Int32.Parse(massRecords.Text));
            recordsList.Add(Int32.Parse(nhRecords.Text));
            recordsList.Add(Int32.Parse(nonContRecords.Text));
            recordsList.Add(Int32.Parse(orRecords.Text));

            return recordsList;

            //totalRecords = Int32.Parse(massCellRecords.Text) + Int32.Parse(miCellRecords.Text) + Int32.Parse(allRecords.Text) + Int32.Parse(massRecords.Text) +Int32.Parse(nhRecords.Text) + Int32.Parse(nonContRecords.Text) + Int32.Parse(orRecords.Text);
        }

        public List<int> CreateTotalDialsList()
        {
            List<int> dialsList = new List<int>();

            dialsList.Add(Int32.Parse(massCellDials.Text));
            dialsList.Add(Int32.Parse(miCellDials.Text));
            dialsList.Add(Int32.Parse(allDials.Text));
            dialsList.Add(Int32.Parse(massDials.Text));
            dialsList.Add(Int32.Parse(nhDials.Text));
            dialsList.Add(Int32.Parse(nonContDials.Text));
            dialsList.Add(Int32.Parse(orDials.Text));

            return dialsList;

            //totalRecords = Int32.Parse(massCellRecords.Text) + Int32.Parse(miCellRecords.Text) + Int32.Parse(allRecords.Text) + Int32.Parse(massRecords.Text) +Int32.Parse(nhRecords.Text) + Int32.Parse(nonContRecords.Text) + Int32.Parse(orRecords.Text);
        }

        public List<TextBox> CreatePercentVolumeList()
        {
            List<TextBox> percentVolumeList = new List<TextBox>();

            percentVolumeList.Add(massCellPercentTotal);
            percentVolumeList.Add(miCellPercentTotal);
            percentVolumeList.Add(allPercentTotal);
            percentVolumeList.Add(massPercentTotal);
            percentVolumeList.Add(nhPercentTotal);
            percentVolumeList.Add(nonContPercentTotal);
            percentVolumeList.Add(orPercentTotal);

            return percentVolumeList;

            //totalRecords = Int32.Parse(massCellRecords.Text) + Int32.Parse(miCellRecords.Text) + Int32.Parse(allRecords.Text) + Int32.Parse(massRecords.Text) +Int32.Parse(nhRecords.Text) + Int32.Parse(nonContRecords.Text) + Int32.Parse(orRecords.Text);
        }

        public void CalculatePercentVolume(List<TextBox> volumes)
        {
            double totalRecords = CalculateTotalRecords(CreateTotalRecordsList());

            List<double> recordsList = CreateTotalRecordsList();
            List<TextBox> percentVolumeList = CreatePercentVolumeList();

            int i = 0;

            foreach(TextBox percentVolume in percentVolumeList)
            {
                percentVolume.Text = Math.Round((recordsList.ElementAt(i) / totalRecords) * 100.0).ToString() + "%";

                i++;
            }
        }

        public double CalculateTotalRecords(List<double> records)
        {
            int totalRecords = 0;

            foreach (int record in records)
            {
                totalRecords += record;
            }

            return totalRecords;

        }

        public int CalculateTotalDials(List<int> dials)
        {
            int totalDials = 0;

            foreach (int dial in dials)
            {
                totalDials += dial;
            }

            return totalDials;

        }

        #endregion

        #region TabControl Events
        private void earlyStageTab_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            Keyboard.Focus(massCellRecords);
            massCellRecords.SelectAll();
        }

        private void lateStageTab_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if(passOneCallsLS.Text == "0")
            {
                Keyboard.Focus(passOneCallsEntryLS);
                passOneCallsEntryLS.SelectAll();
            }
            else if(passOneCallsLS.Text != "0" && passTwoCallsLS.Text == "0")
            {
                Keyboard.Focus(passTwoCallsEntryLS);
                passTwoCallsEntryLS.SelectAll();
            }
            else
            {
                Keyboard.Focus(passThreeCallsEntryLS);
                passThreeCallsEntryLS.SelectAll();
            }
        }

        private void lossMitTab_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (tenCallsLM.Text == "0")
            {
                Keyboard.Focus(tenCallsEntryLM);
                tenCallsEntryLM.SelectAll();
            }
            else if (tenCallsLM.Text != "0" && twoCallsLM.Text == "0")
            {
                Keyboard.Focus(twoCallsEntryLM);
                twoCallsEntryLM.SelectAll();
            }
            else
            {
                Keyboard.Focus(sevenCallsEntryLM);
                sevenCallsEntryLM.SelectAll();
            }
        }
        #endregion

        #region Late Stage Tab Events
        private void passOneCallsEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            passOneCallsLS.Text = passOneCallsEntryLS.Text;
        }

        private void passOneConnectsEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            passOneConnectsLS.Text = passOneConnectsEntryLS.Text;
        }

        private void passOneContactEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            passOneContactLS.Text = passOneContactEntryLS.Text;
        }

        private void passOnePromiseEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            double rpcEntry, rpcTotal, ptp, closure = 0.0;

            ptp = Double.Parse(passOnePromiseEntryLS.Text);
            rpcEntry = Double.Parse(passOneContactEntryLS.Text);

            rpcTotal = rpcEntry + ptp;

            closure = ptp / rpcTotal;

            passOneContactLS.Text = rpcTotal.ToString();
            passOnePromiseLS.Text = ptp.ToString();
            passOneClosure.Text = Math.Round(closure * 100.0).ToString() + "%";
        }

        private void passTwoCallsEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            if(passOneCallsLS.Text != "0")
            {
                passTwoCallsLS.Text = (Double.Parse(passTwoCallsEntryLS.Text) - Double.Parse(passOneCallsLS.Text)).ToString();
            }
        }

        private void passTwoConnectsEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneConnectsLS.Text != "0")
            {
                passTwoConnectsLS.Text = (Double.Parse(passTwoConnectsEntryLS.Text) - Double.Parse(passOneConnectsLS.Text)).ToString();
            }
        }

        private void passTwoContactEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneContactLS.Text != "0")
            {
                passTwoContactLS.Text = (Double.Parse(passTwoContactLS.Text) - Double.Parse(passOneContactLS.Text)).ToString();
            }
        }

        private void passTwoPromiseEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            double rpcPrev, rpcEntry, rpcCurrent, rpcTotal, ptpPrev, ptpEntry, ptpTotal, closure = 0.0;

            ptpEntry = Double.Parse(passTwoPromiseEntryLS.Text);
            rpcEntry = Double.Parse(passTwoContactEntryLS.Text);

            ptpPrev = Double.Parse(passOnePromiseLS.Text);
            rpcPrev = Double.Parse(passOneContactLS.Text);

            ptpTotal = ptpEntry - ptpPrev;
            rpcTotal = rpcEntry + ptpEntry;

            rpcCurrent = rpcTotal - rpcPrev;

            closure = ptpTotal / rpcCurrent;

            passTwoContactLS.Text = rpcCurrent.ToString();
            passTwoPromiseLS.Text = ptpTotal.ToString();
            passTwoClosure.Text = Math.Round(closure * 100.0).ToString() + "%";
        }

        private void passThreeCallsEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneCallsLS.Text != "0" && passTwoCallsLS.Text != "0")
            {
                passThreeCallsLS.Text = (Double.Parse(passThreeCallsEntryLS.Text) - Double.Parse(passTwoCallsLS.Text) - Double.Parse(passOneCallsLS.Text)).ToString();
            }
        }

        private void passThreeConnectsEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneConnectsLS.Text != "0" && passTwoConnectsLS.Text != "0")
            {
                passThreeConnectsLS.Text = (Double.Parse(passThreeConnectsEntryLS.Text) - Double.Parse(passTwoConnectsLS.Text) - Double.Parse(passOneConnectsLS.Text)).ToString();
            }
        }

        private void passThreeContactEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneContactLS.Text != "0" && passTwoContactLS.Text != "0")
            {
                passThreeContactLS.Text = (Double.Parse(passThreeContactLS.Text) - Double.Parse(passTwoContactLS.Text) - Double.Parse(passOneContactLS.Text)).ToString();
            }
        }

        private void passThreePromiseEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            double rpcPrev, rpcEntry, rpcCurrent, rpcTotal, ptpPrev, ptpEntry, ptpTotal, closure = 0.0;

            ptpEntry = Double.Parse(passThreePromiseEntryLS.Text);
            rpcEntry = Double.Parse(passThreeContactEntryLS.Text);

            ptpPrev = Double.Parse(passTwoPromiseLS.Text) + Double.Parse(passOnePromiseLS.Text);
            rpcPrev = Double.Parse(passTwoContactLS.Text) + Double.Parse(passOneContactLS.Text);

            ptpTotal = ptpEntry - ptpPrev;
            rpcTotal = rpcEntry + ptpEntry;

            rpcCurrent = rpcTotal - rpcPrev;

            closure = ptpTotal / rpcCurrent;

            passThreeContactLS.Text = rpcCurrent.ToString();
            passThreePromiseLS.Text = ptpTotal.ToString();
            passThreeClosure.Text = Math.Round(closure * 100.0).ToString() + "%";
        }
        #endregion

        #region Loss Mit Tab Events
        private void tenCallsEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            tenCallsLM.Text = tenCallsEntryLM.Text;
        }

        private void tenConnectsEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            tenConnectsLM.Text = tenConnectsEntryLM.Text;
        }

        private void tenContactEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            tenContactLM.Text = tenContactEntryLM.Text;
        }

        private void tenPromiseEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            double rpcEntry, rpcTotal, ptp, closure = 0.0;

            ptp = Double.Parse(tenPromiseEntryLM.Text);
            rpcEntry = Double.Parse(tenContactEntryLM.Text);

            rpcTotal = rpcEntry + ptp;

            closure = ptp / rpcTotal;

            tenContactLM.Text = rpcTotal.ToString();
            tenPromiseLM.Text = ptp.ToString();
            tenClosure.Text = Math.Round(closure * 100.0).ToString() + "%";
        }

        private void twoCallsEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            if (tenCallsLM.Text != "0")
            {
                twoCallsLM.Text = (Double.Parse(twoCallsEntryLM.Text) - Double.Parse(tenCallsLM.Text)).ToString();
            }
        }

        private void twoConnectsEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            if (tenConnectsLM.Text != "0")
            {
                twoConnectsLM.Text = (Double.Parse(twoConnectsEntryLM.Text) - Double.Parse(tenConnectsLM.Text)).ToString();
            }
        }

        private void twoContactEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            if (tenContactLM.Text != "0")
            {
                twoContactLM.Text = (Double.Parse(twoContactLM.Text) - Double.Parse(tenContactLM.Text)).ToString();
            }
        }

        private void twoPromiseEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            double rpcPrev, rpcEntry, rpcCurrent, rpcTotal, ptpPrev, ptpEntry, ptpTotal, closure = 0.0;

            ptpEntry = Double.Parse(twoPromiseEntryLM.Text);
            rpcEntry = Double.Parse(twoContactEntryLM.Text);

            ptpPrev = Double.Parse(tenPromiseLM.Text);
            rpcPrev = Double.Parse(tenContactLM.Text);

            ptpTotal = ptpEntry - ptpPrev;
            rpcTotal = rpcEntry + ptpEntry;

            rpcCurrent = rpcTotal - rpcPrev;

            closure = ptpTotal / rpcCurrent;

            twoContactLM.Text = rpcCurrent.ToString();
            twoPromiseLM.Text = ptpTotal.ToString();
            twoClosure.Text = Math.Round(closure * 100.0).ToString() + "%";
        }

        private void sevenCallsEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            if (tenCallsLM.Text != "0" && twoCallsLM.Text != "0")
            {
                sevenCallsLM.Text = (Double.Parse(sevenCallsEntryLM.Text) - Double.Parse(twoCallsLM.Text) - Double.Parse(tenCallsLM.Text)).ToString();
            }
        }

        private void sevenConnectsEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            if (tenConnectsLM.Text != "0" && twoConnectsLM.Text != "0")
            {
                sevenConnectsLM.Text = (Double.Parse(sevenConnectsEntryLM.Text) - Double.Parse(twoConnectsLM.Text) - Double.Parse(tenConnectsLM.Text)).ToString();
            }
        }

        private void sevenContactEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            if (tenContactLM.Text != "0" && twoContactLM.Text != "0")
            {
                sevenContactLM.Text = (Double.Parse(sevenContactLM.Text) - Double.Parse(twoContactLM.Text) - Double.Parse(tenContactLM.Text)).ToString();
            }
        }

        private void sevenPromiseEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            double rpcPrev, rpcEntry, rpcCurrent, rpcTotal, ptpPrev, ptpEntry, ptpTotal, closure = 0.0;

            ptpEntry = Double.Parse(sevenPromiseEntryLM.Text);
            rpcEntry = Double.Parse(sevenContactEntryLM.Text);

            ptpPrev = Double.Parse(twoPromiseLM.Text) + Double.Parse(tenPromiseLM.Text);
            rpcPrev = Double.Parse(twoContactLM.Text) + Double.Parse(tenContactLM.Text);

            ptpTotal = ptpEntry - ptpPrev;
            rpcTotal = rpcEntry + ptpEntry;

            rpcCurrent = rpcTotal - rpcPrev;

            closure = ptpTotal / rpcCurrent;

            sevenContactLM.Text = rpcCurrent.ToString();
            sevenPromiseLM.Text = ptpTotal.ToString();
            sevenClosure.Text = Math.Round(closure * 100.0).ToString() + "%";
        }

        #endregion

        #region StackPanel Events
        private void sendWrap_Click(object sender, RoutedEventArgs e)
        {
            if(departmentSelector.SelectedIndex == 2)
            {
                List<WrapReport> report = new List<WrapReport>();
                report = ImportWrap("C:\\Dialer Team Back-End Database\\early stage.xls");

                ExportWrap(report);
            }
            else if (departmentSelector.SelectedIndex == 3)
            {
                List<WrapReport> report = new List<WrapReport>();
                report = ImportWrap("C:\\Dialer Team Back-End Database\\late stage.xls");

                ExportWrap(report);
            }
            else if (departmentSelector.SelectedIndex == 4)
            {
                List<WrapReport> report = new List<WrapReport>();
                report = ImportWrap("C:\\Dialer Team Back-End Database\\loss mit.xls");

                ExportWrap(report);
            }
        }

        private void sendEmail_Click(object sender, RoutedEventArgs e)
        {
            if (departmentSelector.SelectedIndex == 2)
            {
                EmailUpdate();
            }
            else if (departmentSelector.SelectedIndex == 3)
            {
                List<WrapReport> report = new List<WrapReport>();
                report = ImportWrap("C:\\Dialer Team Back-End Database\\late stage.xls");

                ExportWrap(report);
            }
            else if (departmentSelector.SelectedIndex == 4)
            {
                List<WrapReport> report = new List<WrapReport>();
                report = ImportWrap("C:\\Dialer Team Back-End Database\\loss mit.xls");

                ExportWrap(report);
            }
        }
        #endregion

    }

    class OutlookWriterInterop
    {
        Microsoft.Office.Interop.Outlook.Application outlookApp;

        public OutlookWriterInterop()
        {
            outlookApp = new Microsoft.Office.Interop.Outlook.Application();
        }

        public void CreateIntradayUpdate(string messageBody)
        {
            Microsoft.Office.Interop.Outlook.MailItem outlookMail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);
            
            outlookMail.To = "Keven.McBarnes@td.com";
            outlookMail.Subject = "List Penetration Table";
            outlookMail.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
            outlookMail.HTMLBody = messageBody;
            outlookMail.Display();
        }

        public void CreateWrapEmail(List<WrapReport> report)
        {
            Microsoft.Office.Interop.Outlook.MailItem outlookMail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);
            string wrapHighlight = "";
            string rowColor = "<td bgcolor='#FFFFFF'>&nbsp;";

            string msgBody = "";

            msgBody = "<TABLE BORDER='4' CELLSPACING='0' CELLPADDING='0' style='border-collapse: collapse; text-align:center; width: 900px bordercolor='#111111'>" +
                "<FONT COLOR='000000' face='ARIAL' size='2'><b><TR HEIGHT=10 >" +
                "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 100px'>&nbsp;TID</TD>" +
                "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 275px'>&nbsp;Name</TD>" +
                "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 52px'>&nbsp;Calls</TD>" +
                "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Login</TD>" +
                "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Active</TD>" +
                "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Not Ready</TD>" +
                "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Idle</TD>" +
                "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Wrap</TD>" +
                "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Hold</TD>" +
                "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 85px'>&nbsp;Hold Count</TD>" +
                "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Avg Wrap</FONT></b></TD></TR>";

            foreach (WrapReport agent in report)
            {
                if (agent.AvgWrap.TotalSeconds >= 20)
                {
                    wrapHighlight = "<td bgcolor='#da9694'>&nbsp;";
                }

                msgBody += "<tr><FONT COLOR='000000' face='Arial' size='2'>" +
                    "<td bgcolor='#FFFFFF' align='center'>" + agent.Tid + "</td>" +
                    "<td bgcolor='#FFFFFF' align='center'>" + agent.Name + "</td>" +
                    rowColor + agent.Calls.ToString() + "</td>" +
                    rowColor + agent.Login.ToString() + "</td>" +
                    rowColor + agent.Active.ToString() + "</td>" +
                    rowColor + agent.NotReady.ToString() + "</td>" +
                    rowColor + agent.Idle.ToString() + "</td>" +
                    rowColor + agent.Wrap.ToString() + "</td>" +
                    rowColor + agent.Hold.ToString() + "</td>" +
                    rowColor + agent.HoldCount.ToString() + "</td>" +
                    wrapHighlight + agent.AvgWrap.ToString() + "</td>" +
                    "</FONT></tr>";

                wrapHighlight = "<td bgcolor='#ffffff'>&nbsp;";
            }

            //			If Not ((rs.Fields("TID - Name") = 0) Or (rs.Fields("TID - Name") = "Report Total")) Then
            //
            //            If rs.Fields("Avg Wrap") >= #12:00:20 AM# Then
            //                wrapHighlight = "<td bgcolor='#da9694'>&nbsp;"
            //            End If
            //
            //	        esTable = esTable & "<tr><FONT COLOR='000000' face='Arial' size='2'>" & _
            //	            "<td bgcolor='#FFFFFF' align='left'>&nbsp;" & rs.Fields("TID - Name") & "</td>" & _
            //	            rowColor & rs.Fields("Calls") & "</td>" & _
            //	            rowColor & Format(rs.Fields("Login"), "h:mm:ss") & "</td>" & _
            //	            rowColor & Format(rs.Fields("Active"), "h:mm:ss") & "</td>" & _
            //	            rowColor & Format(rs.Fields("Not Ready"), "h:mm:ss") & "</td>" & _
            //	            rowColor & Format(rs.Fields("Idle"), "h:mm:ss") & "</td>" & _
            //	            rowColor & Format(rs.Fields("Wrap"), "h:mm:ss") & "</td>" & _
            //	            rowColor & Format(rs.Fields("Hold"), "h:mm:ss") & "</td>" & _
            //	            rowColor & rs.Fields("Hold Count") & "</td>" & _
            //	            wrapHighlight & Format(rs.Fields("Avg Wrap"), "h:mm:ss") & "</td>" & _
            //	            "</FONT></tr>"
            //
            //        End If

            outlookMail.To = "Keven.McBarnes@td.com";
            outlookMail.Subject = "Wrap Report";
            outlookMail.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
            outlookMail.HTMLBody = msgBody;
            outlookMail.Display();
        }

        //		Microsoft.Office.Interop.Outlook.MailItem olkMail1 =
        //    (MailItem)olkApp1.CreateItem(OlItemType.olMailItem);
        //        olkMail1.To = txtpsnum.Text;
        //        olkMail1.CC = "";
        //        olkMail1.Subject = "Assignment note";
        //        olkMail1.Body = "Assignment note";
        //        olkMail1.Attachments.Add(AssignNoteFilePath, 
        //            Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, 1, 
        //                "Assignment_note");
        //olkMail1.Save();
        //olkMail.Send();
    }

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
                book.Close(false, fileName, null);
                Marshal.ReleaseComObject(book);

                return report;
            }
            catch (System.Exception ex)
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

        private List<WrapReport> ProcessObjects(object[,] valueArray)
        {
            List<WrapReport> report = new List<WrapReport>();

            string identifier = "";

            //valueArray.Length avoid exception being thrown from final row in source sheet having a null value in column 1			
            for (int i = 1; i < valueArray.GetLength(0); i++)
            {
                WrapReport agentWrapReport = new WrapReport();

                //Create string identifier to check whether line contains agent data
                identifier = ((string)valueArray[i, 1]).Substring(0, 1);

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

    public class WrapReport
    {
        string tid;
        string name;
        int calls;
        TimeSpan login;
        TimeSpan active;
        TimeSpan notReady;
        TimeSpan idle;
        TimeSpan wrap;
        TimeSpan hold;
        int holdCount;
        TimeSpan avgWrap;

        public string Tid
        {
            get { return tid; }
            set { tid = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public int Calls
        {
            get { return calls; }
            set { calls = value; }
        }

        public TimeSpan Login
        {
            get { return login; }
            set { login = value; }
        }

        public TimeSpan Active
        {
            get { return active; }
            set { login = value; }
        }

        public TimeSpan NotReady
        {
            get { return notReady; }
            set { notReady = value; }
        }

        public TimeSpan Idle
        {
            get { return idle; }
            set { idle = value; }
        }

        public TimeSpan Wrap
        {
            get { return wrap; }
            set { wrap = value; }
        }

        public TimeSpan Hold
        {
            get { return hold; }
            set { hold = value; }
        }

        public int HoldCount
        {
            get { return holdCount; }
            set { holdCount = value; }
        }

        public TimeSpan AvgWrap
        {
            get { return avgWrap; }
            set { avgWrap = value; }
        }

        public WrapReport()
        {
            tid = "";
            name = "";
            calls = 0;
            login = new TimeSpan();
            active = new TimeSpan();
            notReady = new TimeSpan();
            idle = new TimeSpan();
            wrap = new TimeSpan();
            hold = new TimeSpan();
            holdCount = 0;
            avgWrap = new TimeSpan();
        }

        public WrapReport(string cvTid, string cvName, int cvCalls, TimeSpan cvLogin, TimeSpan cvActive, TimeSpan cvNotReady,
                          TimeSpan cvIdle, TimeSpan cvWrap, TimeSpan cvHold, int cvHoldCount, TimeSpan cvAvgWrap)
        {
            tid = cvTid;
            name = cvName;
            calls = cvCalls;
            login = cvLogin;
            active = cvActive;
            notReady = cvNotReady;
            idle = cvIdle;
            wrap = cvWrap;
            hold = cvHold;
            holdCount = cvHoldCount;
            avgWrap = cvAvgWrap;
        }

        public WrapReport(int cvCalls, TimeSpan cvWrap)
        {
            calls = cvCalls;
            wrap = cvWrap;
        }

        public TimeSpan CalculateAverageWrap(TimeSpan wrapTime, int calls)
        {
            //double avgWrap = 0;
            TimeSpan avgWrap;

            avgWrap = new TimeSpan(0, 0, (int)(wrapTime.TotalSeconds / calls));

            return avgWrap;
        }

        //		public bool Sort(TimeSpan a, TimeSpan b)
        //		{
        //			if(a < b)
        //			{
        //				return true;
        //			}
        //			else if(b < a)
        //			{
        //				return false;
        //			}
        //			else return true;
        //			
        //		}

        public override string ToString()
        {
            string output = "";

            return output;
        }

    }
}
