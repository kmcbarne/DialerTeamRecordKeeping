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
    public partial class MainWindow : System.Windows.Window
    {
        #region Datagrid Events
        private void dialerRecordsDataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if ((e.PropertyType == typeof(DateTime)) && ((string)e.Column.Header == "StartTime") || ((string)e.Column.Header == "EndTime"))
            {
                ((DataGridTextColumn)e.Column).Binding.StringFormat = "hh:mm";
            }
            else if ((e.PropertyType == typeof(DateTime)) && (string)e.Column.Header == "ElapsedTime")
            {
                ((DataGridTextColumn)e.Column).Binding.StringFormat = "mm";
            }
            else if ((e.PropertyType == typeof(DateTime)) && (string)e.Column.Header == "DialerDate")
            {
                ((DataGridTextColumn)e.Column).Binding.StringFormat = "MM/dd/yyyy";
            }
        }

        private void dialerRecordsSelect_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                string queryString = "SELECT * FROM DialerRecords";
                string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    @"Data Source=" + backendFilePath + ";" +
                    @"User Id=;Password=;";

                DateTime filterDate = DateTime.Parse(dialerRecordsSelect.Text);

                FillDataGrid(dialerRecordsDataGrid, queryString, connectionString, filterDate);

                dialerRecordsDataGrid.Columns.ElementAt(8).Width = 200;
            }
        }

        private void dialerRecordsSelect_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            string queryString = "SELECT * FROM DialerRecords";
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source=" + backendFilePath + ";" +
                @"User Id=;Password=;";

            DateTime filterDate = DateTime.Parse(dialerRecordsSelect.Text);

            FillDataGrid(dialerRecordsDataGrid, queryString, connectionString, filterDate);

            dialerRecordsDataGrid.Columns.RemoveAt(11);
            dialerRecordsDataGrid.Columns.RemoveAt(10);
            dialerRecordsDataGrid.Columns.RemoveAt(0);
            dialerRecordsDataGrid.Columns.ElementAt(8).Width = 200;
        }

        private void instrumentIdDataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if ((string)e.Column.Header == "AgentTID")
            {
                e.Column.Header = "Agent TID";
            }
            else if ((string)e.Column.Header == "LastName")
            {
                e.Column.Header = "Last Name";
            }
            else if ((string)e.Column.Header == "FirstName")
            {
                e.Column.Header = "First Name";
            }
            else if ((string)e.Column.Header == "Team")
            {
                e.Column.Header = "Team";
            }
            else if ((string)e.Column.Header == "InstrumentID")
            {
                e.Column.Header = "Instrument ID";
            }
            else if ((string)e.Column.Header == "OnDialerExtension")
            {
                e.Column.Header = "On-Dialer";
            }
            else if ((string)e.Column.Header == "OffDialerExtension")
            {
                e.Column.Header = "Off-Dialer";
            }
        }
        #endregion

        #region Early Stage Records Lost Focus Events
        private void allRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            int records = int.Parse(allRecordsOne.Text);
            int dials = int.Parse(allDialsOne.Text);

            if (records > 0 && dials > 0)
            {
                allSaturationOne.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void massCellRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            int records, dials;

            records = int.Parse(massCellRecords.Text);
            dials = int.Parse(massCellDials.Text);

            if (records > 0 && dials > 0)
            {
                massCellSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void massRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            int records = int.Parse(massRecords.Text);
            int dials = int.Parse(massDials.Text);

            if (records > 0 && dials > 0)
            {
                massSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void miCellRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            int records = int.Parse(miCellRecords.Text);
            int dials = int.Parse(miCellDials.Text);

            if (records > 0 && dials > 0)
            {
                miCellSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void nhRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            int records = int.Parse(nhRecords.Text);
            int dials = int.Parse(nhDials.Text);

            if (records > 0 && dials > 0)
            {
                nhSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void nonContRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            int records = int.Parse(nonContRecords.Text);
            int dials = int.Parse(nonContDials.Text);

            if (records > 0 && dials > 0)
            {
                nonContSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }

        private void orRecords_LostFocus(object sender, RoutedEventArgs e)
        {
            int records = int.Parse(orRecords.Text);
            int dials = int.Parse(orDials.Text);

            if (records > 0 && dials > 0)
            {
                orSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();

            }

            CalculatePercentVolume(CreatePercentVolumeList());
        }
        #endregion

        #region Early Stage Dials Lost Focus Events
        private void allDials_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.Parse(allDialsOne.Text) > 0 && int.Parse(allRecordsOne.Text) > 0)
            {
                double records = double.Parse(allRecordsOne.Text);
                double dials = double.Parse(allDialsOne.Text);

                allSaturationOne.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();
            }

            combinedDials.Text = CalculateTotalDials(CreateDialsList()).ToString();
        }

        private void allDialsOne_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.Parse(allDialsOne.Text) > 0 && int.Parse(allRecordsOne.Text) > 0)
            {
                double records = double.Parse(allRecordsOne.Text);
                double dials = double.Parse(allDialsOne.Text);

                allSaturationOne.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();
            }

            //combinedDials.Text = CalculateTotalDials(CreateDialsList()).ToString();
        }

        private void allDialsTwo_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.Parse(allDialsTwo.Text) > 0 && int.Parse(allRecordsTwo.Text) > 0)
            {
                double records = double.Parse(allRecordsTwo.Text);
                double dials = double.Parse(allDialsTwo.Text);

                allSaturationTwo.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();
            }

            //combinedDials.Text = CalculateTotalDials(CreateDialsList()).ToString();
        }

        private void allDialsThree_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.Parse(allDialsThree.Text) > 0 && int.Parse(allRecordsThree.Text) > 0)
            {
                double records = double.Parse(allRecordsThree.Text);
                double dials = double.Parse(allDialsThree.Text);

                allSaturationThree.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();
            }

            //combinedDials.Text = CalculateTotalDials(CreateDialsList()).ToString();
        }

        private void massCellDials_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.Parse(massCellDials.Text) > 0 && int.Parse(massCellRecords.Text) > 0)
            {
                double records = double.Parse(massCellRecords.Text);
                double dials = double.Parse(massCellDials.Text);

                massCellSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();
            }

            combinedDials.Text = CalculateTotalDials(CreateDialsList()).ToString();
        }

        private void massDials_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.Parse(massDials.Text) > 0 && int.Parse(massRecords.Text) > 0)
            {
                double records = double.Parse(massRecords.Text);
                double dials = double.Parse(massDials.Text);

                massSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();
            }

            combinedDials.Text = CalculateTotalDials(CreateDialsList()).ToString();
        }

        private void miCellDials_LostFocus(object sender, RoutedEventArgs e)
        {
            combinedDials.Text = CalculateTotalDials(CreateDialsList()).ToString();

            double records = double.Parse(miCellRecords.Text);
            double dials = double.Parse(miCellDials.Text);

            if (records > 0 && dials > 0)
            {
                miCellSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();
            }

            if (records > 0)
            {
                int totalRecords = int.Parse(combinedRecords.Text);

                miCellPercentTotal.Text = (Math.Round((CalculatePercentVolume(records, totalRecords)) * 100)).ToString();
            }

        }

        private void nhDials_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.Parse(nhDials.Text) > 0 && int.Parse(nhRecords.Text) > 0)
            {
                double records = double.Parse(nhRecords.Text);
                double dials = double.Parse(nhDials.Text);

                nhSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();
            }

            combinedDials.Text = CalculateTotalDials(CreateDialsList()).ToString();
        }

        private void nonContDials_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.Parse(nonContDials.Text) > 0 && int.Parse(nonContRecords.Text) > 0)
            {
                double records = double.Parse(nonContRecords.Text);
                double dials = double.Parse(nonContDials.Text);

                nonContSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();
            }

            combinedDials.Text = CalculateTotalDials(CreateDialsList()).ToString();
        }

        private void orDials_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.Parse(orDials.Text) > 0 && int.Parse(orRecords.Text) > 0)
            {
                double records = double.Parse(orRecords.Text);
                double dials = double.Parse(orDials.Text);

                orSaturation.Text = (Math.Round((CalculateSaturation(records, dials)) * 100)).ToString();
            }

            combinedDials.Text = CalculateTotalDials(CreateDialsList()).ToString();
        }
        #endregion

        #region Early Stage Saturation Lost Focus Events
        #endregion

        #region Early Stage Percent Total Lost Focus Events
        #endregion

        #region Early Stage Percent Worked Lost Focus Events
        private void allPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {
            double records = double.Parse(allRecords.Text);
            double totalRecords = double.Parse(combinedRecords.Text);
            double percentVolume = CalculatePercentVolumeRaw(double.Parse(allRecords.Text), double.Parse(combinedRecords.Text));
            double passPenetration = double.Parse(allPercentWorked.Text);

            DisplayPassPercent(percentVolume, passPenetration, allPassPercent);
        }

        private void allPercentWorkedOne_LostFocus(object sender, RoutedEventArgs e)
        {
            double records = double.Parse(allRecordsOne.Text);
            double totalRecords = double.Parse(combinedRecords.Text);
            double percentVolume = CalculatePercentVolumeRaw(double.Parse(allRecordsOne.Text), double.Parse(combinedRecords.Text));
            double passPenetration = double.Parse(allPercentWorkedOne.Text);

            DisplayPassPercent(percentVolume, passPenetration, allPassPercentOne);
        }

        private void allPercentWorkedTwo_LostFocus(object sender, RoutedEventArgs e)
        {
            double records = double.Parse(allRecordsOne.Text);
            double totalRecords = double.Parse(combinedRecords.Text);
            double percentVolume = CalculatePercentVolumeRaw(double.Parse(allRecordsOne.Text), double.Parse(combinedRecords.Text));
            double passPenetration = double.Parse(allPercentWorkedTwo.Text);

            DisplayPassPercent(percentVolume, passPenetration, allPassPercentTwo);
        }

        private void allPercentWorkedThree_LostFocus(object sender, RoutedEventArgs e)
        {
            double records = double.Parse(allRecordsOne.Text);
            double totalRecords = double.Parse(combinedRecords.Text);
            double percentVolume = CalculatePercentVolumeRaw(double.Parse(allRecordsOne.Text), double.Parse(combinedRecords.Text));
            double passPenetration = double.Parse(allPercentWorkedThree.Text);

            DisplayPassPercent(percentVolume, passPenetration, allPassPercentThree);
        }

        private void massCellPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {
            double records = double.Parse(massCellRecords.Text);
            double totalRecords = double.Parse(combinedRecords.Text);
            double percentVolume = CalculatePercentVolumeRaw(double.Parse(massCellRecords.Text), double.Parse(combinedRecords.Text));
            double passPenetration = double.Parse(massCellPercentWorked.Text);

            DisplayPassPercent(percentVolume, passPenetration, massCellPassPercent);
        }

        private void massPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {
            double records = double.Parse(massRecords.Text);
            double totalRecords = double.Parse(combinedRecords.Text);
            double percentVolume = CalculatePercentVolumeRaw(double.Parse(massRecords.Text), double.Parse(combinedRecords.Text));
            double passPenetration = double.Parse(massPercentWorked.Text);

            DisplayPassPercent(percentVolume, passPenetration, massPassPercent);
        }

        private void miCellPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {
            //double records = double.Parse(miCellRecords.Text);
            //double totalRecords = double.Parse(combinedRecords.Text);
            //double percentVolume = CalculatePercentVolumeRaw(double.Parse(miCellRecords.Text), double.Parse(combinedRecords.Text));
            //double passPenetration = double.Parse(miCellPercentWorked.Text);

            //DisplayPassPercent(percentVolume, passPenetration, miCellPassPercent);
        }

        private void nhPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {
            double records = double.Parse(nhRecords.Text);
            double totalRecords = double.Parse(combinedRecords.Text);
            double percentVolume = CalculatePercentVolumeRaw(double.Parse(nhRecords.Text), double.Parse(combinedRecords.Text));
            double passPenetration = double.Parse(nhPercentWorked.Text);

            DisplayPassPercent(percentVolume, passPenetration, nhPassPercent);
        }

        private void nonContPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {
            double records = double.Parse(nonContRecords.Text);
            double totalRecords = double.Parse(combinedRecords.Text);
            double percentVolume = CalculatePercentVolumeRaw(double.Parse(nonContRecords.Text), double.Parse(combinedRecords.Text));
            double passPenetration = double.Parse(nonContPercentWorked.Text);

            DisplayPassPercent(percentVolume, passPenetration, nonContPassPercent);
        }

        private void orPercentWorked_LostFocus(object sender, RoutedEventArgs e)
        {
            double records = double.Parse(orRecords.Text);
            double totalRecords = double.Parse(combinedRecords.Text);
            double percentVolume = CalculatePercentVolumeRaw(double.Parse(orRecords.Text), double.Parse(combinedRecords.Text));
            double passPenetration = double.Parse(orPercentWorked.Text);

            DisplayPassPercent(percentVolume, passPenetration, orPassPercent);
        }
        #endregion

        #region Late Stage Calls Entry Lost Focus Events
        private void passOneCallsEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            passOneCallsLS.Text = passOneCallsEntryLS.Text;
        }

        private void passTwoCallsEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneCallsLS.Text != "0")
            {
                passTwoCallsLS.Text = (double.Parse(passTwoCallsEntryLS.Text) - double.Parse(passOneCallsLS.Text)).ToString();
            }
        }

        private void passThreeCallsEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneCallsLS.Text != "0" && passTwoCallsLS.Text != "0")
            {
                passThreeCallsLS.Text = (double.Parse(passThreeCallsEntryLS.Text) - double.Parse(passTwoCallsLS.Text) - double.Parse(passOneCallsLS.Text)).ToString();
            }
        }
        #endregion

        #region Late Stage Connects Entry Lost Focus Events
        private void passOneConnectsEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            passOneConnectsLS.Text = passOneConnectsEntryLS.Text;
        }

        private void passTwoConnectsEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneConnectsLS.Text != "0")
            {
                passTwoConnectsLS.Text = (double.Parse(passTwoConnectsEntryLS.Text) - double.Parse(passOneConnectsLS.Text)).ToString();
            }
        }

        private void passThreeConnectsEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneConnectsLS.Text != "0" && passTwoConnectsLS.Text != "0")
            {
                passThreeConnectsLS.Text = (double.Parse(passThreeConnectsEntryLS.Text) - double.Parse(passTwoConnectsLS.Text) - double.Parse(passOneConnectsLS.Text)).ToString();
            }
        }
        #endregion

        #region Late Stage Contact Entry Lost Focus Events
        private void passOneContactEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            passOneContactLS.Text = passOneContactEntryLS.Text;
        }

        private void passTwoContactEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneContactLS.Text != "0")
            {
                passTwoContactLS.Text = (double.Parse(passTwoContactLS.Text) - double.Parse(passOneContactLS.Text)).ToString();
            }
        }

        private void passThreeContactEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneContactLS.Text != "0" && passTwoContactLS.Text != "0")
            {
                passThreeContactLS.Text = (double.Parse(passThreeContactLS.Text) - double.Parse(passTwoContactLS.Text) - double.Parse(passOneContactLS.Text)).ToString();
            }
        }
        #endregion

        #region Late Stage Promise Entry Lost Focus Events
        private void passOnePromiseEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            double ptp = double.Parse(passOnePromiseEntryLS.Text);
            double rpcEntry = double.Parse(passOneContactEntryLS.Text);
            double rpcTotal = rpcEntry + ptp;
            double closure = ptp / rpcTotal;

            passOneContactLS.Text = rpcTotal.ToString();
            passOnePromiseLS.Text = ptp.ToString();
            passOneClosure.Text = Math.Round(closure * 100.0).ToString() + "%";
        }

        private void passTwoPromiseEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            double ptpEntry = double.Parse(passTwoPromiseEntryLS.Text);
            double rpcEntry = double.Parse(passTwoContactEntryLS.Text);
            double ptpPrev = double.Parse(passOnePromiseLS.Text);
            double rpcPrev = double.Parse(passOneContactLS.Text);
            double ptpTotal = ptpEntry - ptpPrev;
            double rpcTotal = rpcEntry + ptpEntry;
            double rpcCurrent = rpcTotal - rpcPrev;
            double closure = ptpTotal / rpcCurrent;

            passTwoContactLS.Text = rpcCurrent.ToString();
            passTwoPromiseLS.Text = ptpTotal.ToString();
            passTwoClosure.Text = Math.Round(closure * 100.0).ToString() + "%";
        }

        private void passThreePromiseEntryLS_LostFocus(object sender, RoutedEventArgs e)
        {
            double rpcPrev, rpcEntry, rpcCurrent, rpcTotal, ptpPrev, ptpEntry, ptpTotal, closure = 0.0;

            ptpEntry = double.Parse(passThreePromiseEntryLS.Text);
            rpcEntry = double.Parse(passThreeContactEntryLS.Text);

            ptpPrev = double.Parse(passTwoPromiseLS.Text) + double.Parse(passOnePromiseLS.Text);
            rpcPrev = double.Parse(passTwoContactLS.Text) + double.Parse(passOneContactLS.Text);

            ptpTotal = ptpEntry - ptpPrev;
            rpcTotal = rpcEntry + ptpEntry;

            rpcCurrent = rpcTotal - rpcPrev;

            closure = ptpTotal / rpcCurrent;

            passThreeContactLS.Text = rpcCurrent.ToString();
            passThreePromiseLS.Text = ptpTotal.ToString();
            passThreeClosure.Text = Math.Round(closure * 100.0).ToString() + "%";
        }
        #endregion

        #region Loss Mit Calls Entry Lost Focus Events
        private void passOneCallsEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            passOneCallsLM.Text = passOneCallsEntryLM.Text;
        }

        private void passTwoCallsEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneCallsLM.Text != "0")
            {
                passTwoCallsLM.Text = (double.Parse(passTwoCallsEntryLM.Text) - double.Parse(passOneCallsLM.Text)).ToString();
            }
        }

        private void passThreeCallsEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneCallsLM.Text != "0" && passTwoCallsLM.Text != "0")
            {
                passThreeCallsLM.Text = (double.Parse(passThreeCallsEntryLM.Text) - double.Parse(passTwoCallsLM.Text) - double.Parse(passOneCallsLM.Text)).ToString();
            }
        }
        #endregion

        #region Loss Mit Connects Entry Lost Focus Events
        private void passOneConnectsEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            passOneConnectsLM.Text = passOneConnectsEntryLM.Text;
        }

        private void passTwoConnectsEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneConnectsLM.Text != "0")
            {
                passTwoConnectsLM.Text = (double.Parse(passTwoConnectsEntryLM.Text) - double.Parse(passOneConnectsLM.Text)).ToString();
            }
        }

        private void passThreeConnectsEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneConnectsLM.Text != "0" && passTwoConnectsLM.Text != "0")
            {
                passThreeConnectsLM.Text = (double.Parse(passThreeConnectsEntryLM.Text) - double.Parse(passTwoConnectsLM.Text) - double.Parse(passOneConnectsLM.Text)).ToString();
            }
        }
        #endregion

        #region Loss Mit Contact Entry Lost Focus Events
        private void passOneContactEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            passOneContactLM.Text = passOneContactEntryLM.Text;
        }

        private void passTwoContactEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneContactLM.Text != "0")
            {
                passTwoContactLM.Text = (double.Parse(passTwoContactLM.Text) - double.Parse(passOneContactLM.Text)).ToString();
            }
        }

        private void passThreeContactEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            if (passOneContactLM.Text != "0" && passTwoContactLM.Text != "0")
            {
                passThreeContactLM.Text = (double.Parse(passThreeContactLM.Text) - double.Parse(passTwoContactLM.Text) - double.Parse(passOneContactLM.Text)).ToString();
            }
        }
        #endregion

        #region Loss Mit Promise Entry Lost Focus Events
        private void passOnePromiseEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            double ptp = double.Parse(passOnePromiseEntryLM.Text);
            double rpcEntry = double.Parse(passOneContactEntryLM.Text);
            double rpcTotal = rpcEntry + ptp;
            double closure = ptp / rpcTotal;

            passOneContactLM.Text = rpcTotal.ToString();
            passOnePromiseLM.Text = ptp.ToString();
            passOneClosureLM.Text = Math.Round(closure * 100.0).ToString() + "%";
        }

        private void passTwoPromiseEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            double ptpEntry = double.Parse(passTwoPromiseEntryLM.Text);
            double rpcEntry = double.Parse(passTwoContactEntryLM.Text);
            double ptpPrev = double.Parse(passOnePromiseLM.Text);
            double rpcPrev = double.Parse(passOneContactLM.Text);
            double ptpTotal = ptpEntry - ptpPrev;
            double rpcTotal = rpcEntry + ptpEntry;
            double rpcCurrent = rpcTotal - rpcPrev;
            double closure = ptpTotal / rpcCurrent;

            passTwoContactLM.Text = rpcCurrent.ToString();
            passTwoPromiseLM.Text = ptpTotal.ToString();
            passTwoClosureLM.Text = Math.Round(closure * 100.0).ToString() + "%";
        }

        private void passThreePromiseEntryLM_LostFocus(object sender, RoutedEventArgs e)
        {
            double ptpEntry = double.Parse(passThreePromiseEntryLM.Text);
            double rpcEntry = double.Parse(passThreeContactEntryLM.Text);
            double ptpPrev = double.Parse(passTwoPromiseLM.Text) + double.Parse(passOnePromiseLM.Text);
            double rpcPrev = double.Parse(passTwoContactLM.Text) + double.Parse(passOneContactLM.Text);
            double ptpTotal = ptpEntry - ptpPrev;
            double rpcTotal = rpcEntry + ptpEntry;
            double rpcCurrent = rpcTotal - rpcPrev;
            double closure = ptpTotal / rpcCurrent;

            passThreeContactLM.Text = rpcCurrent.ToString();
            passThreePromiseLM.Text = ptpTotal.ToString();
            passThreeClosureLM.Text = Math.Round(closure * 100.0).ToString() + "%";
        }
        #endregion

        #region Main Window Events
        private void currentWindowSize_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Current Height:  " + ActualHeight.ToString() + "\r\n" + "Current Width: " + ActualWidth.ToString());
            MessageBox.Show("Expander Current Height:  " + allListExpander.ActualHeight.ToString() + "\r\n" + "Expander Current Width: " + allListExpander.ActualWidth.ToString());
        }
        #endregion


        private void SelectAllOnFocus_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBox box = (TextBox)sender;

            box.SelectAll();
        }

        #region Record Date Selected Date Changed Events
        private void recordDateES_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            DateTime recordDate = DateTime.Parse(recordDateES.ToString());

            CreateEarlyStageReadQuery(recordDate);
        }

        private void recordDateLM_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void recordDateLS_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void recordDateWE_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        #endregion

        #region StackPanel Button Events
        private void importData_Click(object sender, RoutedEventArgs e)
        {
            string querySql = "";

            List<string> importedData = new List<string>();

            if (departmentSelector.SelectedIndex == 2)
            {
                //Temporarily removed  - unknown usage
                //Get list of departments from Departments table in DialerTeam_be.accdb                       
                //querySql = "SELECT DepartmentName FROM Departments";

                ImportDataFromExcel();
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

            //importedData = ReadAccessDb(querySql);

        }

        private void saveData_Click(object sender, RoutedEventArgs e)
        {
            if (departmentSelector.SelectedIndex == 2)
            {
                DateTime targetDate = DateTime.Parse(recordDateES.ToString());

                string massCellString = "MASSCellRecords = @massCellRecords, MASSCellDials = @massCellDials, MASSCellSaturation = @massCellSaturation," +
                    " MASSCellPercentVolume = @massCellPercentVolume, MASSCellPassPenetration = @massCellPassPenetration, MASSCellTotalPassPenFactor =  @massCellTotalPassPenFactor";
                string miCellString = "MICellRecords = @miCellRecords, MICellDials = @miCellDials, MICellSaturation = @miCellSaturation," +
                    " MICellPercentVolume = @miCellPercentVolume, MICellPassPenetration = @miCellPassPenetration, MICellTotalPassPenFactor =  @miCellTotalPassPenFactor";
                string allString = "AllRecords = @allRecords, AllDials = @allDials, AllSaturation = @allSaturation," +
                    " AllPercentVolume = @allPercentVolume, AllPassPenetration = @allPassPenetrationOne, AllTotalPassPenFactor =  @allTotalPassPenFactorOne," +
                    " AllPassPenetrationTwo = @allPassPenetrationTwo, AllTotalPassPenFactorTwo =  @allTotalPassPenFactorTwo," +
                    " AllPassPenetrationThree = @allPassPenetrationThree, AllTotalPassPenFactorThree =  @allTotalPassPenFactorThree";
                string massString = "MASSRecords = @massRecords, MASSDials = @massDials, MASSSaturation = @massSaturation," +
                    " MASSPercentVolume = @massPercentVolume, MASSPassPenetration = @massPassPenetration, MASSTotalPassPenFactor =  @massTotalPassPenFactor";
                string nhString = "NHRecords = @nhRecords, NHDials = @nhDials, NHSaturation = @nhSaturation," +
                    " NHPercentVolume = @nhPercentVolume, NHPassPenetration = @nhPassPenetration, NHTotalPassPenFactor =  @nhTotalPassPenFactor";
                string nonContString = "NonContRecords = @nonContRecords, NonContDials = @nonContDials, NonContSaturation = @nonContSaturation," +
                    " NonContPercentVolume = @nonContPercentVolume, NonContPassPenetration = @nonContPassPenetration, NonContTotalPassPenFactor =  @nonContTotalPassPenFactor";
                string orString = "ORRecords = @orRecords, ORDials = @orDials, ORSaturation = @orSaturation," +
                    " ORPercentVolume = @orPercentVolume, ORPassPenetration = @orPassPenetration, ORTotalPassPenFactor =  @orTotalPassPenFactor";
                string combinedString = "CombinedRecords = @combinedRecords, CombinedDials = @combinedDials, CombinedSaturation = @combinedSaturation," +
                    " CombinedPassPenetration = @combinedPassPercent," +
                    " CombinedPassPenetrationTwo = @combinedPercentWorkedTwo," +
                    " CombinedPassPenetrationThree =  @combinedPercentWorkedThree";

                string query = "UPDATE ESDialsPerList SET " + massCellString + ", " + miCellString + ", " + allString + ", " + massString + ", " +
                               nhString + ", " + nonContString + ", " + orString + ", " + combinedString +
                               " WHERE CallDate = #" + targetDate.ToString() + "#";

                WriteEarlyStageData(query);
            }
            else if (departmentSelector.SelectedIndex == 3)
            {

            }
            else if (departmentSelector.SelectedIndex == 4)
            {

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

        private void sendWrap_Click(object sender, RoutedEventArgs e)
        {
            if (departmentSelector.SelectedIndex == 2)
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
        #endregion

        #region TabControl Events
        private void earlyStageTab_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            Keyboard.Focus(massCellRecords);
            massCellRecords.SelectAll();

            CreateEarlyStageReadQuery(DateTime.Today);
        }

        private void lateStageTab_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (passOneCallsLS.Text == "0")
            {
                Keyboard.Focus(passOneCallsEntryLS);
                passOneCallsEntryLS.SelectAll();
            }
            else if (passOneCallsLS.Text != "0" && passTwoCallsLS.Text == "0")
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
            if (passOneCallsLM.Text == "0")
            {
                Keyboard.Focus(passOneCallsEntryLM);
                passOneCallsEntryLM.SelectAll();
            }
            else if (passOneCallsLM.Text != "0" && passTwoCallsLM.Text == "0")
            {
                Keyboard.Focus(passTwoCallsEntryLM);
                passTwoCallsEntryLM.SelectAll();
            }
            else
            {
                Keyboard.Focus(passThreeCallsEntryLM);
                passThreeCallsEntryLM.SelectAll();
            }
        }
        #endregion



        


        
    }
}

