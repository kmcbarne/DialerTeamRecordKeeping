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

        public List<TextBox> CreateSaturationList()
        {
            List<TextBox> saturationList = new List<TextBox>();

            saturationList.Add(massCellSaturation);
            saturationList.Add(miCellSaturation);
            saturationList.Add(allSaturationOne);
            saturationList.Add(allSaturationTwo);
            saturationList.Add(allSaturationThree);
            saturationList.Add(massSaturation);
            saturationList.Add(nhSaturation);
            saturationList.Add(nonContSaturation);
            saturationList.Add(orSaturation);

            //if (int.Parse(allDialsTwo.Text) > 0)
            //{
            //    saturationList.Add(int.Parse(allSaturationTwo.Text));

            //    if (int.Parse(allDialsThree.Text) > 0)
            //    {
            //        saturationList.Add(int.Parse(allSaturationThree.Text));
            //    }
            //}

            return saturationList;

            //totalRecords = Int32.Parse(massCellRecords.Text) + Int32.Parse(miCellRecords.Text) + Int32.Parse(allRecords.Text) + Int32.Parse(massRecords.Text) +Int32.Parse(nhRecords.Text) + Int32.Parse(nonContRecords.Text) + Int32.Parse(orRecords.Text);
        }

        //Temporary bool, fix later to match with CreateSaturation number of elements
        public List<double> CreateRecordsList(bool isForSaturation)
        {
            List<double> recordsList = new List<double>();

            recordsList.Add(int.Parse(massCellRecords.Text));
            recordsList.Add(int.Parse(miCellRecords.Text));
            recordsList.Add(int.Parse(allRecordsOne.Text));
            recordsList.Add(int.Parse(allRecordsTwo.Text));
            recordsList.Add(int.Parse(allRecordsThree.Text));
            recordsList.Add(int.Parse(massRecords.Text));
            recordsList.Add(int.Parse(nhRecords.Text));
            recordsList.Add(int.Parse(nonContRecords.Text));
            recordsList.Add(int.Parse(orRecords.Text));

            if (int.Parse(allDialsTwo.Text) > 0)
            {
                recordsList.Add(int.Parse(allRecordsTwo.Text));

                if (int.Parse(allDialsThree.Text) > 0)
                {
                    recordsList.Add(int.Parse(allRecordsThree.Text));
                }
            }

            return recordsList;

            //totalRecords = Int32.Parse(massCellRecords.Text) + Int32.Parse(miCellRecords.Text) + Int32.Parse(allRecords.Text) + Int32.Parse(massRecords.Text) +Int32.Parse(nhRecords.Text) + Int32.Parse(nonContRecords.Text) + Int32.Parse(orRecords.Text);
        }

        public List<double> CreateRecordsList()
        {
            List<double> recordsList = new List<double>();

            recordsList.Add(double.Parse(massCellRecords.Text));
            recordsList.Add(double.Parse(miCellRecords.Text));
            recordsList.Add(double.Parse(allRecordsOne.Text));
            recordsList.Add(double.Parse(massRecords.Text));
            recordsList.Add(double.Parse(nhRecords.Text));
            recordsList.Add(double.Parse(nonContRecords.Text));
            recordsList.Add(double.Parse(orRecords.Text));

            if (double.Parse(allDialsTwo.Text) > 0)
            {
                recordsList.Add(double.Parse(allRecordsTwo.Text));

                if (double.Parse(allDialsThree.Text) > 0)
                {
                    recordsList.Add(double.Parse(allRecordsThree.Text));
                }
            }

            return recordsList;

            //totalRecords = Int32.Parse(massCellRecords.Text) + Int32.Parse(miCellRecords.Text) + Int32.Parse(allRecords.Text) + Int32.Parse(massRecords.Text) +Int32.Parse(nhRecords.Text) + Int32.Parse(nonContRecords.Text) + Int32.Parse(orRecords.Text);
        }

        public List<int> CreateDialsList()
        {
            List<int> dialsList = new List<int>();

            dialsList.Add(int.Parse(massCellDials.Text));
            dialsList.Add(int.Parse(miCellDials.Text));
            dialsList.Add(int.Parse(allDialsOne.Text));
            dialsList.Add(int.Parse(allDialsTwo.Text));
            dialsList.Add(int.Parse(allDialsThree.Text));
            dialsList.Add(int.Parse(massDials.Text));
            dialsList.Add(int.Parse(nhDials.Text));
            dialsList.Add(int.Parse(nonContDials.Text));
            dialsList.Add(int.Parse(orDials.Text));

            //if (int.Parse(allDialsTwo.Text) > 0)
            //{
            //    dialsList.Add(int.Parse(allDialsTwo.Text));

            //    if (int.Parse(allDialsThree.Text) > 0)
            //    {
            //        dialsList.Add(int.Parse(allDialsThree.Text));
            //    }
            //}

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

        public List<TextBox> CreatePercentWorkedList()
        {
            List<TextBox> percentWorkedList = new List<TextBox>();

            percentWorkedList.Add(massCellPercentWorked);
            percentWorkedList.Add(miCellPercentWorked);
            percentWorkedList.Add(allPercentWorkedOne);
            percentWorkedList.Add(allPercentWorkedTwo);
            percentWorkedList.Add(allPercentWorkedThree);
            percentWorkedList.Add(massPercentWorked);
            percentWorkedList.Add(nhPercentWorked);
            percentWorkedList.Add(nonContPercentWorked);
            percentWorkedList.Add(orPercentWorked);

            return percentWorkedList;

            //totalRecords = Int32.Parse(massCellRecords.Text) + Int32.Parse(miCellRecords.Text) + Int32.Parse(allRecords.Text) + Int32.Parse(massRecords.Text) +Int32.Parse(nhRecords.Text) + Int32.Parse(nonContRecords.Text) + Int32.Parse(orRecords.Text);
        }

        public List<TextBox> CreatePassPenetrationList()
        {
            List<TextBox> passPenetrationList = new List<TextBox>();

            passPenetrationList.Add(massCellPassPercent);
            passPenetrationList.Add(miCellPassPercent);
            passPenetrationList.Add(allPassPercentOne);
            passPenetrationList.Add(allPassPercentTwo);
            passPenetrationList.Add(allPassPercentThree);
            passPenetrationList.Add(nhPassPercent);
            passPenetrationList.Add(nonContPassPercent);
            passPenetrationList.Add(orPassPercent);

            return passPenetrationList;
        }
        
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

        /// <summary>
        /// Calculates Percent Volume using formula records divided by totalRecords rounded to 2 decimal places.
        /// </summary>
        /// <param name="records">Number of records in the list.</param>
        /// <param name="totalRecords">Combined number of records in all lists.</param>
        /// <returns>Returns the result of records divided by totalRecords rounded to 2 decimal places.</returns>
        private double CalculatePercentVolume(double records, double totalRecords)
        {
            double percentVolume = 0.0;

            percentVolume = records / totalRecords;

            return Math.Round(percentVolume, 2);
        }

        /// <summary>
        /// Calculates Percent Volume using formula records divided by totalRecords.
        /// </summary>
        /// <param name="records">Number of records in the list.</param>
        /// <param name="totalRecords">Combined number of records in all lists.</param>
        /// <returns>Returns the result of records divided by totalRecords without rounding.</returns>
        private double CalculatePercentVolumeRaw(double records, double totalRecords)
        {
            return records / totalRecords;
        }

        public void CalculatePercentVolume(List<TextBox> volumes)
        {
            double totalRecords = CalculateTotalRecords(CreateRecordsList());

            List<double> recordsList = CreateRecordsList();
            List<TextBox> percentVolumeList = CreatePercentVolumeList();

            int i = 0;

            foreach (TextBox percentVolume in percentVolumeList)
            {
                percentVolume.Text = Math.Round(recordsList.ElementAt(i) / totalRecords * 100.0).ToString();

                i++;
            }
        }

        public List<double> CalculatePassPercent()
        {
            List<TextBox> passPenetration = CreatePassPenetrationList();
            List<TextBox> percentVolume = CreatePercentVolumeList();
            List<double> passPercent = new List<double>();

            for (int i = 0; i < passPenetration.Count; i++)
            {
                passPercent.Add(CalculatePassPercent(double.Parse(passPenetration.ElementAt(i).Text),double.Parse(percentVolume.ElementAt(i).Text)));
            }

            return passPercent;
        }

        public double CalculateTotalRecords(List<double> records)
        {
            double totalRecords = 0;

            foreach (double record in records)
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

        public double CalculatePassPercent(double percentVolume, double passPenetration)
        {
            return percentVolume * passPenetration;
        }

        public void DisplayPassPercent(double percentVolume, double passPenetration, TextBox passPercentTextBox)
        {
            if (percentVolume > 0 && passPenetration > 0)
            {
                passPercentTextBox.Text = Math.Round(CalculatePassPercent(percentVolume, passPenetration), 1).ToString("0.0");
            }

            //UpdateCombinedPassPercent();
        }

        //public void UpdateCombinedPassPercent()
        //{
        //    List<TextBox> passPenetrationBoxes = CreatePassPenetrationList();
        //    double result = 0.0;

        //    foreach (TextBox box in passPenetrationBoxes)
        //    {
        //        result += double.Parse(box.Text);
        //    }

        //    combinedPassPercent.Text = result.ToString("0.0");

        //    //combinedPercentWorkedTwo.Text = double.Parse(allPassPercentTwo.Text).ToString("0.00");
        //    //combinedPercentWorkedThree.Text = double.Parse(allPassPercentThree.Text).ToString("0.00");
        //}

        private void allListExpander_OnCollapsed(object sender, RoutedEventArgs e)
        {
            this.Width -= 420;
        }

        private void AllListExpander_OnExpanded(object sender, RoutedEventArgs e)
        {
            this.Width += 420;
        }
    }
}
