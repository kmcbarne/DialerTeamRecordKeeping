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
        #region Microsoft Access ReadWrite Methods
        /// <summary>
        /// Opens Access database backendFilePath, runs SQL queryString, and closes database.
        /// </summary>
        /// <param name="queryString">SQL query string to run against database</param>
        public List<string> ReadAccessDb(string queryString)
        {
            List<string> departments = new List<string>();

            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source=" + backendFilePath + ";" +
                @"User Id=;Password=;";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (OleDbCommand command = new OleDbCommand(queryString, connection))
            {
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();

                    while (reader != null && reader.Read())
                    {
                        departments.Add(reader.GetString(0));
                    }
                    if (reader != null)
                        reader.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }

            return departments;
        }

        public bool ReadAccessDb(string queryString, bool hasRecords)
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source=" + backendFilePath + ";" +
                @"User Id=;Password=;";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (OleDbCommand command = new OleDbCommand(queryString, connection))
            {
                try
                {
                    connection.Open();

                    object obj = command.ExecuteScalar();
                    int count = obj is DBNull ? 0 : Convert.ToInt32(obj);

                    //This works, but if no records exist it uses the Exception system to halt further action.  Look for better approach.
                    if (count > 0)
                    {
                        hasRecords = true;
                    }
                }
                catch (Exception)
                {
                    // ignored
                }
            }

            return hasRecords;
        }

        public void ReadEarlyStageData(string query)
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source=" + backendFilePath + ";" +
                @"User Id=;Password=;";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (OleDbCommand command = new OleDbCommand(query, connection))
            {
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();

                    while (reader != null && reader.Read())
                    {
                        massCellRecords.Text = (reader["MASSCellRecords"].ToString());
                        massCellDials.Text = (reader["MASSCellDials"].ToString());
                        massCellSaturation.Text = FormatPercentCallables(double.Parse(reader["MASSCellSaturation"].ToString()));
                        massCellPercentTotal.Text = FormatPercentCallables(double.Parse(reader["MASSCellPercentVolume"].ToString()));
                        massCellPercentWorked.Text = (reader["MASSCellPassPenetration"].ToString());
                        massCellPassPercent.Text = FormatPassPercent(double.Parse(reader["MASSCellTotalPassPenFactor"].ToString()), 1);

                        miCellRecords.Text = (reader["MICellRecords"].ToString());
                        miCellDials.Text = (reader["MICellDials"].ToString());
                        miCellSaturation.Text = FormatPercentCallables(double.Parse(reader["MICellSaturation"].ToString()));
                        miCellPercentTotal.Text = FormatPercentCallables(double.Parse(reader["MICellPercentVolume"].ToString()));
                        miCellPercentWorked.Text = (reader["MICellPassPenetration"].ToString());
                        miCellPassPercent.Text = FormatPassPercent(double.Parse(reader["MICellTotalPassPenFactor"].ToString()), 1);

                        allRecordsOne.Text = (reader["AllRecords"].ToString());
                        allDialsOne.Text = (reader["AllDials"].ToString());
                        allSaturationOne.Text = FormatPercentCallables(double.Parse(reader["AllSaturation"].ToString()));
                        allPercentTotalOne.Text = FormatPercentCallables(double.Parse(reader["AllPercentVolume"].ToString()));
                        allPercentWorkedOne.Text = (reader["AllPassPenetration"].ToString());
                        allPassPercentOne.Text = FormatPassPercent(double.Parse(reader["AllTotalPassPenFactor"].ToString()), 1);
                        allPercentWorkedTwo.Text = (reader["AllPassPenetrationTwo"].ToString());
                        allPassPercentTwo.Text = FormatPassPercent(double.Parse(reader["AllTotalPassPenFactorTwo"].ToString()), 1);
                        allPercentWorkedThree.Text = (reader["AllPassPenetrationThree"].ToString());
                        allPassPercentThree.Text = FormatPassPercent(double.Parse(reader["AllTotalPassPenFactorThree"].ToString()), 1);

                        massRecords.Text = (reader["MASSRecords"].ToString());
                        massDials.Text = (reader["MASSDials"].ToString());
                        massSaturation.Text = FormatPercentCallables(double.Parse(reader["MASSSaturation"].ToString()));
                        massPercentTotal.Text = FormatPercentCallables(double.Parse(reader["MASSPercentVolume"].ToString()));
                        massPercentWorked.Text = (reader["MASSPassPenetration"].ToString());
                        massPassPercent.Text = FormatPassPercent(double.Parse(reader["MASSTotalPassPenFactor"].ToString()), 1);

                        nhRecords.Text = (reader["NHRecords"].ToString());
                        nhDials.Text = (reader["NHDials"].ToString());
                        nhSaturation.Text = FormatPercentCallables(double.Parse(reader["NHSaturation"].ToString()));
                        nhPercentTotal.Text = FormatPercentCallables(double.Parse(reader["NHPercentVolume"].ToString()));
                        nhPercentWorked.Text = (reader["NHPassPenetration"].ToString());
                        nhPassPercent.Text = FormatPassPercent(double.Parse(reader["NHTotalPassPenFactor"].ToString()), 1);

                        nonContRecords.Text = (reader["NonContRecords"].ToString());
                        nonContDials.Text = (reader["NonContDials"].ToString());
                        nonContSaturation.Text = FormatPercentCallables(double.Parse(reader["NonContSaturation"].ToString()));
                        nonContPercentTotal.Text = FormatPercentCallables(double.Parse(reader["NonContPercentVolume"].ToString()));
                        nonContPercentWorked.Text = (reader["NonContPassPenetration"].ToString());
                        nonContPassPercent.Text = FormatPassPercent(double.Parse(reader["NonContTotalPassPenFactor"].ToString()), 1);

                        orRecords.Text = (reader["ORRecords"].ToString());
                        orDials.Text = (reader["ORDials"].ToString());
                        orSaturation.Text = FormatPercentCallables(double.Parse(reader["ORSaturation"].ToString()));
                        orPercentTotal.Text = FormatPercentCallables(double.Parse(reader["ORPercentVolume"].ToString()));
                        orPercentWorked.Text = (reader["ORPassPenetration"].ToString());
                        orPassPercent.Text = FormatPassPercent(double.Parse(reader["ORTotalPassPenFactor"].ToString()), 1);

                        combinedRecords.Text = (reader["CombinedRecords"].ToString());
                        combinedDials.Text = (reader["CombinedDials"].ToString());
                        combinedSaturation.Text = FormatPercentCallables(double.Parse(reader["CombinedSaturation"].ToString()));
                        combinedPassPercent.Text = FormatPassPercent(double.Parse(reader["CombinedPassPenetration"].ToString()), 2);
                        //combinedPercentWorkedTwo.Text = FormatPassPercent(double.Parse(reader["CombinedPassPenetrationTwo"].ToString()), 2);
                        //combinedPercentWorkedThree.Text = FormatPassPercent(double.Parse(reader["CombinedPassPenetrationThree"].ToString()), 2);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        public void WriteAccessDb(string queryString)
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
               @"Data Source=" + backendFilePath + ";" +
               @"User Id=;Password=;";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (OleDbCommand command = new OleDbCommand(queryString, connection))
            {
                try
                {
                    connection.Open();
                    OleDbDataReader writer = command.ExecuteReader();

                    //while (writer.Read())
                    //{

                    //}
                    //writer.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        public void WriteEarlyStageData(string query)
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source=" + backendFilePath + ";" +
                @"User Id=;Password=;";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (OleDbCommand command = new OleDbCommand(query, connection))
            {
                try
                {
                    connection.Open();
                    command.Parameters.AddWithValue("@massCellRecords", int.Parse(massCellRecords.Text));
                    command.Parameters.AddWithValue("@massCellDials", int.Parse(massCellDials.Text));
                    command.Parameters.AddWithValue("@massCellSaturation", double.Parse(massCellSaturation.Text));
                    command.Parameters.AddWithValue("@massCellPercentVolume", double.Parse(massCellPercentTotal.Text.Replace("%", "")));
                    command.Parameters.AddWithValue("@massCellPassPenetration", double.Parse(massCellPercentWorked.Text));
                    command.Parameters.AddWithValue("@massCellTotalPassPenFactor", double.Parse(massCellPassPercent.Text));

                    command.Parameters.AddWithValue("@miCellRecords", int.Parse(miCellRecords.Text));
                    command.Parameters.AddWithValue("@miCellDials", int.Parse(miCellDials.Text));
                    command.Parameters.AddWithValue("@miCellSaturation", double.Parse(miCellSaturation.Text));
                    command.Parameters.AddWithValue("@miCellPercentVolume", double.Parse(miCellPercentTotal.Text.Replace("%", "")));
                    command.Parameters.AddWithValue("@miCellPassPenetration", double.Parse(miCellPercentWorked.Text));
                    command.Parameters.AddWithValue("@miCellTotalPassPenFactor", double.Parse(miCellPassPercent.Text));

                    command.Parameters.AddWithValue("@allRecords", int.Parse(allRecordsOne.Text));
                    command.Parameters.AddWithValue("@allDials", int.Parse(allDialsOne.Text));
                    command.Parameters.AddWithValue("@allSaturation", double.Parse(allSaturationOne.Text));
                    command.Parameters.AddWithValue("@allPercentVolume", double.Parse(allPercentTotalOne.Text.Replace("%", "")));
                    command.Parameters.AddWithValue("@allPassPenetrationOne", double.Parse(allPercentWorkedOne.Text));
                    command.Parameters.AddWithValue("@allTotalPassPenFactorOne", double.Parse(allPassPercentOne.Text));
                    command.Parameters.AddWithValue("@allPassPenetrationTwo", double.Parse(allPercentWorkedTwo.Text));
                    command.Parameters.AddWithValue("@allTotalPassPenFactorTwo", double.Parse(allPassPercentTwo.Text));
                    command.Parameters.AddWithValue("@allPassPenetrationThree", double.Parse(allPercentWorkedThree.Text));
                    command.Parameters.AddWithValue("@allTotalPassPenFactorThree", double.Parse(allPassPercentThree.Text));

                    command.Parameters.AddWithValue("@massRecords", int.Parse(massRecords.Text));
                    command.Parameters.AddWithValue("@massDials", int.Parse(massDials.Text));
                    command.Parameters.AddWithValue("@massSaturation", double.Parse(massSaturation.Text));
                    command.Parameters.AddWithValue("@massPercentVolume", double.Parse(massPercentTotal.Text.Replace("%", "")));
                    command.Parameters.AddWithValue("@massPassPenetration", double.Parse(massPercentWorked.Text));
                    command.Parameters.AddWithValue("@massTotalPassPenFactor", double.Parse(massPassPercent.Text));

                    command.Parameters.AddWithValue("@nhRecords", int.Parse(nhRecords.Text));
                    command.Parameters.AddWithValue("@nhDials", int.Parse(nhDials.Text));
                    command.Parameters.AddWithValue("@nhSaturation", double.Parse(nhSaturation.Text));
                    command.Parameters.AddWithValue("@nhPercentVolume", double.Parse(nhPercentTotal.Text.Replace("%", "")));
                    command.Parameters.AddWithValue("@nhPassPenetration", double.Parse(nhPercentWorked.Text));
                    command.Parameters.AddWithValue("@nhTotalPassPenFactor", double.Parse(nhPassPercent.Text));

                    command.Parameters.AddWithValue("@nonContRecords", int.Parse(nonContRecords.Text));
                    command.Parameters.AddWithValue("@nonContDials", int.Parse(nonContDials.Text));
                    command.Parameters.AddWithValue("@nonContSaturation", double.Parse(nonContSaturation.Text));
                    command.Parameters.AddWithValue("@nonContPercentVolume", double.Parse(nonContPercentTotal.Text.Replace("%", "")));
                    command.Parameters.AddWithValue("@nonContPassPenetration", double.Parse(nonContPercentWorked.Text));
                    command.Parameters.AddWithValue("@nonContTotalPassPenFactor", double.Parse(nonContPassPercent.Text));

                    command.Parameters.AddWithValue("@orRecords", int.Parse(orRecords.Text));
                    command.Parameters.AddWithValue("@orDials", int.Parse(orDials.Text));
                    command.Parameters.AddWithValue("@orSaturation", double.Parse(orSaturation.Text));
                    command.Parameters.AddWithValue("@orPercentVolume", double.Parse(orPercentTotal.Text.Replace("%", "")));
                    command.Parameters.AddWithValue("@orPassPenetration", double.Parse(orPercentWorked.Text));
                    command.Parameters.AddWithValue("@orTotalPassPenFactor", double.Parse(orPassPercent.Text));

                    command.Parameters.AddWithValue("@combinedRecords", int.Parse(combinedRecords.Text));
                    command.Parameters.AddWithValue("@combinedDials", int.Parse(combinedDials.Text));
                    command.Parameters.AddWithValue("@combinedSaturation", double.Parse(combinedSaturation.Text));
                    command.Parameters.AddWithValue("@combinedPassPercent", double.Parse(combinedPassPercent.Text));
                    //command.Parameters.AddWithValue("@combinedPercentWorkedTwo", double.Parse(combinedPercentWorkedTwo.Text));
                    //command.Parameters.AddWithValue("@combinedPercentWorkedThree", double.Parse(combinedPercentWorkedThree.Text));

                    command.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }        
        #endregion
    }
}
