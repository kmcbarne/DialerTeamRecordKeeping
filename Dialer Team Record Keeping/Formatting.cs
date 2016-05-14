using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dialer_Team_Record_Keeping
{
    public partial class MainWindow : System.Windows.Window
    {
        public string FormatPassPercent(double decimalPercent, int decimalPlaces)
        {
            string format = "";

            if (decimalPlaces == 1)
            {
                format = "0.0";
            }
            else if (decimalPlaces == 2)
            {
                format = "0.00";
            }

            return Math.Round(decimalPercent, decimalPlaces).ToString(format);
        }

        public string FormatPercentCallables(double decimalPercent)
        {
            return (Math.Round((decimalPercent * 100), 0).ToString());
        }

        
    }
}
