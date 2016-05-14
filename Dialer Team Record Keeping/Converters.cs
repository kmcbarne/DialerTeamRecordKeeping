using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace Dialer_Team_Record_Keeping
{
    public class AddListRecordsConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            double val = 0.0;
            double result = 0.0;

            try
            {
                foreach (object txt in values)
                {
                    if (double.TryParse(txt.ToString(), out val))
                    {
                        result += val;
                    }
                    else
                    {
                        return "0";
                    }
                }
                
                return result.ToString();
            }
            catch (Exception ex)
            {                
                MessageBox.Show(ex.ToString());
            }

            return result;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class SubtractListRecordsConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            double val = 0.0;
            double result = 0.0;
            double[] integerArray = new double[values.Count()];
            double[] sortedArray = new double[values.Count()];

            try
            {
                for (int i = 0; i < values.Count(); i++)
                {
                    if (double.TryParse(values[i].ToString(), out val))
                    {
                        integerArray[i] = val;
                    }
                    else
                    {
                        return "0";
                    }
                }

                sortedArray = integerArray.OrderBy(x => x).ToArray();
                sortedArray.Reverse();

                foreach (double num in sortedArray)
                {
                    result -= num;
                }

                return result.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return result;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        
    }

    public class AverageListRecordsConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            double val = 0.0;
            double result = 0.0;
            int count = 0;

            try
            {
                foreach (object txt in values)
                {
                    if (double.TryParse(txt.ToString(), out val))
                    {
                        result += val;

                        if (val != 0)
                        {
                            count++;
                        }
                    }
                    else
                    {
                        return "0";
                    }                    
                }

                if (result > 0)
                {
                    result /= count;
                }

                return result.ToString("0.0");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return result;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
