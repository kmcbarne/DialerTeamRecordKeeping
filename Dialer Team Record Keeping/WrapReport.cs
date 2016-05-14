using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dialer_Team_Record_Keeping
{
    public class WrapReport
    {
        TimeSpan active;

        public string Tid { get; set; }

        public string Name { get; set; }

        public int Calls { get; set; }

        public TimeSpan Login { get; set; }

        public TimeSpan Active { get; set; }

        public TimeSpan NotReady { get; set; }

        public TimeSpan Idle { get; set; }

        public TimeSpan Wrap { get; set; }

        public TimeSpan Hold { get; set; }

        public int HoldCount { get; set; }

        public TimeSpan AvgWrap { get; set; }

        public WrapReport()
        {
            Tid = "";
            Name = "";
            Calls = 0;
            Login = new TimeSpan();
            active = new TimeSpan();
            NotReady = new TimeSpan();
            Idle = new TimeSpan();
            Wrap = new TimeSpan();
            Hold = new TimeSpan();
            HoldCount = 0;
            AvgWrap = new TimeSpan();
        }

        public WrapReport(string cvTid, string cvName, int cvCalls, TimeSpan cvLogin, TimeSpan cvActive, TimeSpan cvNotReady,
                          TimeSpan cvIdle, TimeSpan cvWrap, TimeSpan cvHold, int cvHoldCount, TimeSpan cvAvgWrap)
        {
            Tid = cvTid;
            Name = cvName;
            Calls = cvCalls;
            Login = cvLogin;
            active = cvActive;
            NotReady = cvNotReady;
            Idle = cvIdle;
            Wrap = cvWrap;
            Hold = cvHold;
            HoldCount = cvHoldCount;
            AvgWrap = cvAvgWrap;
        }

        public WrapReport(int cvCalls, TimeSpan cvWrap)
        {
            Calls = cvCalls;
            Wrap = cvWrap;
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
