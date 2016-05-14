using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dialer_Team_Record_Keeping
{
    public class ListProgressData
    {
        private double callsDialed;
        private double transferredToAgent;
        private double ptp;
        private double rpc;
        private double records;
        private double penetration;

        public double CallsDialed
        {
            get { return callsDialed; }
            set
            {
                if (value < 0)
                    value = 0;
                callsDialed = value;
            }
        }

        public double TransferredToAgent
        {
            get { return transferredToAgent; }
            set
            {
                if (value < 0)
                    value = 0;
                transferredToAgent = value;
            }
        }

        public double PTP
        {
            get { return ptp; }
            set
            {
                if (value < 0)
                    value = 0;
                ptp = value;
            }
        }

        public double RPC
        {
            get { return rpc; }
            set
            {
                if (value < 0)
                    value = 0;
                rpc = value;
            }
        }

        public double Records
        {
            get { return records; }
            set
            {
                if (value < 0)
                    value = 0;
                records = value;
            }
        }

        public double Penetration
        {
            get { return penetration; }
            set
            {
                if (value < 0)
                    value = 0;
                penetration = value;
            }
        }
    }
}
