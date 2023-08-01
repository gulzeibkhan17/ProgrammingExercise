using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgrammingExercise
{
    public class SubrecipientData
    {
        public string SubrecipientName { get; }
        public double Amount { get; set; }

        public SubrecipientData(string subrecipientName, double amount)
        {
            SubrecipientName = subrecipientName;
            Amount = amount;
        }
    }
}
