using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication2
{
    class GetTime
    {
        private string today, yestoday;

        public void setToday(string today) {
            this.today = today;
        }
        public void setYesToday(string yestoday)
        {
            this.yestoday = yestoday;

        }
        public string getDateToday() {
            //return today;
           return today ;

        }
        public string getDateYestoday() {
            // return yestoday;
            return yestoday ;
        }
    }
}
