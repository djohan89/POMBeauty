using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace OPM.Enginee
{
    class LogHelper
    {
        public static void Writelog(string strlogs)
        {
            string strDireactory = @"F:\";
            string strLogFile = "Opmlog.txt";
            //Write file
            using (StreamWriter w = File.AppendText(strDireactory + strLogFile))
            {
                w.Write("\r\n Log Entry \n");
                w.WriteLine("{0} {1} \n", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                w.WriteLine(strlogs);
                w.WriteLine("\n");
                w.WriteLine("..............................");
            }

        }
        
    }
}
