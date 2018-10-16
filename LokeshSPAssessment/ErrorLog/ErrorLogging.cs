using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ErrorLog
{
    public static class ErrorLogging
    {
       public static void ErrorlogWrite(Exception ex)
        {
            string DoubleSpace = "\n\n";
            try
            {
                string filepath = @"D:\ExceptionDetailsFile";  //Text File Path

                if (!Directory.Exists(filepath))
                {
                    Directory.CreateDirectory(filepath);

                }
                filepath = filepath + DateTime.Today.ToString("dd-MM-yy") + ".txt";   //Text File Name
                if (!File.Exists(filepath))
                {


                    File.Create(filepath).Dispose();

                }
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    string error = ex.StackTrace;
                    sw.WriteLine("-----------Exception Details on " + " " + DateTime.Now.ToString() + "-----------------");
                    sw.WriteLine("-------------------------------------------------------------------------------------");
                    sw.WriteLine(ex.Message);
                    sw.WriteLine(DoubleSpace);
                    sw.WriteLine(error);
                    sw.WriteLine("--------------------------------*End*------------------------------------------");
                    sw.WriteLine(DoubleSpace);
                    sw.Flush();
                    sw.Close();

                }

            }

        }
    }
}
//string error = "Log Written Date:" + " " + DateTime.Now.ToString() +
//                                    DoubleSpace + "Error Line No :" + " " + ErrorlineNo +
//                                    DoubleSpace + "Error Message:" + " " + Errormsg +
//                                    DoubleSpace + "Exception Type:" + " " + extype +
//                                    DoubleSpace + "Error Location :" + " " + ErrorLocation +
//                                    DoubleSpace + " Error Page Url:" + " " + exurl + DoubleSpace
//                                    + "User Host IP:" + " " + hostIp + DoubleSpace;