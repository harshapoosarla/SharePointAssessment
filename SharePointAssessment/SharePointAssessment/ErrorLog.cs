using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointAssessment
{
    public class ErrorLog
    {
        public static void Errorlog(Exception e)
        {
            string Message = "--------" + DateTime.Now + Environment.NewLine + "--------" + e.StackTrace +
                Environment.NewLine + "--------" + e.Message + "--------" + Environment.NewLine + Environment.NewLine;
            string Path = @"D:\harsha853\GitHub\SharePointAssessment\SharePointAssessment\SharePointAssessment\Errors.txt";
            File.AppendAllText(Path, Message);
        }
    }

}

