using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using Microsoft.SqlServer.Server;
using System.Collections;
using RestSharp; // for REST API
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text;
using Seagull.BarTender.Print;

namespace PrintProgram
{
    static class AccPrint
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
     
        

    
        public static bool PrintLabel(string LabelPath, string WO, string PN,  string QTY, string No_Number)
        {
            try
            {
                Engine engine = null;
                LabelFormatDocument btFormat = null;

                engine = new Engine();
                engine.Start();
                btFormat = engine.Documents.Open(LabelPath);

                btFormat.SubStrings["WO"].Value = WO;
                btFormat.SubStrings["PN"].Value = PN;
                btFormat.SubStrings["QTY"].Value = QTY;
                btFormat.SubStrings["No_Number"].Value = No_Number;
               
                btFormat.PrintSetup.IdenticalCopiesOfLabel = int.Parse("1"); //列印標籤數
                btFormat.Print();
                engine.Stop();
                return true;
            }
            catch
            {
                return false;
            }
        }

        


    }
}
