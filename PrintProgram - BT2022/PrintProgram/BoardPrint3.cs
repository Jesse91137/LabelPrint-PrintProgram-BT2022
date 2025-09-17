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
    static class BoardPrint3
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
    
        public static bool PrintLabel(string LabelPath, List<string> SNList)
        {
            try
            {
                Engine engine = null;
                LabelFormatDocument btFormat = null;
                engine = new Engine();
                engine.Start();
                btFormat = engine.Documents.Open(LabelPath);

                for (int i = 0; i < SNList.Count; i++)
                {

                    string SN_Name = "SN" + (i + 1).ToString();
                    btFormat.SubStrings[SN_Name].Value = SNList[i].ToString().ToUpper(); //標籤檔中所設定的欄位名稱 。
                }
                btFormat.PrintSetup.IdenticalCopiesOfLabel = int.Parse("1"); //列印標籤數
                btFormat.Print();
                engine.Stop();


                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }        
    }
}
