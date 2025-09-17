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
    static class PhilEPC_Print
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        public static bool PrintLabel(string mod,string labelType, string LabelPath,string ECN,string BIOSVer,string BIOSCS, string SN,string MAC)
        {
            //partNumber, choiceRdb, printlabe, ecn, biosVer, bioscs,radTxt_SN.Text, radTxt_Mac.Text
            try
            {
                Engine engine = null;
                LabelFormatDocument btFormat = null;
                engine = new Engine();
                engine.Start();
                btFormat = engine.Documents.Open(LabelPath);
                //switch (mod)
                //{
                //    case "01R":
                //        btFormat = engine.Documents.Open(LabelPath);
                //        break;
                //    case "02R":
                //        btFormat = engine.Documents.Open(LabelPath);
                //        break;
                //    case "03R":
                //        btFormat = engine.Documents.Open(LabelPath);
                //        break;
                //    case "04R":
                //        btFormat = engine.Documents.Open(LabelPath);
                //        break;
                //    default:
                //        break;
                //}

                if (labelType=="Big")
                {
                    btFormat.SubStrings["SN"].Value = SN.ToUpper();
                    btFormat.SubStrings["MAC"].Value = MAC.ToUpper();
                    btFormat.SubStrings["BIOS"].Value = BIOSVer + "  " + BIOSCS;
                    btFormat.SubStrings["ECN"].Value = ECN;
                }
                else
                {
                    btFormat.SubStrings["MAC"].Value = MAC.ToUpper();
                }
                
                //for (int i = 0; i < 5; i++)
                //{
                //    string SN_Name = "SN" + (i + 1).ToString();
                //    btFormat.SubStrings[SN_Name].Value = SNList[i].ToString(); //標籤檔中所設定的欄位名稱 。
                //}
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
