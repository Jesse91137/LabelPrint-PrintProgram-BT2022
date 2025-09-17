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
using static PrintProgram.RadForm1;

namespace PrintProgram
{
    static class OemPrint2
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]        
        public static bool PrintLabel(string LabelPath, string WO,string QTY,string Bios, List<string> SNList)
        {
            try
            {
                string PN="", No_Number="", ENGSR="",PC="";
                #region Avalue API
                /*
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("{");
                sb.AppendLine("\"Key\":\"" + "@Avalue.ZMO.SOP" + "\",");
                sb.AppendLine("\"moid\":\"" + WO.Trim() + "\",");
                sb.AppendLine("}");

                var test = SFISToJson.reDt3(sb);
                SFIS descJsonStu = JsonConvert.DeserializeObject<SFIS>(test.ToString());//反序列化
                if (!string.IsNullOrEmpty(descJsonStu.powercord))
                {
                    PC = descJsonStu.powercord;
                }
                if (!string.IsNullOrEmpty(descJsonStu.ProductID_MF))
                {
                    ENGSR = descJsonStu.ProductID_MF;
                }
                if (!string.IsNullOrEmpty(descJsonStu.productid))
                {
                    PN = descJsonStu.productid;
                }
                */
                #endregion
                //取得工程編號
                string sqlCmd = "select top(1)Eng_SR from Print_CustomWONO_Table where Avalue_WoNo = '" + WO.Trim() + "' order by sno desc";
                DataSet ds = db.reDs(sqlCmd);
                ENGSR = ds.Tables[0].Rows[0][0].ToString();
                //取得powercode
                DataTable data = new DataTable();
                data = Auto_Route.PowerCord(WO.Trim());
                if (data != null && data.Rows.Count > 0)
                {
                    if (data.Rows.Count == 1)
                    {
                        if (int.Parse(data.Rows[0]["realsendQty"].ToString()) > 0)
                        {
                            PC = data.Rows[0]["materialNo"].ToString();
                        }
                    }
                    else
                    {
                        for (int i = 0; i < data.Rows.Count; i++)
                        {
                            if (int.Parse(data.Rows[i]["realsendQty"].ToString()) > 0)
                            {
                                PC += data.Rows[i]["materialNo"].ToString() + ",";
                            }
                        }
                    }
                }
                //取得機種名稱 e.g. RAD-SYS02-ODIN-C2R
                SFIS descJsonStu = JsonConvert.DeserializeObject<SFIS>(Auto_Route.WipSystem(WO.Trim()));//反序列化
                PN = descJsonStu.itemNo;

                Engine engine = null;
                LabelFormatDocument btFormat = null;
                engine = new Engine();
                engine.Start();
                btFormat = engine.Documents.Open(LabelPath);

                btFormat.SubStrings["WO"].Value = WO;
                btFormat.SubStrings["PN"].Value = PN;
                btFormat.SubStrings["Bios"].Value = Bios;
                //btFormat.SubStrings["PC"].Value = PC;
                btFormat.SubStrings["QTY"].Value = QTY;
                btFormat.SubStrings["No_Number"].Value = No_Number;
                btFormat.SubStrings["ENGSR"].Value = "(" + ENGSR + ")";


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
            catch (Exception ex)
            {
                return false;
            }
        }

        public static bool PrintLabel2(string LabelPath, List<string> SNList)
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
                    btFormat.SubStrings[SN_Name].Value = SNList[i].ToString(); //標籤檔中所設定的欄位名稱 。
                }
                btFormat.PrintSetup.IdenticalCopiesOfLabel = int.Parse("1"); //列印標籤數
                btFormat.Print();
                engine.Stop();


                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }


    }
}
