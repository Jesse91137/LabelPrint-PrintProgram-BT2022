using System.Text.RegularExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RestSharp; // for REST API
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;


namespace PrintProgram
{
    public class AMES
    {
        public string success { get; set; }
        public string msg { get; set; }
        public string status { get; set; }
        public string dataTotal { get; set; }
        public string data { get; set; }

        public string errors { get; set; }

    }
    public class WIPBox
    {
        public string success { get; set; }
        public string msg { get; set; }
        public string status { get; set; }
        public string dataTotal { get; set; }
        public string data { get; set; }

        public string errors { get; set; }

    }

    public class ExtraNo
    {
        public string barcodeNo { get; set; }
        public string barcodeID { get; set; }
    }
    public class UserInfoes
    {
        public string userID { get; set; }
        public string userNo { get; set; }
        public string userName { get; set; }
        public string deptID { get; set; }
        public string loginNo { get; set; }
        public string loginPassword { get; set; }
        public string userEMail { get; set; }
    }
    public class CZmomaterialList
    {
        public string moID { get; set; }
        public string materialNo { get; set; }
        public string demandQty { get; set; }
        public string realsendQty { get; set; }
    }
    static class Auto_Route
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]

        public static string Route(string WO, string unitNo, string line, string userID, string SN)
        {
            try
            {
                #region JSON
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("{");
                sb.AppendLine("\"wipNo\":\"" + WO.Trim() + "\",");
                sb.AppendLine("\"wipID\":" + 0 + ",");
                sb.AppendLine("\"barcode\":\"" + SN.Trim() + "\",");
                sb.AppendLine("\"barcodeID\":" + 0 + ",");
                sb.AppendLine("\"barcodeType\":\"M\",");
                sb.AppendLine("\"ruleStatus\":\"P\",");
                sb.AppendLine("\"unitNo\":\"" + unitNo.Trim() + "\",");
                sb.AppendLine("\"flowRule\":" + 0 + ",");
                sb.AppendLine("\"stationID\":" + 1097 + ",");
                //sb.AppendLine("\"stationID\":" + 1116 + ",");
                sb.AppendLine("\"line\":" + line.Trim() + ",");
                sb.AppendLine("\"inputItems\":");
                sb.AppendLine("[");
                sb.AppendLine("{");
                sb.AppendLine("\"inputType\":\"\",");
                sb.AppendLine("\"inputData\":\"\",");
                sb.AppendLine("\"oldInputData\":\"\"");
                sb.AppendLine("}");
                sb.AppendLine("]" + ",");
                sb.AppendLine("\"outfits\":");
                sb.AppendLine("[");
                sb.AppendLine("{");
                sb.AppendLine("\"inputData\":\"" + "\"");
                sb.AppendLine("}");
                sb.AppendLine("]" + ",");
                sb.AppendLine("\"extNo\":\"" + SN.Trim() + "\",");
                sb.AppendLine("\"userID\":" + userID + "");
                sb.AppendLine("}");
                #endregion

                #region post JSON        
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/BarCodeCheck/PassIngByCheck";
                //string sHttpURLRequest = "http://nportal.avalue.com.tw/SWM_Xfis_tmp/api/AddTERecord?ProductSN=" + sn;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.POST);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", sb, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);

                #endregion
                //判斷是否上傳成功

                if (response.StatusCode.ToString() == "OK")
                {
                    var test = response.Content.ToString();
                    AMES descJsonStu = JsonConvert.DeserializeObject<AMES>(test.ToString());//反序列化
                    if (descJsonStu.success == "true")
                    {
                        return descJsonStu.success;
                    }
                    else
                    {
                        return descJsonStu.msg;
                    }
                }
                else
                {
                    return "過站失敗，AEMS API 連結異常";
                }
            }
            catch (Exception ex)
            {
                return "程式異常，請重新開啟程式";
            }
        }


        public static string ExtraNo_To_Wip(string SN)
        {
            try
            {
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/BarcodeInfoes/ByExtraNo/" + SN;
                //string sHttpURLRequest = cmdtxt;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.GET);

                IRestResponse response = client.Execute(request);

                Object test = response.Content.ToString();
                //var myobjList = JsonConvert.DeserializeObject<List<ExtraNo>>(test.ToString());//反序列化

                //var myobjList2 = JsonConvert.DeserializeObject<List<getWipInfo>>(test.ToString());

                JObject jsonObj = JObject.Parse(test.ToString().Remove(0, 1).Replace("]", ""));
                JToken objToken = jsonObj.SelectToken("getWipInfo.getWipAtt.wipNO");
                return objToken.ToString();


            }
            catch (Exception Ex)
            {

                return "FAIL";
            }
        }
        public static string Wip_To_Unint(string WO)
        {
            try
            {
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/WipInfos/WipInfoByWipNo/" + WO;
                //string sHttpURLRequest = cmdtxt;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.GET);

                IRestResponse response = client.Execute(request);

                Object test = response.Content.ToString();
                //var myobjList = JsonConvert.DeserializeObject<List<ExtraNo>>(test.ToString());//反序列化

                //var myobjList2 = JsonConvert.DeserializeObject<List<getWipInfo>>(test.ToString());

                JObject jsonObj = JObject.Parse(test.ToString().Remove(0, 1).Replace("]", ""));
                JToken objToken = jsonObj.SelectToken("lineID");
                JToken objToken2 = jsonObj.SelectToken("unitNO");
                string lineIDunitNO = objToken.ToString() + "&" + objToken2.ToString();
                return lineIDunitNO;
            }
            catch (Exception Ex)
            {
                return "FAIL";
            }
        }
        public static string Wip_With_Eversun(string WO)
        {
            try
            {
                string sHttpURLRequest = "http://192.168.4.109:5088/api/WipInfos/WipInfoByWipNo/" + WO;
                //string sHttpURLRequest = cmdtxt;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.GET);

                IRestResponse response = client.Execute(request);

                Object test = response.Content.ToString();
                //var myobjList = JsonConvert.DeserializeObject<List<ExtraNo>>(test.ToString());//反序列化

                //var myobjList2 = JsonConvert.DeserializeObject<List<getWipInfo>>(test.ToString());

                JObject jsonObj = JObject.Parse(test.ToString().Remove(0, 1).Replace("]", ""));
                JToken objToken = jsonObj.SelectToken("lineID");
                JToken objToken2 = jsonObj.SelectToken("unitNO");
                string lineIDunitNO = objToken.ToString() + "&" + objToken2.ToString();
                return lineIDunitNO;
            }
            catch (Exception Ex)
            {
                return "FAIL";
            }
        }
        public static string Check_OPID(string OPID)
        {
            try
            {
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/UserInfoes/ByUserNo/" + OPID;
                //string sHttpURLRequest = cmdtxt;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.GET);

                IRestResponse response = client.Execute(request);

                Object test = response.Content.ToString();
                UserInfoes descJsonStu = JsonConvert.DeserializeObject<UserInfoes>(test.ToString());//反序列化

                return descJsonStu.userID;
            }
            catch (Exception Ex)
            {

                return "FAIL";
            }
        }



        public static DataTable PowerCord(string WIP)
        {
            try
            {
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/CZmomaterialList/ByMoID/" + WIP;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.GET);
                var response = client.Execute(request).Content;
                DataTable dt = JsonConvert.DeserializeObject<DataTable>(response.Trim());
                return dt;
            }
            catch
            {
                DataTable dt = new DataTable();
                return dt;
            }
        }
        public static string WipbarcodeOther(string wono)
        {
            try
            {
                //192.168.4.109:5088 / api / WipBarcodeOther / WipNo / 104291801A01
                string sHttpURLRequest = "http://192.168.4.109:5088/api/WipBarcodeOther/WipNo/" + wono;
                //string sHttpURLRequest = cmdtxt;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.GET);
                IRestResponse response = client.Execute(request);

                var test = response.Content.ToString();

                return test;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static string WipBox(string WIP, string SN, string CloseBox)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("{");
                sb.AppendLine("\"wipNo\":\"" + WIP + "\",");
                sb.AppendLine("\"extraNo\":\"" + SN + "\",");
                sb.AppendLine("\"closeBox\":\"" + CloseBox + "\",");
                sb.AppendLine("}");

                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/WipBox/ByBeerAP";
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.POST);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", sb, ParameterType.RequestBody);

                IRestResponse response = client.Execute(request);
                if (response.StatusCode.ToString() == "OK")
                {
                    var test = response.Content.ToString();
                    WIPBox descJsonStu2 = JsonConvert.DeserializeObject<WIPBox>(test.ToString());//反序列化

                    if (descJsonStu2.success == "true")
                    {
                        return descJsonStu2.success;
                    }
                    else
                    {
                        return descJsonStu2.msg;
                    }
                }
                else
                {
                    return "AEMS API 連結異常";
                }
            }
            catch (Exception Ex)
            {

                return "程式異常，請重新開啟程式";
            }
        }

        public static string WipBoard(string WipNo)
        {
            try
            {
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/WipBoard/" + WipNo;
                //string sHttpURLRequest = cmdtxt;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.GET);
                IRestResponse response = client.Execute(request);

                var test = response.Content.ToString();

                return test;
            }
            catch (Exception Ex)
            {

                return "FAIL";
            }
        }
        public static string WipInfoByRelatedWoNo(string WipNo)
        {
            try
            {
                string getWipAtt = string.Empty;
                JToken wipNOToken;
                JToken itemNOToken;
                JToken unitNOToken;
                string sHttpURLRequest = "http://192.168.4.109:5088/api/WipInfos/WipInfoByRelatedWoNo/" + WipNo;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.GET);
                IRestResponse response = client.Execute(request);
                var test = response.Content.ToString();

                string json = test.ToString();                
                JArray jsonArray = JArray.Parse(json);
                foreach (JObject obj in jsonArray)
                {
                    wipNOToken = obj.SelectToken("getWipAtt.wipNO");
                    itemNOToken = obj.SelectToken("getWipAtt.itemNO");                    
                    unitNOToken = obj.SelectToken("unitNO");

                    if (unitNOToken.ToString() == "P" || unitNOToken.ToString() == "O")
                    {
                        getWipAtt = wipNOToken.ToString() + "&" + itemNOToken.ToString();
                    }


                    
                }

                //JToken objToken = jsonObj.SelectToken("getWipAtt.wipNO");
                //JToken objToken2 = jsonObj.SelectToken("getWipAtt.itemNO");                
                //string getWipAtt = objToken + "&" + objToken2;
                return getWipAtt;
            }
            catch (Exception ex)
            {
                return "";
            }
        }
        public static string WipSystem_Eve(string WipNo_Eve)
        {
            try
            {
                string sHttpURLRequest = "http://192.168.4.109:5088/api/wipsystem/" + WipNo_Eve;
                //string sHttpURLRequest = cmdtxt;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.GET);
                IRestResponse response = client.Execute(request);

                var test = response.Content.ToString();

                return test;
            }
            catch (Exception Ex)
            {
                return "FAIL";
            }
        }
        public static string WipBoard_Eve(string WipNo_Eve)
        {
            try
            {
                string sHttpURLRequest = "http://192.168.4.109:5088/api/WipBoard/" + WipNo_Eve;
                //string sHttpURLRequest = cmdtxt;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.GET);
                IRestResponse response = client.Execute(request);

                var test = response.Content.ToString();

                return test;
            }
            catch (Exception Ex)
            {
                return "FAIL";
            }
        }
        public static string WipSystem(string WipNo)
        {
            try
            {
                string sHttpURLRequest = "http://ames-prd.evalue-tech.com:5000/api/WipSystem/" + WipNo;
                //string sHttpURLRequest = cmdtxt;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.GET);
                IRestResponse response = client.Execute(request);

                var test = response.Content.ToString();
                
                return test;
            }
            catch (Exception Ex)
            {
                return "FAIL";
            }
        }

        public static DataTable wipbox(string SN)
        {
            try
            {
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/BarcodeInfoes/ByExtraNo/" + SN;
                //string sHttpURLRequest = cmdtxt;
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.GET);

                IRestResponse response = client.Execute(request);

                Object test = response.Content.ToString();
                //var myobjList = JsonConvert.DeserializeObject<List<ExtraNo>>(test.ToString());//反序列化

                //var myobjList2 = JsonConvert.DeserializeObject<List<getWipInfo>>(test.ToString());

                JObject jsonObj = JObject.Parse(test.ToString().Remove(0, 1).Replace("]", ""));
                JToken objToken = jsonObj.SelectToken("wipID");



                string sHttpURLRequest2 = "https://ames.avalue.com.tw:8443/api/WipBox/ByWipID/" + objToken.ToString();
                var client2 = new RestClient(sHttpURLRequest2);
                var request2 = new RestRequest(Method.GET);
                var response2 = client2.Execute(request2).Content;
                DataTable dt = JsonConvert.DeserializeObject<DataTable>(response2.Trim());
                return dt;
            }
            catch
            {
                DataTable dt = new DataTable();
                return dt;
            }
        }
    }
}
