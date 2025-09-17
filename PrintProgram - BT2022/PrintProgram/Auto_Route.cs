using System;
using System.Text;
using RestSharp; // for REST API
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;


namespace PrintProgram
{
    /// <summary>
    /// AMES 類別用於儲存 API 回傳的結果資料。
    /// </summary>
    public class AMES
    {
        /// <summary>
        /// API 執行是否成功，"true" 表示成功。
        /// </summary>
        public string success { get; set; }

        /// <summary>
        /// API 回傳的訊息內容。
        /// </summary>
        public string msg { get; set; }

        /// <summary>
        /// API 回傳的狀態碼。
        /// </summary>
        public string status { get; set; }

        /// <summary>
        /// API 回傳的資料總數。
        /// </summary>
        public string dataTotal { get; set; }

        /// <summary>
        /// API 回傳的主要資料內容。
        /// </summary>
        public string data { get; set; }

        /// <summary>
        /// API 回傳的錯誤訊息。
        /// </summary>
        public string errors { get; set; }
    }

    /// <summary>
    /// WIPBox 類別用於儲存 WipBox API 回傳的結果資料。
    /// </summary>
    public class WIPBox
    {
        /// <summary>
        /// API 執行是否成功，"true" 表示成功。
        /// </summary>
        public string success { get; set; }

        /// <summary>
        /// API 回傳的訊息內容。
        /// </summary>
        public string msg { get; set; }

        /// <summary>
        /// API 回傳的狀態碼。
        /// </summary>
        public string status { get; set; }

        /// <summary>
        /// API 回傳的資料總數。
        /// </summary>
        public string dataTotal { get; set; }

        /// <summary>
        /// API 回傳的主要資料內容。
        /// </summary>
        public string data { get; set; }

        /// <summary>
        /// API 回傳的錯誤訊息。
        /// </summary>
        public string errors { get; set; }
    }

    /// <summary>
    /// ExtraNo 類別用於儲存條碼相關資訊。
    /// </summary>
    public class ExtraNo
    {
        /// <summary>
        /// 條碼編號。
        /// </summary>
        public string barcodeNo { get; set; }

        /// <summary>
        /// 條碼 ID。
        /// </summary>
        public string barcodeID { get; set; }
    }

    /// <summary>
    /// UserInfoes 類別用於儲存使用者相關資訊。
    /// </summary>
    public class UserInfoes
    {
        /// <summary>
        /// 使用者 ID。
        /// </summary>
        public string userID { get; set; }

        /// <summary>
        /// 使用者編號。
        /// </summary>
        public string userNo { get; set; }

        /// <summary>
        /// 使用者姓名。
        /// </summary>
        public string userName { get; set; }

        /// <summary>
        /// 部門 ID。
        /// </summary>
        public string deptID { get; set; }

        /// <summary>
        /// 登入帳號。
        /// </summary>
        public string loginNo { get; set; }

        /// <summary>
        /// 登入密碼。
        /// </summary>
        public string loginPassword { get; set; }

        /// <summary>
        /// 使用者電子郵件。
        /// </summary>
        public string userEMail { get; set; }
    }

    /// <summary>
    /// CZmomaterialList 類別用於儲存工單物料需求與實際發料資訊。
    /// </summary>
    public class CZmomaterialList
    {
        /// <summary>
        /// 工單編號。
        /// </summary>
        public string moID { get; set; }

        /// <summary>
        /// 物料編號。
        /// </summary>
        public string materialNo { get; set; }

        /// <summary>
        /// 需求數量。
        /// </summary>
        public string demandQty { get; set; }

        /// <summary>
        /// 實際發料數量。
        /// </summary>
        public string realsendQty { get; set; }
    }

    /// <summary>
    /// Auto_Route 類別提供工單過站、查詢工單資訊、查詢使用者資訊、工單物料需求、WipBox 操作等 API 相關功能。
    /// </summary>
    static class Auto_Route
    {
        /// <summary>
        /// 執行工單過站作業，將指定工單、單元、線別、使用者及序號資訊以 JSON 格式送至 AEMS API。
        /// </summary>
        /// <param name="WO">工單編號。</param>
        /// <param name="unitNo">單元編號。</param>
        /// <param name="line">線別編號。</param>
        /// <param name="userID">使用者 ID。</param>
        /// <param name="SN">序號（條碼）。</param>
        /// <returns>
        /// 若 API 回傳成功則回傳 "true"，否則回傳錯誤訊息或異常提示。
        /// </returns>
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
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 POST 請求
                var request = new RestRequest(Method.POST);
                // 加入 Content-Type 標頭，指定請求內容為 JSON 格式
                request.AddHeader("Content-Type", "application/json");
                // 將 JSON 字串加入請求主體
                request.AddParameter("application/json", sb, ParameterType.RequestBody);
                // 執行 API 請求
                IRestResponse response = client.Execute(request);
                #endregion

                // 判斷是否上傳成功
                if (response.StatusCode.ToString() == "OK")
                {
                    // 取得回傳內容
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


        /// <summary>
        /// 依據額外條碼編號 (ExtraNo) 取得對應的工單編號 (WIP)。
        /// </summary>
        /// <param name="SN">額外條碼編號。</param>
        /// <returns>
        /// 成功時回傳工單編號 (WIP)，失敗時回傳 "FAIL"。
        /// </returns>
        public static string ExtraNo_To_Wip(string SN)
        {
            try
            {
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/BarcodeInfoes/ByExtraNo/" + SN;
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 GET 請求
                var request = new RestRequest(Method.GET);
                // 執行 API 請求
                IRestResponse response = client.Execute(request);

                // 取得回傳內容
                Object test = response.Content.ToString();
                // 反序列化 JSON 並解析工單編號
                JObject jsonObj = JObject.Parse(test.ToString().Remove(0, 1).Replace("]", ""));
                // 取得 JSON 物件中的 getWipInfo.getWipAtt.wipNO 欄位
                JToken objToken = jsonObj.SelectToken("getWipInfo.getWipAtt.wipNO");
                // 回傳工單編號
                return objToken.ToString();
            }
            catch (Exception Ex)
            {
                return "FAIL";
            }
        }

        /// <summary>
        /// 依據工單編號 (WO) 取得對應的線別編號 (lineID) 與單元編號 (unitNO)。
        /// </summary>
        /// <param name="WO">工單編號。</param>
        /// <returns>
        /// 成功時回傳 "lineID&unitNO" 字串，失敗時回傳 "FAIL"。
        /// </returns>
        public static string Wip_To_Unint(string WO)
        {
            try
            {
                // 依據工單編號 (WO) 取得對應的線別與單元資訊
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/WipInfos/WipInfoByWipNo/" + WO;
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 GET 請求
                var request = new RestRequest(Method.GET);

                // 執行 API 請求
                IRestResponse response = client.Execute(request);
                // 取得回傳內容
                Object test = response.Content.ToString();

                // 解析回傳 JSON，取得 lineID 與 unitNO
                JObject jsonObj = JObject.Parse(test.ToString().Remove(0, 1).Replace("]", ""));
                // 取得 JSON 物件中的 lineID 欄位
                JToken objToken = jsonObj.SelectToken("lineID");
                // 取得 JSON 物件中的 unitNO 欄位
                JToken objToken2 = jsonObj.SelectToken("unitNO");
                // 將 lineID 與 unitNO 組合成字串，格式為 "lineID&unitNO"
                string lineIDunitNO = objToken.ToString() + "&" + objToken2.ToString();
                return lineIDunitNO;
            }
            catch (Exception Ex)
            {
                // 發生例外時回傳 "FAIL"
                return "FAIL";
            }
        }

        /// <summary>
        /// 依據工單編號 (WO) 取得對應的線別編號 (lineID) 與單元編號 (unitNO)（Eversun 系統）。
        /// </summary>
        /// <param name="WO">工單編號。</param>
        /// <returns>
        /// 成功時回傳 "lineID&unitNO" 字串，失敗時回傳 "FAIL"。
        /// </returns>
        public static string Wip_With_Eversun(string WO)
        {
            try
            {
                // 依據工單編號 (WO) 取得對應的線別與單元資訊
                string sHttpURLRequest = "http://192.168.4.109:5088/api/WipInfos/WipInfoByWipNo/" + WO;
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 GET 請求
                var request = new RestRequest(Method.GET);
                // 執行 API 請求
                IRestResponse response = client.Execute(request);
                // 取得回傳內容
                Object test = response.Content.ToString();
                // 解析回傳 JSON，取得 lineID 與 unitNO
                JObject jsonObj = JObject.Parse(test.ToString().Remove(0, 1).Replace("]", ""));
                // 取得 JSON 物件中的 lineID 欄位
                JToken objToken = jsonObj.SelectToken("lineID");
                // 取得 JSON 物件中的 unitNO 欄位
                JToken objToken2 = jsonObj.SelectToken("unitNO");
                // 將 lineID 與 unitNO 組合成字串，格式為 "lineID&unitNO"
                string lineIDunitNO = objToken.ToString() + "&" + objToken2.ToString();
                return lineIDunitNO;
            }
            catch (Exception Ex)
            {
                return "FAIL";
            }
        }

        /// <summary>
        /// 根據操作員編號 (OPID) 查詢並回傳對應的使用者 ID。
        /// </summary>
        /// <param name="OPID">操作員編號。</param>
        /// <returns>
        /// 成功時回傳使用者 ID，失敗時回傳 "FAIL"。
        /// </returns>
        public static string Check_OPID(string OPID)
        {
            try
            {
                // 依據操作員編號 (OPID) 查詢使用者資訊
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/UserInfoes/ByUserNo/" + OPID;
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 GET 請求
                var request = new RestRequest(Method.GET);

                // 執行 API 請求
                IRestResponse response = client.Execute(request);

                // 取得回傳內容
                Object test = response.Content.ToString();
                // 反序列化 JSON 並取得使用者資訊
                UserInfoes descJsonStu = JsonConvert.DeserializeObject<UserInfoes>(test.ToString());

                // 回傳使用者 ID
                return descJsonStu.userID;
            }
            catch (Exception Ex)
            {

                return "FAIL";
            }
        }

        /// <summary>
        /// 根據工單編號 (WIP) 取得對應的工單物料需求與實際發料資訊。
        /// </summary>
        /// <param name="WIP">工單編號。</param>
        /// <returns>
        /// 成功時回傳包含物料需求與發料資訊的 <see cref="DataTable"/>，失敗時回傳空的 <see cref="DataTable"/>。
        /// </returns>
        public static DataTable PowerCord(string WIP)
        {
            try
            {
                // 建立 API 請求網址，依據工單編號查詢物料需求與發料資訊
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/CZmomaterialList/ByMoID/" + WIP;
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 GET 請求
                var request = new RestRequest(Method.GET);
                // 執行 API 請求並取得回傳內容
                var response = client.Execute(request).Content;
                // 反序列化 JSON 資料為 DataTable
                DataTable dt = JsonConvert.DeserializeObject<DataTable>(response.Trim());
                return dt;
            }
            catch
            {
                // 發生例外時回傳空的 DataTable
                DataTable dt = new DataTable();
                return dt;
            }
        }

        /// <summary>
        /// 依據工單編號 (wono) 取得對應的其他條碼資訊。
        /// </summary>
        /// <param name="wono">工單編號。</param>
        /// <returns>
        /// 成功時回傳 API 回傳的 JSON 字串，失敗時丟出例外。
        /// </returns>
        public static string WipbarcodeOther(string wono)
        {
            try
            {
                // 建立 API 請求網址，依據工單編號查詢其他條碼資訊
                //192.168.4.109:5088 / api / WipBarcodeOther / WipNo / 104291801A01
                string sHttpURLRequest = "http://192.168.4.109:5088/api/WipBarcodeOther/WipNo/" + wono;
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 GET 請求
                var request = new RestRequest(Method.GET);
                // 執行 API 請求
                IRestResponse response = client.Execute(request);
                // 取得回傳內容
                var test = response.Content.ToString();

                // 回傳 API 回傳的 JSON 字串
                return test;
            }
            catch (Exception)
            {
                // 發生例外時直接丟出
                throw;
            }
        }

        /// <summary>
        /// 執行 WipBox API，將工單編號、額外條碼及箱子關閉狀態以 JSON 格式送至 AEMS API。
        /// </summary>
        /// <param name="WIP">工單編號。</param>
        /// <param name="SN">額外條碼。</param>
        /// <param name="CloseBox">箱子關閉狀態（"true" 或 "false"）。</param>
        /// <returns>
        /// 若 API 回傳成功則回傳 "true"，否則回傳錯誤訊息或異常提示。
        /// </returns>
        public static string WipBox(string WIP, string SN, string CloseBox)
        {
            // 執行 WipBox API，將工單編號、額外條碼及箱子關閉狀態以 JSON 格式送至 AEMS API。
            try
            {
                // 建立 JSON 字串
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("{");
                sb.AppendLine("\"wipNo\":\"" + WIP + "\",");
                sb.AppendLine("\"extraNo\":\"" + SN + "\",");
                sb.AppendLine("\"closeBox\":\"" + CloseBox + "\",");
                sb.AppendLine("}");

                // 設定 API 請求網址
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/WipBox/ByBeerAP";
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 POST 請求並加入 JSON 參數
                var request = new RestRequest(Method.POST);
                // 加入 Content-Type 標頭，指定請求內容為 JSON 格式
                request.AddHeader("Content-Type", "application/json");
                // 將 JSON 字串加入請求主體
                request.AddParameter("application/json", sb, ParameterType.RequestBody);

                // 執行 API 請求
                IRestResponse response = client.Execute(request);
                // 判斷回傳狀態
                if (response.StatusCode.ToString() == "OK")
                {
                    var test = response.Content.ToString();
                    // 反序列化 API 回傳結果
                    WIPBox descJsonStu2 = JsonConvert.DeserializeObject<WIPBox>(test.ToString());

                    // 判斷 API 執行是否成功
                    if (descJsonStu2.success == "true")
                    {
                        // 成功時回傳 "true"
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
                // 發生例外時回傳錯誤訊息
                return "程式異常，請重新開啟程式";
            }
        }

        /// <summary>
        /// 依據工單編號 (WipNo) 取得對應的工單看板資訊。
        /// </summary>
        /// <param name="WipNo">工單編號。</param>
        /// <returns>
        /// 成功時回傳 API 回傳的 JSON 字串，失敗時回傳 "FAIL"。
        /// </returns>
        public static string WipBoard(string WipNo)
        {
            try
            {
                // 建立 API 請求網址，依據工單編號查詢工單看板資訊
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/WipBoard/" + WipNo;
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 GET 請求
                var request = new RestRequest(Method.GET);
                // 執行 API 請求
                IRestResponse response = client.Execute(request);
                // 取得回傳內容
                var test = response.Content.ToString();
                // 回傳 API 回傳的 JSON 字串
                return test;
            }
            catch (Exception Ex)
            {
                // 發生例外時回傳 "FAIL"
                return "FAIL";
            }
        }

        /// <summary>
        /// 依據關聯工單編號 (WipNo) 取得對應的工單編號與料號 (itemNO)，僅回傳 unitNO 為 "P" 或 "O" 的資料。
        /// </summary>
        /// <param name="WipNo">關聯工單編號。</param>
        /// <returns>
        /// 成功時回傳 "wipNO&itemNO" 字串，若無符合條件則回傳空字串。
        /// </returns>
        public static string WipInfoByRelatedWoNo(string WipNo)
        {
            // 依據關聯工單編號 (WipNo) 取得對應的工單編號與料號 (itemNO)，僅回傳 unitNO 為 "P" 或 "O" 的資料。
            try
            {

                /// <summary>
                /// 儲存 API 回傳的工單編號與料號、單元編號的暫存字串。
                /// </summary>
                string getWipAtt = string.Empty;
                /// <summary>
                /// 儲存 JSON 解析後的工單編號 Token。
                /// </summary>
                JToken wipNOToken;
                /// <summary>
                /// 儲存 JSON 解析後的料號 Token。
                /// </summary>
                JToken itemNOToken;
                /// <summary>
                /// 儲存 JSON 解析後的單元編號 Token。
                /// </summary>
                JToken unitNOToken;

                /// WipNo 關聯工單編號
                // 建立 API 請求網址，依據關聯工單編號查詢工單資訊。
                string sHttpURLRequest = "http://192.168.4.109:5088/api/WipInfos/WipInfoByRelatedWoNo/" + WipNo;
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 透過 RestSharp 建立 GET 請求，取得 API 回傳內容。
                var request = new RestRequest(Method.GET);
                // 執行 API 請求
                IRestResponse response = client.Execute(request);
                // 取得回傳內容
                var test = response.Content.ToString();
                // 解析回傳的 JSON 陣列
                string json = test.ToString();
                // 將 JSON 字串轉換為 JArray 物件
                JArray jsonArray = JArray.Parse(json);
                // 逐筆解析 JSON 陣列
                foreach (JObject obj in jsonArray)
                {
                    wipNOToken = obj.SelectToken("getWipAtt.wipNO");
                    itemNOToken = obj.SelectToken("getWipAtt.itemNO");
                    unitNOToken = obj.SelectToken("unitNO");

                    // 只回傳 unitNO 為 "P" 或 "O" 的資料
                    if (unitNOToken.ToString() == "P" || unitNOToken.ToString() == "O")
                    {
                        getWipAtt = wipNOToken.ToString() + "&" + itemNOToken.ToString();
                    }
                }
                return getWipAtt;
            }
            catch (Exception ex)
            {
                // 發生例外時回傳空字串
                return "";
            }
        }

        /// <summary>
        /// 依據 Eversun 系統工單編號 (WipNo_Eve) 取得工單資訊。
        /// </summary>
        /// <param name="WipNo_Eve">Eversun 系統工單編號。</param>
        /// <returns>
        /// 成功時回傳 API 回傳的 JSON 字串，失敗時回傳 "FAIL"。
        /// </returns>
        public static string WipSystem_Eve(string WipNo_Eve)
        {
            try
            {
                // 建立 API 請求網址，依據 Eversun 系統工單編號查詢工單資訊
                string sHttpURLRequest = "http://192.168.4.109:5088/api/wipsystem/" + WipNo_Eve;
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 GET 請求
                var request = new RestRequest(Method.GET);
                // 執行 API 請求
                IRestResponse response = client.Execute(request);

                // 取得回傳內容
                var test = response.Content.ToString();

                // 回傳 API 回傳的 JSON 字串
                return test;
            }
            catch (Exception Ex)
            {
                // 發生例外時回傳 "FAIL"
                return "FAIL";
            }
        }

        /// <summary>
        /// 依據 Eversun 系統工單編號 (WipNo_Eve) 取得工單看板資訊。
        /// </summary>
        /// <param name="WipNo_Eve">Eversun 系統工單編號。</param>
        /// <returns>
        /// 成功時回傳 API 回傳的 JSON 字串，失敗時回傳 "FAIL"。
        /// </returns>
        public static string WipBoard_Eve(string WipNo_Eve)
        {
            try
            {
                // 建立 API 請求網址，依據 Eversun 系統工單編號查詢工單看板資訊
                string sHttpURLRequest = "http://192.168.4.109:5088/api/WipBoard/" + WipNo_Eve;
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 GET 請求
                var request = new RestRequest(Method.GET);
                // 執行 API 請求
                IRestResponse response = client.Execute(request);

                // 取得回傳內容
                var test = response.Content.ToString();

                // 回傳 API 回傳的 JSON 字串
                return test;
            }
            catch (Exception Ex)
            {
                // 發生例外時回傳 "FAIL"
                return "FAIL";
            }
        }

        /// <summary>
        /// 依據工單編號 (WipNo) 取得工單系統資訊。
        /// </summary>
        /// <param name="WipNo">工單編號。</param>
        /// <returns>
        /// 成功時回傳 API 回傳的 JSON 字串，失敗時回傳 "FAIL"。
        /// </returns>
        public static string WipSystem(string WipNo)
        {
            try
            {
                // 建立 API 請求網址，依據工單編號查詢工單系統資訊
                string sHttpURLRequest = "http://ames-prd.evalue-tech.com:5000/api/WipSystem/" + WipNo;
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 GET 請求
                var request = new RestRequest(Method.GET);
                // 執行 API 請求
                IRestResponse response = client.Execute(request);

                // 取得回傳內容
                var test = response.Content.ToString();

                // 回傳 API 回傳的 JSON 字串
                return test;
            }
            catch (Exception Ex)
            {
                // 發生例外時回傳 "FAIL"
                return "FAIL";
            }
        }

        /// <summary>
        /// 依據額外條碼編號 (SN) 取得對應的 WipBox 資料表。
        /// </summary>
        /// <param name="SN">額外條碼編號。</param>
        /// <returns>
        /// 成功時回傳 WipBox 資料表，失敗時回傳空的 <see cref="DataTable"/>。
        /// </returns>
        public static DataTable wipbox(string SN)
        {
            // 依據額外條碼編號查詢工單資訊
            try
            {
                // 建立 API 請求網址
                string sHttpURLRequest = "https://ames.avalue.com.tw:8443/api/BarcodeInfoes/ByExtraNo/" + SN;
                // 建立 REST API 客戶端
                var client = new RestClient(sHttpURLRequest);
                // 建立 GET 請求
                var request = new RestRequest(Method.GET);

                // 執行 API 請求
                IRestResponse response = client.Execute(request);

                // 取得回傳內容
                Object test = response.Content.ToString();

                // 從 JSON 物件中選取 wipID 欄位
                // 解析 JSON 取得 wipID
                JObject jsonObj = JObject.Parse(test.ToString().Remove(0, 1).Replace("]", ""));

                // 解析 JSON 取得 wipID
                JToken objToken = jsonObj.SelectToken("wipID"); // 從 JSON 物件中選取 wipID 欄位

                // 依據 wipID 查詢 WipBox 資料
                string sHttpURLRequest2 = "https://ames.avalue.com.tw:8443/api/WipBox/ByWipID/" + objToken.ToString();
                // 建立 REST API 客戶端
                var client2 = new RestClient(sHttpURLRequest2);
                // 建立 GET 請求
                var request2 = new RestRequest(Method.GET);
                // 執行 API 請求並取得回傳內容
                var response2 = client2.Execute(request2).Content;

                // 反序列化 JSON 為 DataTable
                DataTable dt = JsonConvert.DeserializeObject<DataTable>(response2.Trim());
                return dt;
            }
            catch
            {
                // 發生例外時回傳空的 DataTable
                DataTable dt = new DataTable();
                return dt;
            }
        }
    }
}
