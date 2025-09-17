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

namespace PrintProgram
{
    static class SFISToJson
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
     
        

    
        public static DataTable reDt(string cmdtxt)
        {
            try
            {

                string sHttpURLRequest = cmdtxt;
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

        public static object reDt2( StringBuilder wo)
        {
            try
            {


                string sHttpURLRequest = "http://nportal.avalue.com.tw/PTD_CartonLabel/api/CartonLabel";
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.POST);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", wo, ParameterType.RequestBody);

                IRestResponse response = client.Execute(request);
                var jsonstr = JsonConvert.DeserializeObject(response.Content.ToString());

                //var response = client.Execute(request).Content;
                //DataTable dt = JsonConvert.DeserializeObject<DataTable>(response.Trim());
                return jsonstr;
            }
            catch (Exception ex)
            {

                return "";
            }
        }
        public static object  reDt3(StringBuilder wo)
        {
            try
            {
                string sHttpURLRequest = "http://nportal.avalue.com.tw/SFIS_MO/api/MO_Info";                
                var client = new RestClient(sHttpURLRequest);
                var request = new RestRequest(Method.POST);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", wo, ParameterType.RequestBody);

                IRestResponse response = client.Execute(request);
                var jsonstr = JsonConvert.DeserializeObject(response.Content.ToString());
               
                //var response = client.Execute(request).Content;
                //DataTable dt = JsonConvert.DeserializeObject<DataTable>(response.Trim());
                return jsonstr;
            }
            catch(Exception ex)
            {
                
                return "";
            }
        }


    }
}
