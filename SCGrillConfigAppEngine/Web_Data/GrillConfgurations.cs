using SCGrillConfig.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace SCGrillConfigAppEngine.Web_Data
{
    class GrillConfgurationsWebData
    {
        // global static vars
        private static HttpClient client = new HttpClient();

        // GET: api/GrillConfgurationsData
        public static List<GrillConfguration> GetGrillConfgurations()
        {
            // return the data or perform an action using the remote webApiUrl
            string webApiPath = "api/GrillConfgurationsData";
            string results = "";
            try
            {
                results = client.GetAsync(AppCommon.BuildUrl(AppCommon.GetRemoteWebApiUrl(), webApiPath)).Result.Content.ReadAsStringAsync().Result;
                return JsonConvert.DeserializeObject<List<GrillConfguration>>(results);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("GetGrillConfgurations: " + e.Message, e);
                throw new Exception(message);
            }
        } // GetGrillConfgurations
    }
}

