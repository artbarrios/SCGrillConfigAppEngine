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
    class GrillSizesWebData
    {
        // global static vars
        private static HttpClient client = new HttpClient();

        // GET: api/GrillSizesData
        public static List<GrillSize> GetGrillSizes()
        {
            // return the data or perform an action using the remote webApiUrl
            string webApiPath = "api/GrillSizesData";
            string results = "";
            try
            {
                results = client.GetAsync(AppCommon.BuildUrl(AppCommon.GetRemoteWebApiUrl(), webApiPath)).Result.Content.ReadAsStringAsync().Result;
                return JsonConvert.DeserializeObject<List<GrillSize>>(results);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("GetGrillSizes: " + e.Message, e);
                throw new Exception(message);
            }
        } // GetGrillSizes
    }
}

