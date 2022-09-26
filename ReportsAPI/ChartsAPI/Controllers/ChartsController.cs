namespace reports.Controllers

{
    using Newtonsoft.Json;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System.Xml;
    using System.Xml.Linq;
    using Models;
    using ChartsAPI.Services.ConfigurationProvider;

    /// <summary>
    /// 
    /// </summary>
    public class ChartsController : ApiController
    {

        // TODO: use IoC

        private string baseUrl = ConfigurationManager.AppSettings["ChartUrl"];
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="token"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("api/charts/views")]
        public async Task<HttpResponseMessage> GetViews()
        {
            try
            {

                var config = new ConfigProvider();

                //Authenticate the API
                string token = await this.GetAuthToken();

                string reportViewsURL = baseUrl + "/api/2.2/sites/0d8956fe-3e21-470c-a195-2da4b0ece5e4/views";
                string url = baseUrl + "/api/3.6/sites/0d8956fe-3e21-470c-a195-2da4b0ece5e4/views/{0}/image?resolution=high&maxAge={1}";
                var result = new HttpResponseMessage();

                var httpRequest = new HttpClient();
                httpRequest.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                var response = await httpRequest.GetAsync(reportViewsURL);
                var body = await response.Content.ReadAsStringAsync();
                XmlDocument xml = new XmlDocument();
                xml.LoadXml(body);
                var views = xml.GetElementsByTagName("view");
                List<ChartsModel> chartsModel = new List<ChartsModel>();
                for (int i = 0; i < views.Count; i++)
                {
                    var name = views.Item(i).Attributes.GetNamedItem("contentUrl").Value;
                    var id = views.Item(i).Attributes.GetNamedItem("id").Value;
                    chartsModel.Add(new ChartsModel()
                    {
                        Name = name,
                        ViewId = id,
                        Url = string.Format(url, id, config.ChartsMaxAge),
                        Sort = 0
                    });
                }
                var newReport = baseUrl + "/t/buildee/authoringNewWorkbook/1dv6l8v3t$ccl9-2f-d7-ub-d01ojz/Buildee_Analytics#4";
                chartsModel.Add(new ChartsModel()
                {
                    Name = "New Report",
                    ViewId = "",
                    Url = newReport,
                    Sort = 1
                });

                return Request.CreateResponse(HttpStatusCode.OK, chartsModel);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, ex.ToString());
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="token"></param>
        /// <param name="viewId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("api/charts/{viewId}/{buildingName}")]
        public async Task<HttpResponseMessage> GetChart(string viewId, string buildingName, string year = "", string utilType = "")
        {
            try
            {

                var config = new ConfigProvider();

                //Authenticate the API
                string token = await this.GetAuthToken();
                string url = baseUrl + "/api/3.6/sites/0d8956fe-3e21-470c-a195-2da4b0ece5e4/views/{0}/image?resolution=high&maxAge={1}&vf_BuildingId={2}"; //&vf_Year={2}&vf_Utility_Type={3}";

                if (!string.IsNullOrEmpty(year))
                {
                    url += "&vf_Year=" + year;
                }

                if (!string.IsNullOrEmpty(utilType))
                {
                    url += "&vf_Utility_Type=" + utilType;
                }

                var formattedUrl = string.Format(url, viewId, config.ChartsMaxAge, buildingName);

                var httpRequest = new HttpClient();
                httpRequest.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                var response = await httpRequest.GetAsync(formattedUrl);

                return response;
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, ex.ToString());
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="token"></param>
        /// <param name="viewId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("api/charts")]
        public async Task<HttpResponseMessage> GetChart()
        {
            try
            {
                string url = "";
                var reqParams = this.Request.GetQueryNameValuePairs();
                url = reqParams.FirstOrDefault(t => t.Key == "url").Value;
                foreach (var item in reqParams)
                {
                    if(item.Key != "url")
                    {
                        url = url + "&" + item.Key + "=" + item.Value;
                    }
                    
                }

                url = System.Net.WebUtility.HtmlDecode(url);

                string token = await this.GetAuthToken();
                var httpRequest = new HttpClient();
                httpRequest.DefaultRequestHeaders.CacheControl = new CacheControlHeaderValue() { NoCache = true };
                httpRequest.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                var response = await httpRequest.GetAsync(url);

                return response;
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, ex.ToString());
            }
        }

        /// <summary>
        /// Returns the auth token
        /// </summary>
        /// <returns></returns>
        private async Task<string> GetAuthToken()
        {
            string authURL = baseUrl + "/api/2.2/auth/signin";

            // Preparing the request body
            var dict = new Dictionary<string, string>();
            dict.Add("contentUrl", "buildee");
            //Get the token
            var values = new AuthRequest()
            {
                credentials = new CredetialsBody()
                {
                    name = ConfigurationManager.AppSettings["TableauUserName"],
                    password = ConfigurationManager.AppSettings["TableauPassword"],
                    site = dict
                }
            };

            var content = JsonConvert.SerializeObject(values);
            var httpContent = new StringContent(content, Encoding.UTF8, "application/json");
            var authRequest = new HttpClient();
            var tokenResponse = await authRequest.PostAsync(authURL, httpContent);
            var responseBody = await tokenResponse.Content.ReadAsStringAsync();
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(responseBody);
            var crendetialsTag = xml.GetElementsByTagName("credentials");
            var token = crendetialsTag[0].Attributes.GetNamedItem("token").Value;

            return token;
        }
    }
}