using System;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;

namespace Nexus.WebAPI.Handlers
{
    /// <summary>
    /// 
    /// </summary>
    public class AuthenticationServiceHandler
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="userId"></param>
        /// <param name="code"></param>
        /// <returns></returns>
        public static async Task<Boolean> PostProcessStep(string userId, string code)
        {
            using (var client = new HttpClient())
            {
                var issuer = ConfigurationManager.AppSettings["as:Issuer"];
                string baseUrl = issuer.TrimEnd(Convert.ToChar("/")) + "/";

                client.BaseAddress = new Uri(baseUrl);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                HttpResponseMessage response = await client.GetAsync("api/accounts/verifyusertoken?userId=" + userId + "&code=" + HttpUtility.UrlEncode(code));

                string responseString;

                if (response.IsSuccessStatusCode)
                {
                    responseString = await response.Content.ReadAsAsync<string>();
                }
                else
                {
                    responseString = "false";
                }

                return Convert.ToBoolean(responseString);

            }
        }
    }
}