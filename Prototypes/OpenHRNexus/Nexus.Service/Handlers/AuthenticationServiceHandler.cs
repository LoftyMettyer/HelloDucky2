using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

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
         /// <param name="issuer"></param>
         /// <param name="userId"></param>
         /// <param name="purpose"></param>
         /// <returns></returns>
        private static async Task<string> GetServiceResult(string issuer, string url, Guid userId, Guid purpose)
        {
            using (var client = new HttpClient())
            {
                string baseUrl = issuer.TrimEnd(Convert.ToChar("/")) + "/";

                client.BaseAddress = new Uri(baseUrl);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                var serviceCall = string.Format("{0}/{1}/{2}", url, userId.ToString(), purpose.ToString());
                HttpResponseMessage response = await client.GetAsync(serviceCall);



                string responseString;

                if (response.IsSuccessStatusCode)
                {
                    responseString = await response.Content.ReadAsAsync<string>();
                }
                else
                {
                    responseString = "false";
                }

                return responseString;

            }
        }

        public static async Task<string> GetUserToken(string issuer, Guid userId, Guid purpose)
        {
            string result;

            try
            {
                result = await GetServiceResult(issuer, "api/accounts/getusertoken", userId, purpose);
            }

            catch
            {
                result = "";
            }

            return result;
        }

    }
}