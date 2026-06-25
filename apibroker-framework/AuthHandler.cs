using System;
using System.Web;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace NetDocumentsProxy.Framework
{
    public class AuthHandler : IHttpHandler
    {
        public bool IsReusable => false;

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "application/json";
            context.Response.AddHeader("Access-Control-Allow-Origin", "*");
            context.Response.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
            context.Response.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

            if (context.Request.HttpMethod == "OPTIONS")
            {
                context.Response.StatusCode = 200;
                return;
            }

            try
            {
                var path = context.Request.Url.AbsolutePath.ToLower();
                
                if (path.EndsWith("/authorize"))
                {
                    HandleAuthorize(context);
                }
                else if (path.EndsWith("/token"))
                {
                    HandleTokenAsync(context).Wait();
                }
                else
                {
                    context.Response.StatusCode = 404;
                    context.Response.Write("{\"error\":\"Not found\"}");
                }
            }
            catch (Exception ex)
            {
                context.Response.StatusCode = 500;
                context.Response.Write($"{{\"error\":\"{ex.Message}\"}}");
            }
        }

        private void HandleAuthorize(HttpContext context)
        {
            var clientId = context.Request.QueryString["clientId"];
            var redirectUri = context.Request.QueryString["redirectUri"];

            if (string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(redirectUri))
            {
                context.Response.StatusCode = 400;
                context.Response.Write("{\"error\":\"clientId and redirectUri are required\"}");
                return;
            }

            var authUrl = "https://api.netdocuments.com/v1/OAuth" +
                         $"?response_type=code" +
                         $"&client_id={Uri.EscapeDataString(clientId)}" +
                         $"&redirect_uri={Uri.EscapeDataString(redirectUri)}" +
                         $"&scope=read write" +
                         $"&state={Guid.NewGuid()}";

            var response = new { authorizationUrl = authUrl };
            context.Response.Write(JsonConvert.SerializeObject(response));
        }

        private async Task HandleTokenAsync(HttpContext context)
        {
            var requestBody = new System.IO.StreamReader(context.Request.InputStream).ReadToEnd();
            var tokenRequest = JsonConvert.DeserializeObject<TokenRequest>(requestBody);

            if (tokenRequest == null || string.IsNullOrEmpty(tokenRequest.Code))
            {
                context.Response.StatusCode = 400;
                context.Response.Write("{\"error\":\"code, clientId, and clientSecret are required\"}");
                return;
            }

            using (var httpClient = new HttpClient())
            {
                var formData = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("grant_type", "authorization_code"),
                    new KeyValuePair<string, string>("code", tokenRequest.Code),
                    new KeyValuePair<string, string>("client_id", tokenRequest.ClientId),
                    new KeyValuePair<string, string>("client_secret", tokenRequest.ClientSecret),
                    new KeyValuePair<string, string>("redirect_uri", tokenRequest.RedirectUri ?? "")
                });

                var response = await httpClient.PostAsync("https://api.netdocuments.com/v1/OAuth/Token", formData);
                var responseContent = await response.Content.ReadAsStringAsync();

                context.Response.StatusCode = (int)response.StatusCode;
                context.Response.Write(responseContent);
            }
        }
    }

    public class TokenRequest
    {
        public string Code { get; set; }
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string RedirectUri { get; set; }
    }
}