using System.Xml;

namespace Yukarisan
{
    internal class SlackAPIHttpClient
    {
        private readonly string baseUrl;
        private readonly HttpClient httpClient;
        private string TOKEN = String.Empty;

        public SlackAPIHttpClient(string baseUrl)
        {
            this.baseUrl = baseUrl;
            this.httpClient = new HttpClient();
            GetTokenFromXml();
        }

        private void GetTokenFromXml()
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(@"SlackToken.xml");
            var xmlToken = xmlDoc.SelectSingleNode("root/slack/token/value").InnerText;
            TOKEN = xmlToken;
        }

        private HttpRequestMessage CreateRequest(HttpMethod httpMethod, string requestEndPoint)
        {
            var request = new HttpRequestMessage(httpMethod, requestEndPoint);
            return this.AddHeaders(request);
        }

        private HttpRequestMessage AddHeaders(HttpRequestMessage request)
        {
            request.Headers.Add("Accept", "application/json");
            request.Headers.Add("Accept-Charset", "utf-8");
            request.Headers.Add($"Authorization", string.Format("Bearer {0}", this.TOKEN));
            return request;
        }

        private HttpResponseMessage GetMethod(string endPoint)
        {
            HttpRequestMessage request = this.CreateRequest(HttpMethod.Get, endPoint);

            Task<HttpResponseMessage> response;
            try
            {
                response = httpClient.SendAsync(request);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return new HttpResponseMessage();
            }

            return response.Result;
        }

        public string GetChannelList()
        {
            string requestEndPoint = this.baseUrl + "/conversations.list";

            var response = GetMethod(requestEndPoint);
            if (response == null)
            {
                Console.WriteLine("Failed to get request");
                return string.Empty;
            }

            var resStatus = response.StatusCode;
            var resBody = response.Content.ReadAsStringAsync().Result;
            if (!resStatus.Equals(System.Net.HttpStatusCode.OK))
            {
                Console.WriteLine(resStatus.ToString());
                return String.Empty;
            }
            if (string.IsNullOrEmpty(resBody))
            {
                Console.WriteLine("Response body is empty.");
                return string.Empty;
            }
            return resBody;
        }

        public string GetChannelHistory(string channelId, string timestamp)
        {
            if (string.IsNullOrEmpty(channelId))
            {
                return string.Empty;
            }
            string requestEndPoint = this.baseUrl + "/conversations.history?channel=" + channelId;
            if (!string.IsNullOrEmpty(timestamp))
            {
                requestEndPoint += "&oldest=" + timestamp;
            }

            var response = GetMethod(requestEndPoint);
            if (response == null)
            {
                Console.WriteLine("Failed to get request");
                return string.Empty;
            }

            var resStatus = response.StatusCode;
            var resBody = response.Content.ReadAsStringAsync().Result;
            if (!resStatus.Equals(System.Net.HttpStatusCode.OK))
            {
                Console.WriteLine(resStatus.ToString());
                return String.Empty;
            }
            if (string.IsNullOrEmpty(resBody))
            {
                Console.WriteLine("Response body is empty.");
                return string.Empty;
            }
            return resBody;
        }

    }
}
