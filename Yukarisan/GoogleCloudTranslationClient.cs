using System.Xml;
using Newtonsoft.Json;

namespace Yukarisan
{
    internal class GoogleCloudTranslationClient
    {
        private readonly string baseUrl = "https://translation.googleapis.com/language/translate/v2?";
        private readonly HttpClient httpClient;
        private string TOKEN = string.Empty;

        public GoogleCloudTranslationClient()
        {
            this.httpClient = new HttpClient();
            GetTokenFromXml();
        }
        private void GetTokenFromXml()
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(@"SlackToken.xml");
            var xmlToken = xmlDoc.SelectSingleNode("root/google/translate/key").InnerText;
            TOKEN = xmlToken;
        }

        public string GetTranslation(string text)
        {
            var request = baseUrl + $"key={TOKEN}&q={text}&detectedSourceLanguage=en&target=ja";
            var response = PostMethod(request);
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

            return ExtractTranslatedString(resBody);
        }

        private HttpRequestMessage CreateRequest(HttpMethod httpMethod, string requestEndPoint)
        {
            var request = new HttpRequestMessage(httpMethod, requestEndPoint);
            return request;
        }

        private HttpResponseMessage PostMethod(string endPoint)
        {
            HttpRequestMessage request = this.CreateRequest(HttpMethod.Post, endPoint);

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
        private string ExtractTranslatedString(string json)
        {
            var messageObjects = JsonConvert.DeserializeObject<ResultObject>(json);
            if (messageObjects == null || messageObjects.Data.Translations.Count == 0)
            {
                return String.Empty;
            }

            return messageObjects.Data.Translations.First().TranslatedText;
        }
        public class ResultObject
        {
            public Data Data { get; set; } = new Data();
        }
        public class Data
        {
            public List<Translate> Translations { get; set; } = new List<Translate>();
        }
        public class Translate
        {
            public string TranslatedText { get; set; } = string.Empty;
            public string DetectedSourceLanguage { get; set; } = string.Empty;
        }
    }
}
