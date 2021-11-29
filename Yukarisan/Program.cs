using System.Text;
using System.Diagnostics;
using Codeer.Friendly.Windows;
using Codeer.Friendly.Windows.Grasp;
using RM.Friendly.WPFStandardControls;
using Newtonsoft.Json;
using System.Windows.Automation;
using System.Xml;
using System.Text.RegularExpressions;

namespace Yukarisan
{
    class GoogleCloudTranslationClient
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
                return null;
            }

            return response.Result;
        }
        private string ExtractTranslatedString(string json)
        {
            var messageObjects = JsonConvert.DeserializeObject<ResultObject>(json);
            if(messageObjects.Data.Translations.Count == 0)
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
    class SlackAPIHttpClient
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
                return null;
            }

            return response.Result;
        }

        public string GetChannelList()
        {
            string requestEndPoint = this.baseUrl + "/conversations.list";
            
            var response = GetMethod(requestEndPoint);
            if(response == null)
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

        public string GetChannelHistory(string channelId)
        {
            string requestEndPoint = this.baseUrl + "/conversations.history?channel=" + channelId;

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
    class Program
    {
        public class Messages
        {
            public List<Message> messages { get; set; } = new List<Message>();
        }
        public class Message
        {
            public string UserName { get; set; } = string.Empty;
            public string Text { get; set; } = string.Empty;
            public List<Attachments> Attachments { get; set; } = new List<Attachments>();
        }
        public class Attachments
        {
            public string Title { get; set; } = string.Empty;
            public string Title_Link { get; set; } = string.Empty;
            public string Text { get; set; } = string.Empty;
        }

        public class SlackChannel
        {
            public string name = string.Empty;
            public string desc = string.Empty;
            public string id = string.Empty;
        }

        static List<SlackChannel> ListedChannels()
        {
            var channels = new List<SlackChannel>();

            var xmlDoc = new XmlDocument();
            xmlDoc.Load(@"SlackToken.xml");
            var xmlToken = xmlDoc.SelectNodes("root/slack/channels/channel");
            foreach(XmlNode channel in xmlToken)
            {
                var ch = new SlackChannel();
                ch.name = channel.SelectSingleNode("name").InnerText;
                ch.desc = channel.SelectSingleNode("desc").InnerText;
                ch.id = channel.SelectSingleNode("id").InnerText;

                channels.Add(ch);
            }

            return channels;
        }
        static void Main(string[] args)
        {
            List<Process> aivoiceProcess = Process.GetProcessesByName("AIVoiceEditor").ToList();
            if (aivoiceProcess.Count == 0)
            {
                Console.WriteLine("AIVoiceEditor.exe is not runnning. Try to start.");
                Process p = Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\\AI\\AIVoice\\AIVoiceEditor\\AIVoiceEditor.exe");
                Console.WriteLine("Start process, wait 10 sec for complete to start...");
                Thread.Sleep(10 * 1000);
                if(p == null || p.HasExited)
                {
                    Console.WriteLine("Failed to start process. Aborted");
                    return;
                }
                aivoiceProcess.Add(p);
            }
            Process yukarisan = aivoiceProcess[0];

            ReadRssMessageViaSlack(yukarisan);
        }

        static bool isEnglish(string text)
        {
            return !Regex.IsMatch(text, @"[\p{IsHiragana}\p{IsKatakana}\p{IsCJKUnifiedIdeographs}]+");
        }

        static void ReadRssMessageViaSlack(Process proc)
        {
            WindowsAppFriend app = new WindowsAppFriend(proc);

            var channelList = ListedChannels();
            const string baseUrl = "https://slack.com/api";
            var slackClient = new SlackAPIHttpClient(baseUrl);
            var translate = new GoogleCloudTranslationClient();

            foreach (var channel in channelList)
            {
                Console.WriteLine(channel.name + ":" + channel.desc);
                var responseBody = slackClient.GetChannelHistory(channel.id);

                string regexLink = @"<(?<URL>https?.*?)\|(?<Title>.*?)>";
                Regex regex = new Regex(string.Format(@"^{0}(?<Text>.*)", regexLink), RegexOptions.Singleline);

                AutomationElement elem = AutomationElement.FromHandle(proc.MainWindowHandle);

                var messageObjects = JsonConvert.DeserializeObject<Messages>(responseBody);
                foreach (var messageObject in messageObjects.messages)
                {
                    Console.WriteLine(messageObject.UserName);

                    var targetText = messageObject.Text.Replace("\r", "").Replace("\n", "");
                    var textMatch = regex.Match(targetText);
                    if (textMatch.Success)
                    {
                        string url = textMatch.Groups["URL"].Value.ToString();
                        string title = textMatch.Groups["Title"].Value.ToString();
                        string text = textMatch.Groups["Text"].Value.ToString();
                        text = Regex.Replace(text, regexLink, "");

                        if (text == String.Empty)
                        {
                            text = messageObject.Attachments.FirstOrDefault().Text;
                        }

                        foreach(string d in new string[] {"…", "..." })
                        {
                            int dot = text.LastIndexOf(d);
                            if (dot != -1)
                            {
                                text = text.Remove(dot);
                            }
                        }

                        Console.WriteLine(title);
                        Console.WriteLine(url);
                        Console.WriteLine(text);

                        foreach (var v in new Dictionary<string, string> { { "title", title }, { "content", text } })
                        {
                            string str = v.Value;
                            if (isEnglish(v.Value))
                            {
                                var translated = translate.GetTranslation(v.Value);
                                str = translated;
                            }
                            if (v.Key == "title")
                            {
                                str = "タイトル: " + str;
                            }
                            Program.FromGUI(app, str);

                            var status = elem.FindFirst(TreeScope.Element | TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "StatusBarItem"));
                            if (status == null)
                            {
                                Console.WriteLine("Could not detect status bar text. Sleep 10 sec.");
                                Thread.Sleep(10000);
                                continue;
                            }

                            uint count = 0;
                            char[] bars = { '/', '-', '\\', '|' };
                            while (!status.Current.Name.Equals("テキストの読み上げは完了しました。"))
                            {
                                Console.CursorLeft = 0;
                                Console.Write("Now Yukarisan reading " + bars[count % 4]);
                                Thread.Sleep(100);
                                count++;
                            }
                        }

                    }
                    Console.WriteLine("\n+++++++++++++");
                }

            }            
        }

        static void FromGUI(WindowsAppFriend app, string text)
        {
            var topLevel = app.GetTopLevelWindows();
            var editview = topLevel.Single().GetFromTypeFullName("AI.Talk.Editor.TextEditView");

            // Detect TextBox and edit text.
            var textbox = editview.Single().LogicalTree().ByType<System.Windows.Controls.TextBox>();
            var talkTextBox = new WPFTextBox(textbox.Single());
            talkTextBox.EmulateChangeText(text);

            // Detect "Play" button and emulate click.
            // NOTE: In detection, below code suppose first button element is "Play" button. 
            var button = editview.Single().VisualTree().ByType<System.Windows.Controls.Button>();
            var talkPlayButton = new WPFButtonBase(button.First());
            talkPlayButton.EmulateClick();

        }
    }
}