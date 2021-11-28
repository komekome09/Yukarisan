using System.Text;
using System.Diagnostics;
using Codeer.Friendly.Windows;
using Codeer.Friendly.Windows.Grasp;
using RM.Friendly.WPFStandardControls;
using Newtonsoft.Json;
using System.Windows.Automation;
using System.Xml;

namespace Yukarisan
{
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
            var xmlToken = xmlDoc.SelectSingleNode("slack/token/value").InnerText;
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

        public string GetChannelList()
        {
            string requestEndPoint = this.baseUrl + "/conversations.list";
            HttpRequestMessage request = this.CreateRequest(HttpMethod.Get, requestEndPoint);

            string resBody;
            System.Net.HttpStatusCode resStatus = System.Net.HttpStatusCode.NotFound;
            Task<HttpResponseMessage> response;
            try
            {
                response = httpClient.SendAsync(request);
                resBody = response.Result.Content.ReadAsStringAsync().Result;
                resStatus = response.Result.StatusCode;
            }catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return String.Empty;
            }

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
            HttpRequestMessage request = this.CreateRequest(HttpMethod.Get, requestEndPoint);

            string resBody;
            System.Net.HttpStatusCode resStatus = System.Net.HttpStatusCode.NotFound;
            Task<HttpResponseMessage> response;
            try
            {
                response = httpClient.SendAsync(request);
                resBody = response.Result.Content.ReadAsStringAsync().Result;
                resStatus = response.Result.StatusCode;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return String.Empty;
            }

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
            public List<Attachments> Attachments { get; set; } = new List<Attachments>();
        }
        public class Attachments
        {
            public string Title { get; set; } = string.Empty;
            public string Title_Link { get; set; } = string.Empty;
            public string Text { get; set; } = string.Empty;
        }
        static void Main(string[] args)
        {
            Process[] aivoiceProcess = Process.GetProcessesByName("AIVoiceEditor");
            if (aivoiceProcess.Length == 0)
            {
                Console.WriteLine("AIVoiceEditor.exe is not runnning. Exit");
                return;
            }
            Process yukarisan = aivoiceProcess[0];

            ReadRssMessageViaSlack(yukarisan);
        }

        static void ReadRssMessageViaSlack(Process proc)
        {
            WindowsAppFriend app = new WindowsAppFriend(proc);

            const string baseUrl = "https://slack.com/api";
            const string RSS_VEHICLE = "C018Z1JLM7V";
            var slackClient = new SlackAPIHttpClient(baseUrl);
            var responseBody = slackClient.GetChannelHistory(RSS_VEHICLE);
            //Console.WriteLine(channelList);

            var messageObjects = JsonConvert.DeserializeObject<Messages>(responseBody);
            foreach (var messageObject in messageObjects.messages)
            {
                Console.WriteLine(messageObject.UserName);
                foreach (var attachment in messageObject.Attachments)
                {
                    Console.WriteLine(attachment.Title);
                    Console.WriteLine(attachment.Title_Link);
                    Console.WriteLine(attachment.Text);
                    Program.FromGUI(app, attachment.Text);

                    AutomationElement elem = AutomationElement.FromHandle(proc.MainWindowHandle);
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
                Console.WriteLine("\n+++++++++++++");
            }
        }

        static void FromGUI(WindowsAppFriend app, string text)
        {
            var topLevel = app.GetTopLevelWindows();
            var editview = topLevel.Single().GetFromTypeFullName("AI.Talk.Editor.TextEditView");
            var textbox = editview.Single().LogicalTree().ByType<System.Windows.Controls.TextBox>();
            var talkTextBox = new WPFTextBox(textbox.Single());
            talkTextBox.EmulateChangeText("あいうえo");

            // テキスト入力欄と再生ボタンを特定する
            WindowControl ui_tree_top = WindowControl.FromZTop(app);
            var text_edit_view = ui_tree_top.GetFromTypeFullName("AI.Talk.Editor.TextEditView")[0].LogicalTree();
            WPFTextBox talk_text_box = new WPFTextBox(text_edit_view[7]);
            WPFButtonBase play_button = new WPFButtonBase(text_edit_view[9]);

            // テキストを入力し、再生する
            talk_text_box.EmulateChangeText(text);
            play_button.EmulateClick();
        }
    }
}