using System.Text;
using System.Diagnostics;
using Codeer.Friendly.Windows;
using Codeer.Friendly.Dynamic;
using Codeer.Friendly.Windows.Grasp;
using AI.Talk.Core;
using NAudio.Wave;
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
        private string TOKEN;

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
            var xmlToken = xmlDoc.SelectSingleNode("token/value").InnerText;
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
            public string Text { get;set; } = string.Empty;
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

            //UIAutomation(yukarisan);
            //Program.FromGUI(yukarisan, "こんにちは");

            ReadRssMessageViaSlack(yukarisan);
            //Program.FromAPI(app);
        }

        static void UIAutomation(Process proc)
        {
            AutomationElement elem = AutomationElement.FromHandle(proc.MainWindowHandle);
            var hoge = elem.FindFirst(TreeScope.Element | TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "StatusBarItem"));
            if (hoge == null) return;

            Console.WriteLine(hoge.Current.Name);

            var tab = elem.FindFirst(TreeScope.Element | TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "TextEditView"));
            if(tab == null) return;

            var tabControl = tab.FindFirst(TreeScope.Element | TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "TabControl"));
            if( tabControl == null) return;

            var rawWalker = TreeWalker.RawViewWalker;
            Queue<AutomationElement> queue = new Queue<AutomationElement>();
            queue.Enqueue(tab);
            while(queue.Count > 0)
            {
                var q = queue.Dequeue();
                Console.WriteLine(String.Format("{0}{1}: {2}", new String('-', (int)(queue.Count*2)), q.Current.ClassName, q.Current.Name));

                var sibling = rawWalker.GetNextSibling(q);
                if(sibling != null) queue.Enqueue(sibling);

                var qchild = rawWalker.GetFirstChild(q);
                if(qchild != null) queue.Enqueue(qchild);
            }

            var textbox = tab.FindFirst(TreeScope.Element | TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "TextBox"));
            if(textbox == null) return;
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
                    if(status == null)
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
                        Console.Write("Now Yukarisan reading " + bars[count%4]);
                        Thread.Sleep(100);
                        count++;
                    }
                }
                Console.WriteLine("\n+++++++++++++");
            }
        }

        static void FromAPI(WindowsAppFriend app)
        {
            WindowsAppExpander.LoadAssembly(app, typeof(Program).Assembly);
            dynamic injected_program = app.Type(typeof(Program));

            string kana = "<S>ア^フリカオーリツガ!クイン<F>";
            //kana = injected_program.TextToKana("こんにちは");
            if (kana == String.Empty)
            {
                Console.WriteLine("Failed to TextToKana()");
                return;
            }

            short[] rawBuf;
            rawBuf = injected_program.TextToSpeech(kana);
            if (rawBuf.Length == 0)
            {
                Console.WriteLine("Failed to TextToSpeec()");
                return;
            }

            byte[] pcm_bytes = new byte[rawBuf.Length * 2];
            for (int i = 0; i < rawBuf.Length; i++)
            {
                byte[] b = BitConverter.GetBytes(rawBuf[i]);
                pcm_bytes[i * 2] = b[0];
                pcm_bytes[i * 2 + 1] = b[1];
            }
            MemoryStream ms = new MemoryStream();
            ms.Write(pcm_bytes, 0, pcm_bytes.Length);
            ms.Position = 0;

            var wave_stream = new RawSourceWaveStream(ms, new WaveFormat(44100, 16, 1));
            var wave_out = new WaveOutEvent();
            wave_out.Init(wave_stream);
            wave_out.Play();
            while (wave_out.PlaybackState == PlaybackState.Playing)
            {
                Thread.Sleep(100);
            }
            wave_out.Dispose();

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

        static string TextToKana(string text)
        {
            AITalkResultCode resultCode;
            AITalk_TJobParam jobParam;
            jobParam.modeInOut = AITalkJobInOut.AITALKIOMODE_PLAIN_TO_AIKANA;
            jobParam.userData = IntPtr.Zero;            

            int jobId;
            resultCode = AITalkAPI.TextToKana(out jobId, ref jobParam, text);
            if(resultCode != AITalkResultCode.AITALKERR_SUCCESS)
            {
                Console.WriteLine("TextToKana() returned: " + resultCode);
                return String.Empty;
            }
            Console.WriteLine($"Job ID: {jobId}");

            StringBuilder sb = new StringBuilder(0x100);
            uint size, pos;
            resultCode = AITalkAPI.GetKana(jobId, sb, (uint)sb.Capacity, out size, out pos);
            string kana = sb.ToString();
            if(resultCode != AITalkResultCode.AITALKERR_SUCCESS)
            {
                Console.WriteLine("GetKana() returned: " + resultCode);
                return String.Empty;
            }

            resultCode = AITalkAPI.CloseKana(jobId);
            if(resultCode != AITalkResultCode.AITALKERR_SUCCESS)
            {
                Console.WriteLine("CloseKana() returned" + resultCode);
                return String.Empty;
            }

            return kana;
        }

        static short[] TextToSpeech(string text)
        {
            AITalk_TJobParam jobParam;
            jobParam.modeInOut = AITalkJobInOut.AITALKIOMODE_AIKANA_TO_WAVE;
            jobParam.userData = IntPtr.Zero;

            int jobId;
            AITalkResultCode resultCode = AITalkAPI.TextToSpeech(out jobId, ref jobParam, text);
            if(resultCode != AITalkResultCode.AITALKERR_SUCCESS)
            {
                Console.WriteLine("TextToSpeec() returned" + resultCode);
                return new short[0];
            }

            AITalkStatusCode status;
            do
            {
                Thread.Sleep(1);
                AITalkAPI.GetStatus(jobId, out status);
            } while ((status == AITalkStatusCode.AITALKSTAT_INPROGRESS) || (status == AITalkStatusCode.AITALKSTAT_STILL_RUNNING));
            if(status != AITalkStatusCode.AITALKSTAT_DONE)
            {
                Console.WriteLine($"{status.ToString()}", "AITalkStatus is wrong state.");
                return new short[0];
            }

            short[] rawBuf = new short[0x2B110];
            uint size;
            resultCode = AITalkAPI.GetData(jobId, rawBuf, 0x2B110, out size);
            if(resultCode != AITalkResultCode.AITALKERR_SUCCESS)
            {
                Console.WriteLine("GetKana() returned" + resultCode);
                return new short[0];
            }

            resultCode = AITalkAPI.CloseSpeech(jobId);
            if(resultCode != AITalkResultCode.AITALKERR_SUCCESS)
            {
                Console.WriteLine("CloseSpeec() returned" + resultCode);
                return new short[0];
            }
            Array.Resize<short>(ref rawBuf, (int)size);
            return rawBuf;
        }
    }
}