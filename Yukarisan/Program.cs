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
    class SlackChannelMessages
    {
        public List<Message> Messages { get; set; } = new List<Message>();

        [Serializable]
        public class Message
        {
            public string UserName { get; set; } = string.Empty;
            public string Text { get; set; } = string.Empty;
            public List<Attachment> Attachments { get; set; } = new List<Attachment>();

            [Serializable]
            public class Attachment
            {
                public string Title { get; set; } = string.Empty;
                public string Title_Link { get; set; } = string.Empty;
                public string Text { get; set; } = string.Empty;

            }
        }
    }
    class Yukarisan
    {
        private class SlackChannelInXml
        {
            public string name = string.Empty;
            public string desc = string.Empty;
            public string id = string.Empty;
        }

        private List<SlackChannelInXml> ListedChannels()
        {
            var channels = new List<SlackChannelInXml>();

            var xmlDoc = new XmlDocument();
            xmlDoc.Load(@"SlackToken.xml");
            var xmlToken = xmlDoc.SelectNodes("root/slack/channels/channel");
            foreach (XmlNode channel in xmlToken)
            {
                var ch = new SlackChannelInXml();
                ch.name = channel.SelectSingleNode("name").InnerText;
                ch.desc = channel.SelectSingleNode("desc").InnerText;
                ch.id = channel.SelectSingleNode("id").InnerText;

                channels.Add(ch);
            }

            return channels;
        }
        private  bool isEnglish(string text)
        {
            return !Regex.IsMatch(text, @"[\p{IsHiragana}\p{IsKatakana}\p{IsCJKUnifiedIdeographs}]+");
        }

        private SlackAPIHttpClient slackClient;
        private GoogleCloudTranslationClient translateClient;
        private WindowsAppFriend appFriend;
        private AutomationElement autoElement;
        private Process voiceProcess;
        private readonly string baseUrl = "https://slack.com/api";
        private readonly string regexLink = @"<(?<URL>https?.*?)\|(?<Title>.*?)>";

        public Yukarisan(Process proc)
        {
            slackClient = new SlackAPIHttpClient(baseUrl);
            translateClient = new GoogleCloudTranslationClient();
            voiceProcess = proc;
            appFriend = new WindowsAppFriend(voiceProcess);
            autoElement = AutomationElement.FromHandle(voiceProcess.MainWindowHandle);
        }

        public void ReadRssMessageViaSlack()
        {
            var channelList = ListedChannels();

            Regex regex = new Regex(string.Format(@"^{0}(?<Text>.*)", regexLink), RegexOptions.Singleline);

            foreach (var channel in channelList)
            {
                var responseBody = slackClient.GetChannelHistory(channel.id);
                var messageObjects = JsonConvert.DeserializeObject<SlackChannelMessages>(responseBody);
                if(messageObjects == null)
                {
                    continue;
                }

                Console.WriteLine(channel.name + ":" + channel.desc);
                TalkAndWait(channel.desc + "についての記事を読むよ。全部で" + messageObjects.Messages.Count + "記事あるよ。");
                var messageList = messageObjects.Messages;
                if(messageList.Count > 10)
                {
                    TalkAndWait("記事が多いから最新 10記事を読むね");
                    messageList = messageList.GetRange(0, 10);
                }
                foreach (var messageObject in messageList)
                {
                    Console.WriteLine(messageObject.UserName);

                    var targetText = messageObject.Text.Replace("\r", "").Replace("\n", "");
                    var textMatch = regex.Match(targetText);
                    if (!textMatch.Success)
                    {
                        continue;
                    }
                    string url = textMatch.Groups["URL"].Value.ToString();
                    string title = textMatch.Groups["Title"].Value.ToString();
                    string text = textMatch.Groups["Text"].Value.ToString();
                    text = Regex.Replace(text, regexLink, "");

                    if (text == String.Empty)
                    {
                        text = messageObject.Attachments.FirstOrDefault().Text;
                    }

                    foreach (string d in new string[] { "…", "..." })
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
                            var translated = translateClient.GetTranslation(v.Value);
                            str = translated;
                        }
                        if (v.Key == "title")
                        {
                            str = "タイトル: " + str;
                        }
                        TalkAndWait(str);
                    }
                    Console.WriteLine("\n+++++++++++++");
                }

                TalkAndWait("このチャンネルの記事は全て読み終えたので次のチャンネルに行くね");
            }
        }

        private void TalkAndWait(string str)
        {
            FromGUI(appFriend, str);

            AutomationElement? status = autoElement.FindFirst(
                TreeScope.Element | TreeScope.Descendants, 
                new PropertyCondition(AutomationElement.ClassNameProperty, "StatusBarItem"));

            uint count = 0;
            char[] bars = { '/', '-', '\\', '|' };
            while (!status.Current.Name.Equals("テキストの読み上げは完了しました。"))
            {
                Console.CursorLeft = 0;
                Console.Write("Now Yukarisan reading " + bars[count % 4]);
                Thread.Sleep(100);
                count++;
            }
            Thread.Sleep(100);
        }

        private void FromGUI(WindowsAppFriend app, string text)
        {
            var topLevel = app.GetTopLevelWindows();
            var editview = topLevel.First().GetFromTypeFullName("AI.Talk.Editor.TextEditView").FirstOrDefault();

            // Detect TextBox and edit text.
            var textbox = editview.LogicalTree().ByType<System.Windows.Controls.TextBox>();
            var talkTextBox = new WPFTextBox(textbox.Single());
            talkTextBox.EmulateChangeText(text);

            // Detect "Play" button and emulate click.
            // NOTE: In detection, below code suppose first button element is "Play" button. 
            var button = editview.VisualTree().ByType<System.Windows.Controls.Button>();
            var talkPlayButton = new WPFButtonBase(button.First());
            talkPlayButton.EmulateClick();
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            List<Process> aivoiceProcesses = Process.GetProcessesByName("AIVoiceEditor").ToList();
            if (aivoiceProcesses.Count == 0)
            {
                Console.WriteLine("AIVoiceEditor.exe is not runnning. Try to start.");
                Process p = Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\\AI\\AIVoice\\AIVoiceEditor\\AIVoiceEditor.exe");
                Console.WriteLine("Start process.");
                if (p == null || p.HasExited)
                {
                    Console.WriteLine("Failed to start process. Aborted");
                    return;
                }

                string windowTitle = string.Empty;
                do
                {
                    Thread.Sleep(100);
                    aivoiceProcesses = Process.GetProcessesByName("AIVoiceEditor").ToList();
                    windowTitle = aivoiceProcesses[0].MainWindowTitle;
                } while(windowTitle != "A.I.VOICE Editor - (新規プロジェクト)");
            }
            Process aivoiceProcess = aivoiceProcesses[0];

            var yukarisan = new Yukarisan(aivoiceProcess);
            yukarisan.ReadRssMessageViaSlack();
        }
    }
}