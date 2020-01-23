using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Speech.Recognition;
using Microsoft.Speech.Synthesis;

namespace VoiceRecognition
{
    class Program
    {
        static int Main(string[] args)
        {
            SpeechRec speechRec = new SpeechRec();

            Console.WriteLine("Please press <Return> to exit.");
            Console.ReadLine();

            return 0;
        }
    }

    public class SpeechRec
    {

        private System.Speech.Synthesis.SpeechSynthesizer speechSynthesizer { get; set; } = new System.Speech.Synthesis.SpeechSynthesizer();
        private string[] choicesList { get; set; }
        private Choices choices { get; set; } = new Choices();
        private SpeechRecognitionEngine speechRecognitionEngine { get; set; } = new SpeechRecognitionEngine();
        private bool listen { get; set; } = false;

        public SpeechRec()
        {
            if (SpeechRecognitionEngine.InstalledRecognizers().Count == 0)
            {
                throw new Exception("InstalledRecognizers returned 0");
            }

            choicesList = new string[]
            {
                "computer", "cancel query", "add task", "list task",
                "time", "start chrome", "start notepad", "start battle",
                "start snip", "list choices"
            };

            choices.Add(choicesList);
            Grammar grammar = new Grammar(new GrammarBuilder(choices));
            
            speechRecognitionEngine.RequestRecognizerUpdate();
            speechRecognitionEngine.LoadGrammar(grammar);
            speechRecognitionEngine.SpeechRecognized += SpeechRecognitionEngine_SpeechRecognized;
            speechRecognitionEngine.SetInputToDefaultAudioDevice();
            speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);

            speechSynthesizer.Speak("Ready");
        }

        public void SpeechRecognitionEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string message = e.Result.Text;

            Console.WriteLine("Command {0}", message);

            if (message == "computer" && !listen)
            {
                speechSynthesizer.Speak("Now listening");
                
                listen = true;
            }
            else if (listen)
            {
                if (message == "cancel query")
                {
                    speechSynthesizer.Speak("Listening canceled");
                }
                if (message == "start chrome")
                {
                    if (System.Diagnostics.Process.GetProcessesByName("chrome").Count() == 0) 
                    {
                        System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe");
                    }                    
                }
                if (message == "start notepad")
                {
                    System.Diagnostics.Process.Start(@"C:\Windows\System32\notepad.exe");
                }
                if (message == "start battle")
                {
                    System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Battle.net\Battle.net Launcher.exe");
                }
                if (message == "start snip")
                {
                    System.Diagnostics.Process.Start(@"C:\Windows\System32\SnippingTool.exe");
                }
                if (message == "list task")
                {
                    string[] text = File.ReadAllLines(@"C:\Users\Andreas\Desktop\Tasks.txt");

                    speechSynthesizer.Speak("Tasks to complete");
                    System.Threading.Thread.Sleep(200);
                    foreach (string item in text)
                    {
                        speechSynthesizer.Speak(item);
                        System.Threading.Thread.Sleep(200);
                    }
                }
                if (message == "add task")
                {
                    string input = Microsoft.VisualBasic.Interaction.InputBox("Task info: ", "New task");
                    File.AppendAllText(@"C:\Users\Andreas\Desktop\Tasks.txt", input);
                }
                if (message == "time")
                {
                    speechSynthesizer.Speak(DateTime.Now.ToShortTimeString());
                }
                if (message == "list choices")
                {
                    speechSynthesizer.Speak("The choices are");
                    System.Threading.Thread.Sleep(200);
                    foreach (string item in choicesList)
                    {
                        speechSynthesizer.Speak(item);
                        System.Threading.Thread.Sleep(200);
                    }                    
                }


                listen = false;
            }
            Console.WriteLine("\tlisten == {0}", listen);
        }
    }
}
