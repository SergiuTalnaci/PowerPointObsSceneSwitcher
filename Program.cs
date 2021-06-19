using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
//Thanks to CSharpFritz and EngstromJimmy for their gists, snippets, and thoughts.
//Thanks to Scott Hanselman for his youtube video and github repo

//Slightly modified by Sergiu Vlad Talnaci

namespace PowerPointToOBSSceneSwitcher
{
    class Program
    {
        private static Application ppt = new Microsoft.Office.Interop.PowerPoint.Application();
        private static ObsLocal OBS;
        static async Task Main(string[] args)
        {
            Console.Write("Connecting to PowerPoint...");
            ppt.SlideShowNextSlide += App_SlideShowNextSlide;
            Console.WriteLine("connectedz\n");

            Console.WriteLine("Input password to connect to OBS:");
            var password = GetPassword(); 
            Console.WriteLine("Connecting to OBS...");
            OBS = new ObsLocal(password);
            await OBS.Connect();
            Console.WriteLine("connected");

            OBS.GetScenes();

            Console.ReadLine();
        }


        async static void App_SlideShowNextSlide(SlideShowWindow Wn)
        {
            if (Wn != null)
            {
                Console.WriteLine($"Moved to Slide Number {Wn.View.Slide.SlideNumber}");
                //Text starts at Index 2 ¯\_(ツ)_/¯
                var note = String.Empty;
                try { note = Wn.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text; }
                catch { /*no notes*/ }

                bool sceneHandled = false;


                var notereader = new StringReader(note);
                string line;
                while ((line = notereader.ReadLine()) != null)
                {
                    if (line.StartsWith("OBS:"))
                    {
                        line = line.Substring(4).Trim();

                        if (!sceneHandled)
                        {
                            Console.WriteLine($"  Switching to OBS Scene named \"{line}\"");
                            try
                            {
                                sceneHandled = OBS.ChangeScene(line);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"  ERROR: {ex.Message.ToString()}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"  WARNING: Multiple scene definitions found.  I used the first and have ignored \"{line}\"");
                        }
                    }

                    if (line.StartsWith("OBSDEF:"))
                    {
                        OBS.DefaultScene = line.Substring(7).Trim();
                        Console.WriteLine($"  Setting the default OBS Scene to \"{OBS.DefaultScene}\"");
                    }

                    if (line.StartsWith("**START"))
                    {
                        OBS.StartRecording();
                    }

                    if (line.StartsWith("**STOP"))
                    {
                        OBS.StopRecording();
                    }

                    if (!sceneHandled)
                    {
                        OBS.ChangeScene(OBS.DefaultScene);
                        Console.WriteLine($"  Switching to OBS Default Scene named \"{OBS.DefaultScene}\"");
                    }
                }
            }
        }

        public static string GetPassword()
        {
            var pwd = "";
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.Remove(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else if (i.KeyChar != '\u0000') // KeyChar == '\u0000' if the key pressed does not correspond to a printable character, e.g. F1, Pause-Break, etc
                {
                    pwd += (i.KeyChar);
                    Console.Write("*");
                }
            }
            Console.WriteLine();
            return pwd;
        }
    }
}