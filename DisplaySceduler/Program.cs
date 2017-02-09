using System;
using System.Collections.Generic;
using System.Diagnostics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;



namespace DisplaySceduler
{
    class Program
    {

        static int DebugLevel = 0;
        static void Main(string[] args)
        {
            string Path = "C:\\slides\\";
            int slideDuration = 10;
            try
            {
                if (args.Length > 0)
                {
                    for (int i = 0; i < args.Length; i++)
                    {
                        if (args[i].Equals("-v"))
                        {
                            DebugLevel = 10;
                        }
                        else if (args[i].Equals("-d"))
                        {
                            i++;
                            slideDuration = Convert.ToInt32(args[i]);
                        }
                        else if (args[i].Equals("-p"))
                        {
                            i++;
                            Path = args[i];
                        }
                        else throw new Exception();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("DisplaySceduler kann mit folgenden optionen Aufgerufen werden:");
                Console.WriteLine("-v : zeige Debugausgaben und halte Ausgabe an");
                Console.WriteLine("-d Zahl: Die Zahl gibt die Anzahl an Sekunden an die jede Folie gezeigt werden soll, für die kein Automatischer Übergang festgelegt ist. Als Standardwert ist 10 Sekunden festgelegt.");
                Console.WriteLine("-p Pfad : Pfad gibt den Pfad mit den Präsentationsdateien an. Standard C:\\slides\\");
                Console.WriteLine("press Enter to continue");
                Console.ReadLine();
                return;
            }
            if(!Path.EndsWith("\\"))
            {
                Path = Path + "\\";
            }
            if( DebugLevel > 0 )
            {
                Console.WriteLine("Pfad : {0}", Path);
                Console.WriteLine("sildeDuration : {0}", slideDuration);
                Console.WriteLine("DebugLevel : {0}", DebugLevel);
                Console.Write("press Enter to continue");
                Console.ReadLine();
            }
            List<string> AcceptedSuffixes = new List<string>();
            AcceptedSuffixes.Add(".ppt");
            AcceptedSuffixes.Add(".pptx");
//            AcceptedSuffixes.Add(".pdf");
//            AcceptedSuffixes.Add(".pwww");
            Console.WriteLine("Pfad für Präsentationen: {0}", Path);
            PresenterFolder Folder = new PresenterFolder(AcceptedSuffixes, Path);
            if (DebugLevel > 0)
            {
                Console.WriteLine(Folder.ToString());
                Console.Write("press Enter to continue");
                Console.ReadLine();
            }
            PresenterFile CurrentPresenterFile = Folder.GetCurrentPresenterFile();
            if (DebugLevel > 0)
            {
                Console.WriteLine(CurrentPresenterFile.ToString());
                Console.Write("press Enter to continue");
                Console.ReadLine();
            }
            try
            {
                Console.WriteLine("schliesse alte Präsentationen");
                KillProcessByName("POWERPNT");
                PowerPoint.Application oPPT;
                PowerPoint.Presentations objPresSet;
                PowerPoint.Presentation objPres;

                //the location of your powerpoint presentation
                string strPres = CurrentPresenterFile.Filename;

                //Create an instance of PowerPoint.
                oPPT = new Microsoft.Office.Interop.PowerPoint.Application();

                // Show PowerPoint to the user.
                oPPT.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

                objPresSet = oPPT.Presentations;

                //open the presentation
                objPres = objPresSet.Open(strPres, MsoTriState.msoFalse,
                MsoTriState.msoTrue, MsoTriState.msoTrue);
                objPres.SlideShowSettings.LoopUntilStopped = MsoTriState.msoTrue;
                foreach (PowerPoint.Slide slide in objPres.Slides)
                {
                    if(slide.SlideShowTransition.AdvanceOnTime != MsoTriState.msoTrue)
                    {
                        Console.WriteLine(slide.SlideShowTransition.AdvanceTime);
                        slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
                        slide.SlideShowTransition.AdvanceTime = slideDuration;
                    }

                }
                objPres.SlideShowSettings.Run();
            }
            catch (Exception Ex)
            {
                Console.WriteLine("Die gewählte Datei kann nicht abgespielt werden.");
                Console.WriteLine(Ex.Message);
            }


        }
        static void TestFileName(string Filename)
        {
            PresenterFile testPresenterFile = new PresenterFile(Filename);
            Console.WriteLine(testPresenterFile.ToString());
        }
        public static void KillProcessByName(string nameToKill)
        {
            foreach (Process process in Process.GetProcesses())
            {
                if (process.ProcessName == nameToKill)
                    process.Kill();
            }
        }

    }
}
