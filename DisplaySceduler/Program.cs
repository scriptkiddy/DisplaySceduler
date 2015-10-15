using System;
using System.Collections.Generic;
using System.Diagnostics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;



namespace DisplaySceduler
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("schliesse alte Präsentationen");
            KillProcessByName("POWERPNT");

            List<string> AcceptedSuffixes = new List<string>();
            AcceptedSuffixes.Add(".ppt");
            AcceptedSuffixes.Add(".pptx");
            AcceptedSuffixes.Add(".pdf");
            AcceptedSuffixes.Add(".pwww");
            string Path = "C:\\slides\\";
            PresenterFolder Folder = new PresenterFolder(AcceptedSuffixes, Path);
            Console.WriteLine(Folder.ToString());
            Console.Write("press any key to continue ...");
            Console.ReadLine();
            PresenterFile CurrentPresenterFile = Folder.GetCurrentPresenterFile();
            Console.WriteLine(CurrentPresenterFile.ToString());
            Console.Write("press any key to continue ...");
            Console.ReadLine();

            try
            {
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
                    Console.WriteLine(slide.SlideShowTransition.AdvanceTime);
                    slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
                    slide.SlideShowTransition.AdvanceTime = 1;

                }
                //Console.ReadLine();
                //TextReader tr = new StreamReader(folder + "zusatz_info\\an.txt");
                //if (tr.ReadLine() == "voll")
                //{
                //    foreach (Microsoft.Office.Interop.PowerPoint.Slide s in objPres.Slides)
                //    {
                //        s.Shapes.AddPicture(folder + "zusatz_info\\das_haus_ist_voll.gif", MsoTriState.msoTrue, MsoTriState.msoFalse, 0, 500, 720, 40);
                //    }
                //}
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
