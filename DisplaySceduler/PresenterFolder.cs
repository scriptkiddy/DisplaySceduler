using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DisplaySceduler
{
    class PresenterFolder
    {
        public List<PresenterFile> PresenterFiles { get; private set; }
        private List<string> AcceptedSuffixes { get; set; }
        private string Path { get; set; }
        public PresenterFolder(List<String> AcceptedSuffixes, string Path)
        {
            this.AcceptedSuffixes = AcceptedSuffixes;
            this.Path = Path;
            this.PresenterFiles = new List<PresenterFile>();
            string[] Files = System.IO.Directory.GetFiles(this.Path);
            //            string[] Folders = System.IO.Directory.GetDirectories(Root);
            foreach(string File in Files)
            {
                foreach (string Suffix in AcceptedSuffixes)
                {
                    if (File.EndsWith(Suffix))
                    {
                        PresenterFile PresenterFile = new PresenterFile(File);
                        PresenterFiles.Add(PresenterFile);
                        //FileArray.Add(Files[i].ToString());
                        }
                }
            }
        }

        public PresenterFile GetCurrentPresenterFile()
        {
            DateTime Now = DateTime.Now;
            PresenterFile CurrentPresenterFile = new PresenterFile(this.Path + "default.ppt", new DateTime(1,1,1,0,0,0), new DateTime(4000,12,31,23,59,59));
            foreach (PresenterFile PF in this.PresenterFiles)
            {
                if (PF.IsCurrent())
                {
                    if(PF.GetTimeSpan() < CurrentPresenterFile.GetTimeSpan())
                    {
                        CurrentPresenterFile = PF;
                    }
                }
            }
            return CurrentPresenterFile;
        }

        public override string ToString()
        {
            StringBuilder SB = new StringBuilder();
            foreach (PresenterFile PF in PresenterFiles)
            {
                SB.Append(PF.ToString());
            }
            return SB.ToString();
        }
    }
}
