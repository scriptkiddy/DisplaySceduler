using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace DisplaySceduler
{
    enum PresentationType { PowerPoint, PDF, WWW}
    class PresenterFile
    {
        public string Filename { get; private set; }
        public string DateTimeString { get; private set; }        
        public DateTime StartOfDisplayTime { get; private set; }
        public DateTime EndOfDisplayTime { get; private set; }
        public PresentationType Type{ get; private set; }
        public PresenterFile(string Filename, DateTime StartOfDisplayTime, DateTime EndOfDisplayTime)
        {
            this.Filename = Filename;
            this.DateTimeString = "_default";
            this.Type = PresentationType.PowerPoint;
            this.StartOfDisplayTime = StartOfDisplayTime;
            this.EndOfDisplayTime = EndOfDisplayTime;
        }
        public PresenterFile (string Filename)
        {
            this.Filename = Filename;
            Console.WriteLine("hier " + this.Filename);
            if (Filename.EndsWith(".pptx"))
            {
                Type = PresentationType.PowerPoint;
                this.DateTimeString = Path.GetFileName(Filename).Substring(4);
                this.DateTimeString = this.DateTimeString.Substring(0, DateTimeString.IndexOf('.'));
            }
            else if (Filename.EndsWith(".ppt"))
            {
                Type = PresentationType.PowerPoint;
                //this.DateTimeString = Path.GetFileName(Filename).Substring(4, Path.GetFileName(Filename).Length - 8);
                this.DateTimeString = Path.GetFileName(Filename).Substring(4);
                this.DateTimeString = this.DateTimeString.Substring(0, DateTimeString.IndexOf('.'));
            }
            else if (Filename.EndsWith(".pdf"))
            {
                Type = PresentationType.PDF;
                //this.DateTimeString = Path.GetFileName(Filename).Substring(4, Path.GetFileName(Filename).Length - 8);
                this.DateTimeString = Path.GetFileName(Filename).Substring(4);
                this.DateTimeString = this.DateTimeString.Substring(0, DateTimeString.IndexOf('.'));
            }
            else if (Filename.EndsWith(".pwww"))
            {
                Type = PresentationType.WWW;
                //this.DateTimeString = Path.GetFileName(Filename).Substring(4, Path.GetFileName(Filename).Length - 9);
                this.DateTimeString = Path.GetFileName(Filename).Substring(4);
                this.DateTimeString = this.DateTimeString.Substring(0, DateTimeString.IndexOf('.'));
            }
            else
            {
                throw new PresenterFileException("File must be .pptx, .ppt, .pdf or .pwww");
            }
            if(this.DateTimeString.Equals("_default"))
            {
                StartOfDisplayTime = new DateTime(1, 1, 1, 0, 0, 0);
                EndOfDisplayTime = new DateTime(4000, 12, 31, 23, 59, 59);
            }
            else if(this.DateTimeString.Contains("-"))
            {
                if(DateTimeString.Length == 27)
                {
                    short startyear = Int16.Parse(DateTimeString.Substring(0, 4));
                    short startmonth = Int16.Parse(DateTimeString.Substring(4, 2));
                    short startday = Int16.Parse(DateTimeString.Substring(6, 2));
                    short starthour = Int16.Parse(DateTimeString.Substring(9, 2));
                    short startminute = Int16.Parse(DateTimeString.Substring(11, 2));
                    short endyear = Int16.Parse(DateTimeString.Substring(14, 4));
                    short endmonth = Int16.Parse(DateTimeString.Substring(18, 2));
                    short endday = Int16.Parse(DateTimeString.Substring(20, 2));
                    short endhour = Int16.Parse(DateTimeString.Substring(23, 2));
                    short endminute = Int16.Parse(DateTimeString.Substring(25, 2));
                    StartOfDisplayTime = new DateTime(startyear, startmonth, startday, starthour, startminute, 0);
                    EndOfDisplayTime = new DateTime(endyear, endmonth, endday, endhour, endminute, 00);
                }

            }else
            {
                if(DateTimeString.Length == 8)
                {
                    short year = Int16.Parse(DateTimeString.Substring(0, 4));
                    short month = Int16.Parse(DateTimeString.Substring(4, 2));
                    short day = Int16.Parse(DateTimeString.Substring(6, 2));
                    StartOfDisplayTime = new DateTime(year, month, day, 0, 0, 0);
                    EndOfDisplayTime = new DateTime(year, month, day, 23, 59, 59);
                }
            }
            if (StartOfDisplayTime > EndOfDisplayTime)
            {
                throw new PresenterFileException("This Presentation seems to end before it started");
            }
        }

        public override string ToString()
        {
            return String.Format("--------------------------------------------------------------------------------{0}\n{1}\n{2}\n{3}\n{4}\n{5}\n{6}\n--------------------------------------------------------------------------------",
                this.Filename,
                this.Type,
                this.DateTimeString,
                this.StartOfDisplayTime,
                this.EndOfDisplayTime,
                this.GetTimeSpan(),
                DateTime.Now
                );
        }

        public TimeSpan GetTimeSpan()
        {
            return EndOfDisplayTime - StartOfDisplayTime;
        }

        public bool IsCurrent()
        {
            DateTime Now = DateTime.Now;
            return this.StartOfDisplayTime <= Now && this.EndOfDisplayTime >= Now;
        }

    }
}
