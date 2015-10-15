using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DisplaySceduler
{
    class PresenterFileException : Exception
    {
        public PresenterFileException(string Message)
            : base(Message)
        {
        }
    }
}
