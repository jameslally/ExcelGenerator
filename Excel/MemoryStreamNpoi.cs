using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Excel.Npoi
{

    public class MemoryStreamNpoi : MemoryStream
    {
        public MemoryStreamNpoi()
        {
            AllowClose = true;
        }

        public bool AllowClose { get; set; }

        protected override void Dispose(bool disposing)
        {
            if (AllowClose)
                base.Dispose(disposing);
        }
    }
}
