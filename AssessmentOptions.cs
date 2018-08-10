using System;
using System.Collections.Generic;
using System.Text;

namespace AssessmentReportsV2
{
    public class AssessmentOptions
    {
        public string Filename { get; set; }
        public string SheetName { get; set; }
        public string CurrentSemester { get; set; }
        public string StartColumn { get; set; }
        public string LastColumn { get; set; }
        public string[] SkipColumns { get; set; }
        public bool ValidateOnly { get; set; }
    }
}
