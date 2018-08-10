using System;
using System.Collections.Generic;
using System.Text;

namespace AssessmentReportsV2
{
    public class AssessmentScore
    {
        public string StudentIdentifier { get; set; }
        public string Semester { get; set; }
        public string SemesterSort { get; set; }
        public string ClassStanding { get; set; }
        public string Emphasis { get; set; }
        public string ScoreName { get; set; }
        public decimal Score { get; set; }
    }
}
