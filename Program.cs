﻿using System;
using System.Linq;

namespace AssessmentReportsV2
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new AssessmentOptions
            {
                Filename = @"E:\Users\Eric\Desktop\Juries\Theatre Juries fall 18 orig.xlsx",
                CurrentSemester = "Fall 2018",
                SheetName = "Sheet1",
                StartColumn = "I",
                LastColumn = "AP",
                SkipColumns = new[] { "AH", "AI", "AJ" },
                ValidateOnly = args.Contains("-v")
            };

            var index = args.IndexOf("-file");
            if (index >= 0 && args.Length >= index + 2)
            {
                options.Filename = args[index + 1];
            }
            index = args.IndexOf("-semester");
            if (index >= 0 && args.Length >= index + 2)
            {
                options.CurrentSemester = args[index + 1];
            }
            index = args.IndexOf("-sheet");
            if (index >= 0 && args.Length >= index + 2)
            {
                options.SheetName = args[index + 1];
            }
            index = args.IndexOf("-start");
            if (index >= 0 && args.Length >= index + 2)
            {
                options.StartColumn = args[index + 1];
            }
            index = args.IndexOf("-end");
            if (index >= 0 && args.Length >= index + 2)
            {
                options.LastColumn = args[index + 1];
            }
            index = args.IndexOf("-skip");
            if (index >= 0 && args.Length >= index + 2)
            {
                options.SkipColumns = args[index + 1].Split(',', StringSplitOptions.RemoveEmptyEntries);
            }
            var analyzer = new AssessmentAnalyzer(options);
            analyzer.Execute();
            Console.WriteLine("Finished.");
            Console.ReadLine();
        }
    }
}
