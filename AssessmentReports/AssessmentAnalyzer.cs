using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DuoVia.FuzzyStrings;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace AssessmentReportsV2
{
    public class AssessmentAnalyzer
    {
        private readonly AssessmentOptions _options;

        public AssessmentAnalyzer(AssessmentOptions options)
        {
            _options = options;
        }

        public void Execute()
        {
            var scores = ReadScores();

            Console.WriteLine($"{scores.Count} scores read.");

            if (!string.IsNullOrWhiteSpace(_options.CurrentSemester))
            {
                var studentNames = scores.Where(s => s.Semester == _options.CurrentSemester)
                                         .Select(s => s.StudentIdentifier)
                                         .ToHashSet();
                scores = scores.Where(s => studentNames.Contains(s.StudentIdentifier)).ToList();
            }

            Console.WriteLine($"Using {scores.Count} scores from students included in the {_options.CurrentSemester} semester scores.");

            var messages = ScoreValidator.ValidateScores(scores);
            foreach (var message in messages.OrderBy(m => ConvertToSortableValue(m.Item1)))
            {
                Console.Error.WriteLine(message.Item2);
            }

            if (_options.ValidateOnly)
                return;

            WriteResults(scores);
        }

        private List<AssessmentScore> ReadScores()
        {
            var scores = new List<AssessmentScore>();
            using (var package = new ExcelPackage(new FileInfo(_options.Filename)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault(s => _options.SheetName.Equals(s.Name, StringComparison.OrdinalIgnoreCase));
                if (worksheet == null)
                {
                    throw new InvalidOperationException($"Worksheet {_options.SheetName} not found.");
                }
                if (worksheet.Dimension == null)
                {
                    throw new InvalidOperationException($"Worksheet {_options.SheetName} has no data.");
                }

                var startCol = GetNumberForColumn(_options.StartColumn ?? "A");
                int endCol;
                if (_options.LastColumn == null)
                    endCol = worksheet.Dimension.Columns - 1;
                else
                    endCol = GetNumberForColumn(_options.LastColumn);
                var skipCols = _options.SkipColumns?.Select(c => GetNumberForColumn(c)).ToArray();

                var map = new Dictionary<string, int>();
                var approachMap = new Dictionary<int, string>();
                var approachesSeen = 0;
                for (var col = 1; col <= endCol; ++col)
                {
                    var headerValue = CleanString(worksheet.Cells[1, col].GetValue<string>());
                    if (string.IsNullOrWhiteSpace(headerValue)) continue;
                    if ("Approach".Equals(headerValue, StringComparison.OrdinalIgnoreCase))
                    {
                        switch (++approachesSeen)
                        {
                            case 1:
                                approachMap[col] = "Acting/Directing Approach";
                                break;
                            case 2:
                                approachMap[col] = "Design Approach";
                                break;
                            case 3:
                                approachMap[col] = "Educational Approach";
                                break;
                            case 4:
                                approachMap[col] = "Other Project Approach";
                                break;
                        }
                    }
                    else
                    {
                        map[headerValue] = col;
                    }

                }
                for (var row = 2; row <= worksheet.Dimension.Rows; ++row)
                {
                    var firstName = CleanString(worksheet.Cells[row, map["First Name"]].GetValue<string>());
                    var lastName = CleanString(worksheet.Cells[row, map["Last Name"]].GetValue<string>());

                    var name = firstName + " " + lastName;
                    if (string.IsNullOrWhiteSpace(name)) continue;

                    var classStanding = CleanString(worksheet.Cells[row, map["Class Standing"]].GetValue<string>());
                    var emphasis = CleanString(worksheet.Cells[row, map["Emphasis"]].GetValue<string>());
                    var semester = CleanString(worksheet.Cells[row, map["Semester/Date"]].GetValue<string>());

                    if (string.IsNullOrWhiteSpace(firstName))
                    {
                        Console.Error.WriteLine($"Row {row}: First Name is missing");
                    }
                    if (string.IsNullOrWhiteSpace(lastName))
                    {
                        Console.Error.WriteLine($"Row {row}: Last Name is missing");
                    }
                    if (string.IsNullOrWhiteSpace(classStanding))
                    {
                        Console.Error.WriteLine($"Row {row}: Class Standing is missing");
                    }
                    if (string.IsNullOrWhiteSpace(emphasis))
                    {
                        Console.Error.WriteLine($"Row {row}: Emphasis is missing");
                    }
                    if (string.IsNullOrWhiteSpace(semester) || !semester.Contains(" "))
                    {
                        Console.Error.WriteLine($"Row {row}: Semester is missing or an unexpected value: '{semester}'; skipping row.");
                        continue;
                    }

                    if (worksheet.Cells[row, startCol, row, endCol].Value is object[,] values)
                    {
                        var hasValue = false;
                        for (var i = 0; i < values.GetLength(1); ++i)
                        {
                            if (values[0, i] != null && !string.IsNullOrWhiteSpace(values[0, i].ToString()))
                            {
                                hasValue = true;
                                break;
                            }
                        }
                        if (!hasValue)
                            continue;
                    }

                    if ((classStanding == "1 - Freshman" || semester.Contains("New Student", StringComparison.OrdinalIgnoreCase)) && emphasis == "Musical Theatre")
                    {
                        emphasis = "Musical Theatre Applicant";
                    }

                    for (var col = startCol; col <= endCol; ++col)
                    {
                        if (skipCols?.Contains(col) ?? false)
                        {
                            continue;
                        }
                        var scoreValue = GetScoreValue(worksheet.Cells[row, col].GetValue<string>());
                        if (!scoreValue.HasValue)
                            continue;
                        var score = new AssessmentScore
                        {
                            FirstName = firstName,
                            LastName = lastName,
                            StudentIdentifier = name,
                            ClassStanding = classStanding,
                            Emphasis = emphasis,
                            Semester = semester,
                            SemesterSort = ConvertToSortableValue(semester),
                            ScoreName = approachMap.ContainsKey(col) ? approachMap[col] : CleanString(worksheet.Cells[1, col].GetValue<string>()),
                            Score = scoreValue.Value
                        };
                        scores.Add(score);
                    }
                }
            }

            return scores;
        }

        private decimal? GetScoreValue(string value)
        {
            if (!string.IsNullOrWhiteSpace(value) && decimal.TryParse(value, out var result))
            {
                return result;
            }

            return null;
        }

        private void WriteResults(List<AssessmentScore> scores)
        {
            //var overallAverage = scores.Where(s => s.ScoreName == "Metacognition")
            //                           .Average(s => s.Score);
            //Console.WriteLine("Overall Metacognition average: " + overallAverage);

            var scoreAverages = scores.GroupBy(s => new
            {
                s.StudentIdentifier,
                s.Semester,
                s.SemesterSort,
                s.ScoreName
            })
            .Select(g => new
            {
                g.Key.StudentIdentifier,
                g.Key.ScoreName,
                g.Key.Semester,
                g.Key.SemesterSort,
                g.First().LastName,
                AverageScore = g.Average(s => s.Score),
                Scores = g.ToArray()
            })
            .OrderBy(g => g.SemesterSort);

            var studentAverages = scoreAverages.GroupBy(g => g.StudentIdentifier).Select(g => new
            {
                StudentIdentifier = g.Key,
                LastName = g.First().LastName,
                Averages = g.ToArray()
            })
            .OrderBy(g => g.LastName);

            var basePath = Path.Combine(Path.GetDirectoryName(_options.Filename), "output");
            if (!Directory.Exists(basePath))
            {
                Directory.CreateDirectory(basePath);
            }
            else
            {
                foreach (var file in Directory.GetFiles(basePath))
                {
                    File.Delete(file);
                }
            }
            Console.WriteLine("Writing results for:");
            foreach (var student in studentAverages)
            {
                Console.WriteLine($"\t{student.StudentIdentifier}...");
                using (var doc = CreateDefaultEmptyWordDocument(Path.Combine(basePath, student.StudentIdentifier + ".docx")))
                {
                    var body = doc.MainDocumentPart.Document.Body;

                    var p = body.AppendChild(new Paragraph());
                    if (p.Elements<ParagraphProperties>().Count() == 0)
                    {
                        p.PrependChild(new ParagraphProperties());
                    }
                    var pp = p.Elements<ParagraphProperties>().First();
                    var styleId = new ParagraphStyleId { Val = "Heading1" };
                    pp.Append(styleId);

                    var run = p.AppendChild(new Run());
                    run.AppendChild(new Text($"Assessment Scores For {student.StudentIdentifier}"));

                    p = body.AppendChild(new Paragraph());
                    run = p.AppendChild(new Run());
                    run.AppendChild(new Text($"\"Average\" is the average score for all students with the same Emphasis and Class Standing in a semester."));

                    var table = new Table();
                    table.AppendChild(new TableProperties(
                        new TableBorders(
                        new TopBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                        new BottomBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                        new LeftBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                        new RightBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                        new InsideHorizontalBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                        new InsideVerticalBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        })));

                    var tr = new TableRow();
                    AddTableCell(tr, "Assessment", true);
                    AddTableCell(tr, "");

                    var semesters = student.Averages.Select(a => a.Semester).Distinct().OrderBy(a => ConvertToSortableValue(a)).ToArray();
                    foreach (var semester in semesters)
                    {
                        AddTableCell(tr, semester, true);
                    }
                    table.Append(tr);

                    var tests = student.Averages.GroupBy(a => a.ScoreName).OrderBy(t => t.Key == "Overall Success" ? 1 : 0).ToArray();
                    foreach (var test in tests)
                    {
                        tr = new TableRow();
                        if (test.Key == "Overall Success")
                        {
                            var tc = AddTableCell(tr, new ParagraphStyleId { Val = "IntenseEmphasis" }, test.Key);
                            AddTableCell(tr, new ParagraphStyleId { Val = "IntenseEmphasis" }, "Yours:", "Average:");
                        }
                        else
                        {
                            AddTableCell(tr, test.Key);
                            AddTableCell(tr, "Yours:", "Average:");
                        }

                        foreach (var semester in semesters)
                        {
                            var average = test.FirstOrDefault(t => t.Semester == semester);
                            if (average != null)
                            {
                                // Get the comparative average
                                var compQ = scores.Where(s => s.SemesterSort == average.SemesterSort
                                                           && s.ClassStanding == average.Scores.First().ClassStanding
                                                           && s.Emphasis == average.Scores.First().Emphasis
                                                           && s.ScoreName == average.ScoreName)
                                                  .Select(s => s.Score);
                                if (test.Key == "Overall Success")
                                {
                                    AddTableCell(tr, new ParagraphStyleId { Val = "IntenseEmphasis" }, average.AverageScore.ToString("0.##"), (compQ.Any() ? compQ.Average().ToString("0.##") : "N/A"));
                                }
                                else
                                {
                                    AddTableCell(tr, average.AverageScore.ToString("0.##"), (compQ.Any() ? compQ.Average().ToString("0.##") : "N/A"));
                                }
                            }
                            else
                            {
                                AddTableCell(tr, "N/A");
                            }
                        }

                        table.Append(tr);
                    }

                    body.AppendChild(table);

                    doc.Save();
                }
            }
        }

        private static TableCell AddTableCell(TableRow tr, string text, bool header = false)
        {
            var tc = new TableCell();
            var run = new Run();
            if (header)
            {
                run.AppendChild(new Bold());
            }
            run.AppendChild(new Text(text));
            tc.Append(new Paragraph(run));
            tc.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));
            tr.Append(tc);
            return tc;
        }

        private static TableCell AddTableCell(TableRow tr, ParagraphStyleId style, params string[] text)
        {
            var tc = new TableCell();
            foreach (var t in text)
            {
                var run = new Run(
                    new RunProperties(new RunStyle { Val = style.Val }),
                    new Text(t)
                );
                var pp = new ParagraphProperties(style.CloneNode(true));
                tc.Append(new Paragraph(pp, run));
            }
            tc.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));
            tr.Append(tc);
            return tc;
        }

        private static TableCell AddTableCell(TableRow tr, params string[] text)
        {
            var tc = new TableCell();
            foreach (var t in text)
            {
                tc.Append(new Paragraph(new Run(new Text(t))));
            }
            tc.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));
            tr.Append(tc);
            return tc;
        }

        private static WordprocessingDocument CreateDefaultEmptyWordDocument(string filename)
        {
            var doc = WordprocessingDocument.Create(filename, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            mainPart.Document.Body = mainPart.Document.AppendChild(new Body());
            var stylePart = doc.MainDocumentPart.StyleDefinitionsPart ?? AddStylesPartToPackage(doc);
            var themePart = doc.MainDocumentPart.ThemePart ?? AddThemePartToPackage(doc);
            // TODO: This really shouldn't be hardcoded.
            //var templatePath = @"C:\Program Files (x86)\Microsoft Office\root\Office16\1033\QuickStyles\Default.dotx";
            var templatePath = @"C:\Program Files\Microsoft Office\root\Office16\1033\QuickStyles\Default.dotx";
            if (File.Exists(templatePath))
            {
                using (WordprocessingDocument wordTemplate = WordprocessingDocument.Open(templatePath, false))
                {
                    foreach (var templateTheme in wordTemplate.MainDocumentPart.ThemePart.Theme)
                    {
                        themePart.Theme.Append(templateTheme.CloneNode(true));
                    }
                    foreach (var templateStyle in wordTemplate.MainDocumentPart.StyleDefinitionsPart.Styles)
                    {
                        stylePart.Styles.Append(templateStyle.CloneNode(true));
                    }
                }
            }
            return doc;
        }

        public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            var root = new Styles();
            root.Save(part);
            return part;
        }

        public static ThemePart AddThemePartToPackage(WordprocessingDocument doc)
        {
            ThemePart part;
            part = doc.MainDocumentPart.AddNewPart<ThemePart>();
            var root = new DocumentFormat.OpenXml.Drawing.Theme();
            root.Save(part);
            return part;
        }

        static readonly char[] _whitespace = new[] { (char)160 };
        private string CleanString(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;
            return new string(value.Select(c => _whitespace.Contains(c) ? ' ' : c).ToArray()).Trim();
        }

        public static string ConvertToSortableValue(string semester)
        {
            if (!semester.Contains(" "))
            {
                throw new Exception($"Unexpected semester value {semester}");
            }
            var s = semester.Split(' ');
            if (semester.Contains("New Student", StringComparison.OrdinalIgnoreCase))
                return s[s.Length - 1] + "0";
            else if (semester.Contains("Fall", StringComparison.OrdinalIgnoreCase) || semester.Contains("Spring", StringComparison.OrdinalIgnoreCase))
                return s[s.Length - 1] + s[s.Length - 2].ToUpper(CultureInfo.CurrentCulture).Replace("FALL", "2").Replace("SPRING", "1");
            else
                return semester;
        }

        private static readonly int _unicodeA = Char.ConvertToUtf32("A", 0);
        private int GetNumberForColumn(string column)
        {
            var intVal = 0;
            var exp = 0;
            var rv = 0;
            foreach (var ch in column.Reverse())
            {
                var charVal = Char.ConvertToUtf32(ch.ToString(), 0);
                if (exp == 0)
                    intVal = charVal - _unicodeA;
                else
                    intVal = charVal - _unicodeA + 1;
                rv += intVal * (int)Math.Pow(26, exp);
                ++exp;
            }
            return rv + 1;
        }
    }
}
