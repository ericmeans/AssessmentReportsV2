using DuoVia.FuzzyStrings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AssessmentReportsV2
{
    public static class ScoreValidator
    {
        private const int _maximumLevenshteinDistance = 2;

        public static List<Tuple<string, string>> ValidateScores(List<AssessmentScore> scores)
        {
            // Validation: Find students who differ only by first or only by last name, or whose first and last name got flipped.
            var splitNames = scores.Select(s => new[] { s.FirstName, s.LastName }).Distinct();
            var possibleFirstMatches = splitNames.GroupBy(s => s[1])
                                                 .Where(g => g.Select(x => x[0]).Distinct().Count() > 1)
                                                 .Select(g => new
                                                 {
                                                     FirstNames = g.Select(s => s[0]).Distinct().ToArray(),
                                                     LastName = g.Key
                                                 })
                                                 .SelectMany(m => m.FirstNames.Select(fn => new { FirstName = fn, m.LastName }));
            var onlyFirstNames = scores.GroupBy(s => s.StudentIdentifier)
                                       .Join(possibleFirstMatches,
                                             g => g.First().LastName,
                                             m => m.LastName,
                                             (s, m) =>
                                             {
                                                var closeFirstNames = GetSimilarNames(s, score => score.FirstName, m.FirstName);
                                                return new
                                                {
                                                    Scores = closeFirstNames.ToArray(),
                                                    m.FirstName,
                                                    m.LastName
                                                };
                                             })
                                       .Where(x => x.Scores.Count() > 0)
                                       .GroupBy(x => x.LastName)
                                       .Select(g => new
                                       {
                                           FirstName = g.OrderBy(x => x.Scores.Count()).First().FirstName,
                                           LastName = g.Key,
                                           MismatchScores = g.OrderBy(x => x.Scores.Count()).First().Scores
                                       })
                                       .ToArray();
            var possibleLastMatches = splitNames.GroupBy(s => s[0])
                                                .Where(g => g.Select(x => x[1]).Distinct().Count() > 1)
                                                .Select(g => new
                                                {
                                                    FirstName = g.Key,
                                                    LastNames = g.Select(s => s[1]).Distinct().ToArray()
                                                })
                                                .SelectMany(m => m.LastNames.Select(ln => new { m.FirstName, LastName = ln }));
            var onlyLastNames = scores.GroupBy(s => s.StudentIdentifier)
                                      .Join(possibleLastMatches,
                                            g => g.First().FirstName,
                                            m => m.FirstName,
                                            (s, m) =>
                                            {
                                                var closeLastNames = GetSimilarNames(s, score => score.LastName, m.LastName);
                                                return new
                                                {
                                                    Scores = closeLastNames.ToArray(),
                                                    m.FirstName,
                                                    m.LastName
                                                };
                                            })
                                      .Where(x => x.Scores.Count() > 0)
                                      .GroupBy(x => x.FirstName)
                                      .Select(g => new
                                      {
                                          FirstName = g.Key,
                                          LastName = g.OrderBy(x => x.Scores.Count()).First().LastName,
                                          MismatchScores = g.OrderBy(x => x.Scores.Count()).First().Scores
                                      })
                                      .ToArray();
            var reversedNames = scores.GroupBy(s => s.LastName + " " + s.FirstName).Select(g => new { ReversedName = g.Key, Count = g.Count(), PossibleMatches = g.ToArray() });
            var flippedNames = scores.GroupBy(s => s.StudentIdentifier)
                                     .Join(reversedNames, g => g.Key, r => r.ReversedName, (g1, g2) => new { Student = g1, Reversed = g2 })
                                     .Where(g => g.Student.Count() > g.Reversed.Count)
                                     .Select(g => new
                                     {
                                         StudentIdentifier = g.Student.Key,
                                         PossibleMatches = g.Reversed.PossibleMatches
                                     })
                                     .ToArray();

            var dashReplacedWithSpace = scores.GroupBy(s => s.FirstName.Replace('-', ' ')).Select(g => new { NoDashes = g.Key, Count = g.Count(), PossibleMatches = g.ToArray() });
            var lostDashes = scores.GroupBy(s => s.StudentIdentifier)
                                     .Join(dashReplacedWithSpace, g => g.Key, r => r.NoDashes, (g1, g2) => new { Student = g1, NoDashes = g2 })
                                     .Where(g => g.Student.Count() > g.NoDashes.Count)
                                     .Select(g => new
                                     {
                                         StudentIdentifier = g.Student.Key,
                                         PossibleMatches = g.NoDashes.PossibleMatches
                                     })
                                     .ToArray();

            // Validation: Find students who have more than one class standing or emphasis in a given semester.
            var multipleStandings = scores.GroupBy(s => new
                {
                    s.StudentIdentifier,
                    s.Semester
                })
                .Where(g => g.Select(x => x.ClassStanding).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1)
                .Select(g => new
                {
                    g.Key.StudentIdentifier,
                    g.Key.Semester,
                    Scores = g.ToArray()
                }).ToArray()
                .ToArray();
            var multipleEmphasis = scores.GroupBy(s => new
                {
                    s.StudentIdentifier,
                    s.Semester
                })
                .Where(g => g.Select(x => x.Emphasis).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1)
                .Select(g => new
                {
                    g.Key.StudentIdentifier,
                    g.Key.Semester,
                    Scores = g.ToArray()
                }).ToArray();

            var messages = new List<Tuple<string, string>>();
            foreach (var fn in onlyFirstNames)
            {
                var mostPopular = fn.FirstName + " " + fn.LastName;
                var mismatches = fn.MismatchScores.Select(s => new { FirstName = s.FirstName, s.Semester }).Distinct();
                foreach (var m in mismatches)
                {
                    var sortValue = m.Semester + "_" + mostPopular;
                    var message = $"Student {mostPopular} has potential misspelled first name {m.FirstName} in semester {m.Semester}; using {fn.FirstName}";
                    messages.Add(new Tuple<string, string>(sortValue, message));
                }
                foreach (var score in fn.MismatchScores)
                {
                    score.StudentIdentifier = mostPopular;
                }
            }
            foreach (var ln in onlyLastNames)
            {
                var mostPopular = ln.FirstName + " " + ln.LastName;
                var mismatches = ln.MismatchScores.Select(s => new { LastName = s.LastName, s.Semester }).Distinct();
                foreach (var m in mismatches)
                {
                    var sortValue = m.Semester + "_" + mostPopular;
                    var message = $"Student {mostPopular} has potential misspelled last name {m.LastName} in semester {m.Semester}; using {ln.LastName}";
                    messages.Add(new Tuple<string, string>(sortValue, message));
                }
                foreach (var score in ln.MismatchScores)
                {
                    score.StudentIdentifier = mostPopular;
                }
            }
            foreach (var f in flippedNames)
            {
                var mostPopular = f.StudentIdentifier;
                var mismatches = f.PossibleMatches.Select(s => new { s.StudentIdentifier, s.Semester }).Distinct();
                foreach (var m in mismatches)
                {
                    var sortValue = m.Semester + "_" + mostPopular;
                    var message = $"Student {mostPopular} has potential flipped first and last name {m.StudentIdentifier} in semester {m.Semester}";
                    messages.Add(new Tuple<string, string>(sortValue, message));
                }
                foreach (var score in f.PossibleMatches)
                {
                    score.StudentIdentifier = mostPopular;
                }
            }
            foreach (var ld in lostDashes)
            {
                var mostPopular = ld.StudentIdentifier;
                var mismatches = ld.PossibleMatches.Select(s => new { s.StudentIdentifier, s.Semester }).Distinct();
                foreach (var m in mismatches)
                {
                    var sortValue = m.Semester + "_" + mostPopular;
                    var message = $"Student {mostPopular} has potential missing dash {m.StudentIdentifier} in semester {m.Semester}";
                    messages.Add(new Tuple<string, string>(sortValue, message));
                }
                foreach (var score in ld.PossibleMatches)
                {
                    score.StudentIdentifier = mostPopular;
                }
            }
            foreach (var standing in multipleStandings)
            {
                var mostPopular = standing.Scores.GroupBy(s => s.ClassStanding).OrderByDescending(g => g.Count()).Select(g => g.Key).First();
                var sortValue = standing.Semester + "_" + standing.StudentIdentifier;
                var message = $"Student {standing.StudentIdentifier} has multiple class standings for semester {standing.Semester}; picking {mostPopular}";
                messages.Add(new Tuple<string, string>(sortValue, message));
                foreach (var score in standing.Scores)
                {
                    score.ClassStanding = mostPopular;
                }
            }
            foreach (var emphasis in multipleEmphasis)
            {
                var mostPopular = emphasis.Scores.GroupBy(s => s.Emphasis).OrderByDescending(g => g.Count()).Select(g => g.Key).First();
                var sortValue = emphasis.Semester + "_" + emphasis.StudentIdentifier;
                var message = $"Student {emphasis.StudentIdentifier} has multiple emphases for semester {emphasis.Semester}; picking {mostPopular}";
                messages.Add(new Tuple<string, string>(sortValue, message));
                foreach (var score in emphasis.Scores)
                {
                    score.Emphasis = mostPopular;
                }
            }
            return messages;
        }

        private static IEnumerable<AssessmentScore> GetSimilarNames(IGrouping<string, AssessmentScore> groupedScores, Func<AssessmentScore, string> getNameFunc, string compareToName)
        {
            var similarNames = groupedScores.Where(score =>
            {
                var name = getNameFunc(score);
                return name != compareToName
                           && (Soundex.Generate(name) == Soundex.Generate(compareToName)
                               || name.LevenshteinDistance(compareToName) <= _maximumLevenshteinDistance
                               || (name[0] == compareToName[0]
                                   && name != compareToName
                                   && (Math.Abs(name.Length - compareToName.Length) <= 3
                                       || name.Contains(compareToName)
                                       || compareToName.Contains(name)
                                       )
                                  )
                               );
            });

            return similarNames;
        }
    }
}
