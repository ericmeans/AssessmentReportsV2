using DuoVia.FuzzyStrings;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace AssessmentReportsV2
{
    public static class ScoreValidator
    {
        private const int _maximumLevenshteinDistance = 1;

        public static HashSet<Tuple<string, string>> ValidateScores(List<AssessmentScore> scores)
        {
            var messages = new HashSet<Tuple<string, string>>();

            // Validation: Find students whose first and last name got flipped, or a dash got added or removed.
            var reversedNames = scores.GroupBy(s => s.LastName + " " + s.FirstName).Select(g => new { ReversedName = g.Key, Count = g.Count(), PossibleMatches = g.ToArray() });
            var flippedNames = scores.GroupBy(s => s.StudentIdentifier)
                                     .Join(reversedNames, g => g.Key, r => r.ReversedName, (g1, g2) => new { Student = g1, Reversed = g2 })
                                     .Where(g => g.Student.Count() > g.Reversed.Count)
                                     .Select(g => new
                                     {
                                         StudentIdentifier = g.Student.Key,
                                         FirstName = g.Student.First().LastName,
                                         LastName = g.Student.First().FirstName,
                                         PossibleMatches = g.Reversed.PossibleMatches
                                     })
                                     .ToArray();

            var dashReplacedWithSpace = scores.GroupBy(s => s.FirstName?.Replace('-', ' ')).Select(g => new { NoDashes = g.Key, Count = g.Count(), PossibleMatches = g.ToArray() });
            var lostDashes = scores.GroupBy(s => s.StudentIdentifier)
                                     .Join(dashReplacedWithSpace, g => g.Key, r => r.NoDashes, (g1, g2) => new { Student = g1, NoDashes = g2 })
                                     .Where(g => g.Student.Count() > g.NoDashes.Count)
                                     .Select(g => new
                                     {
                                         StudentIdentifier = g.Student.Key,
                                         FirstName = g.Student.First().FirstName,
                                         LastName = g.Student.First().LastName,
                                         PossibleMatches = g.NoDashes.PossibleMatches
                                     })
                                     .ToArray();

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
                    score.FirstName = f.FirstName;
                    score.LastName = f.LastName;
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
                    score.FirstName = ld.FirstName;
                    score.LastName = ld.LastName;
                }
            }

            // Validation: Find students with misspelled names.
            var firstNames = scores.Select(score => score.FirstName).Distinct().Where(s => !string.IsNullOrWhiteSpace(s));
            var similarFirstNames = GetSupersets(firstNames.Select(s => GetSimilarNames(firstNames, s)));
            var lastNames = scores.Select(score => score.LastName).Distinct().Where(s => !string.IsNullOrWhiteSpace(s));
            var similarLastNames = GetSupersets(lastNames.Select(s => GetSimilarNames(lastNames, s)));

            var scoresCopy = new List<AssessmentScore>(scores);

            while (scoresCopy.Any())
            {
                var score = scoresCopy.First();
                var similarFirstList = similarFirstNames.Where(n => n.Contains(score.FirstName)).FirstOrDefault() ?? new SortedSet<string>(new[] { score.FirstName });
                var similarLastList = similarLastNames.Where(n => n.Contains(score.LastName)).FirstOrDefault() ?? new SortedSet<string>(new[] { score.LastName });

                if (similarFirstList.Count == 1 && similarLastList.Count == 1)
                {
                    scoresCopy.RemoveAt(0);
                    continue;
                }

                var similarScores = scoresCopy.Where(s =>
                        (s.FirstName.Equals(score.FirstName, StringComparison.CurrentCultureIgnoreCase)
                         || similarFirstList.Contains(s.FirstName, StringComparer.CurrentCultureIgnoreCase))
                        && (s.LastName.Equals(score.LastName, StringComparison.CurrentCultureIgnoreCase)
                         || similarLastList.Contains(s.LastName, StringComparer.CurrentCultureIgnoreCase))
                    ).ToArray();

                if (!similarScores.Any())
                {
                    scoresCopy.RemoveAt(0);
                    continue;
                }

                var firstName = similarScores.GroupBy(s => s.FirstName).OrderByDescending(g => g.Count()).First().Key;
                var lastName = similarScores.GroupBy(s => s.LastName).OrderByDescending(g => g.Count()).First().Key;
                var fullName = firstName + " " + lastName;

                if (score.FirstName != firstName || score.LastName != lastName)
                {
                    var sortValue = score.Semester + "_" + fullName;
                    var message = $"Student {fullName} has potential misspelled name {score.FirstName} {score.LastName} in semester {score.Semester}; using {fullName}";
                    messages.Add(new Tuple<string, string>(sortValue, message));
                    score.FirstName = firstName;
                    score.LastName = lastName;
                    score.StudentIdentifier = fullName;
                    scoresCopy.RemoveAt(0);
                }

                foreach (var s in similarScores)
                {
                    scoresCopy.Remove(s);
                    if (s.FirstName == firstName && s.LastName == lastName)
                        continue;

                    var sortValue = s.Semester + "_" + fullName;
                    var message = $"Student {fullName} has potential misspelled name {s.FirstName} {s.LastName} in semester {s.Semester}; using {fullName}";
                    messages.Add(new Tuple<string, string>(sortValue, message));
                    s.FirstName = firstName;
                    s.LastName = lastName;
                    s.StudentIdentifier = fullName;
                }
            }

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
                                              })
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
                                             })
                                         .ToArray();

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

        private static SortedSet<string> GetSimilarNames(IEnumerable<string> possibleNames, string compareToName)
        {
            var similarNames = possibleNames.Where(name =>
            {
                return name == compareToName
                       || Soundex.Generate(name) == Soundex.Generate(compareToName)
                       || name.LevenshteinDistance(compareToName) <= _maximumLevenshteinDistance
                       || name.StartsWith(compareToName, StringComparison.OrdinalIgnoreCase)
                       || compareToName.StartsWith(name, StringComparison.OrdinalIgnoreCase)
                       || name.EndsWith(compareToName, StringComparison.OrdinalIgnoreCase)
                       || compareToName.EndsWith(name, StringComparison.OrdinalIgnoreCase);
            });

            return new SortedSet<string>(similarNames);
        }

        private static IEnumerable<SortedSet<string>> GetSupersets(IEnumerable<SortedSet<string>> sets)
        {
            // Sort by length, then copy only entries that aren't a subset of any other set.
            var supersets = new List<SortedSet<string>>();

            var orderedSets = sets.OrderBy(s => s.Count).ToList();
            for (var i = 0; i < orderedSets.Count; ++i)
            {
                var set = orderedSets[i];
                var hasSuper = false;
                for (var j = i + 1; j < orderedSets.Count(); ++j)
                {
                    if (set.IsSubsetOf(orderedSets[j]))
                    {
                        hasSuper = true;
                        break;
                    }
                }
                if (!hasSuper)
                    supersets.Add(set);
            }

            return supersets;
        }
    }
}
