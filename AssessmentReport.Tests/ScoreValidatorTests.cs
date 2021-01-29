using AssessmentReportsV2;
using NUnit.Framework;
using System.Collections.Generic;
using System.Linq;

namespace Tests
{
    public class ScoreValidatorTests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void FirstNameMissingCharacter()
        {
            var scores = new List<AssessmentScore>
            {
                new AssessmentScore
                {
                    FirstName = "Allison",
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Allison Means",
                    ScoreName = "Oral",
                    Score = 4
                },
                new AssessmentScore
                {
                    FirstName = "Alison",
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Alison Means",
                    ScoreName = "Oral",
                    Score = 3
                },
                new AssessmentScore
                {
                    FirstName = "Alison",
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Alison Means",
                    ScoreName = "Oral",
                    Score = 4
                }
            };

            var validationResult = ScoreValidator.ValidateScores(scores);
            Assert.AreEqual(1, validationResult.Count);
            Assert.IsFalse(scores.Any(s => s.FirstName != "Alison"));
        }

        [Test]
        public void LastNameMissingCharacter()
        {
            var scores = new List<AssessmentScore>
            {
                new AssessmentScore
                {
                    FirstName = "Eric",
                    LastName = "Albrecht",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Eric Albrecht",
                    ScoreName = "Oral",
                    Score = 4
                },
                new AssessmentScore
                {
                    FirstName = "Eric",
                    LastName = "ALbrecht",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Eric Albrecht",
                    ScoreName = "Oral",
                    Score = 3
                },
                new AssessmentScore
                {
                    FirstName = "Eric",
                    LastName = "Albrecht",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Eric Albrecht",
                    ScoreName = "Oral",
                    Score = 4
                },
                new AssessmentScore
                {
                    FirstName = "Eric",
                    LastName = "Albrect",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Eric Albrect",
                    ScoreName = "Oral",
                    Score = 4
                }
            };

            var validationResult = ScoreValidator.ValidateScores(scores);
            Assert.AreEqual(2, validationResult.Count);
            Assert.IsFalse(scores.Any(s => s.LastName != "Albrecht"));
        }

        [Test]
        public void WoodAndSnow()
        {
            var scores = new List<AssessmentScore>
            {
                new AssessmentScore
                {
                    FirstName = "Eric",
                    LastName = "Snow",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Eric Snow",
                    ScoreName = "Oral",
                    Score = 4
                },
                new AssessmentScore
                {
                    FirstName = "Eric",
                    LastName = "Snow",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Eric Snow",
                    ScoreName = "Oral",
                    Score = 3
                },
                new AssessmentScore
                {
                    FirstName = "Eric",
                    LastName = "Wood",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Eric Wood",
                    ScoreName = "Oral",
                    Score = 4
                }
            };

            var validationResult = ScoreValidator.ValidateScores(scores);
            Assert.AreEqual(0, validationResult.Count);
            Assert.IsTrue(scores.Any(s => s.LastName == "Wood"));
            Assert.IsTrue(scores.Any(s => s.LastName == "Snow"));
        }

        [Test]
        public void Plural()
        {
            var scores = new List<AssessmentScore>
            {
                new AssessmentScore
                {
                    FirstName = "Eric",
                    LastName = "Wood",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Eric Wood",
                    ScoreName = "Oral",
                    Score = 4
                },
                new AssessmentScore
                {
                    FirstName = "Eric",
                    LastName = "Wood",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Eric Wood",
                    ScoreName = "Oral",
                    Score = 3
                },
                new AssessmentScore
                {
                    FirstName = "Eric",
                    LastName = "Woods",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Eric Woods",
                    ScoreName = "Oral",
                    Score = 4
                }
            };

            var validationResult = ScoreValidator.ValidateScores(scores);
            Assert.AreEqual(1, validationResult.Count);
            Assert.IsFalse(scores.Any(s => s.LastName != "Wood"));
        }

        [Test]
        public void FirstNameIsNull()
        {
            var scores = new List<AssessmentScore>
            {
                new AssessmentScore
                {
                    FirstName = null,
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Allison Means",
                    ScoreName = "Oral",
                    Score = 4
                },
                new AssessmentScore
                {
                    FirstName = "Alison",
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Alison Means",
                    ScoreName = "Oral",
                    Score = 3
                },
                new AssessmentScore
                {
                    FirstName = "Alison",
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Alison Means",
                    ScoreName = "Oral",
                    Score = 4
                }
            };

            var validationResult = ScoreValidator.ValidateScores(scores);
            Assert.AreEqual(0, validationResult.Count);
            Assert.IsTrue(scores.Any(s => s.FirstName == null));
        }

        [Test]
        public void FirstNameIsEmpty()
        {
            var scores = new List<AssessmentScore>
            {
                new AssessmentScore
                {
                    FirstName = "",
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Allison Means",
                    ScoreName = "Oral",
                    Score = 4
                },
                new AssessmentScore
                {
                    FirstName = "Alison",
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Alison Means",
                    ScoreName = "Oral",
                    Score = 3
                },
                new AssessmentScore
                {
                    FirstName = "Alison",
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Alison Means",
                    ScoreName = "Oral",
                    Score = 4
                }
            };

            var validationResult = ScoreValidator.ValidateScores(scores);
            Assert.AreEqual(0, validationResult.Count);
            Assert.IsTrue(scores.Any(s => s.FirstName == ""));
        }

        [Test]
        public void LastNameIsNull()
        {
            var scores = new List<AssessmentScore>
            {
                new AssessmentScore
                {
                    FirstName = "Alison",
                    LastName = null,
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Allison Means",
                    ScoreName = "Oral",
                    Score = 4
                },
                new AssessmentScore
                {
                    FirstName = "Alison",
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Alison Means",
                    ScoreName = "Oral",
                    Score = 3
                },
                new AssessmentScore
                {
                    FirstName = "Alison",
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Alison Means",
                    ScoreName = "Oral",
                    Score = 4
                }
            };

            var validationResult = ScoreValidator.ValidateScores(scores);
            Assert.AreEqual(0, validationResult.Count);
            Assert.IsTrue(scores.Any(s => s.LastName == null));
        }

        [Test]
        public void LastNameIsEmpty()
        {
            var scores = new List<AssessmentScore>
            {
                new AssessmentScore
                {
                    FirstName = "Alison",
                    LastName = "",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Allison Means",
                    ScoreName = "Oral",
                    Score = 4
                },
                new AssessmentScore
                {
                    FirstName = "Alison",
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Alison Means",
                    ScoreName = "Oral",
                    Score = 3
                },
                new AssessmentScore
                {
                    FirstName = "Alison",
                    LastName = "Means",
                    ClassStanding = "4 - Senior",
                    Emphasis = "Generalist",
                    Semester = "Spring 2019",
                    SemesterSort = AssessmentAnalyzer.ConvertToSortableValue("Spring 2019"),
                    StudentIdentifier = "Alison Means",
                    ScoreName = "Oral",
                    Score = 4
                }
            };

            var validationResult = ScoreValidator.ValidateScores(scores);
            Assert.AreEqual(0, validationResult.Count);
            Assert.IsTrue(scores.Any(s => s.LastName == ""));
        }
    }
}