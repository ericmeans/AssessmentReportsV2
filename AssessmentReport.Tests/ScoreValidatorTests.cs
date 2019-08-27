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
            Assert.AreEqual("Student Alison Means has potential misspelled first name Allison in semester Spring 2019; using Alison", validationResult.First().Item2);
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
                    LastName = "Albrecht",
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
            Assert.AreEqual(1, validationResult.Count);
            Assert.AreEqual("Student Eric Albrecht has potential misspelled last name Albrect in semester Spring 2019; using Albrecht", validationResult.First().Item2);
        }
    }
}