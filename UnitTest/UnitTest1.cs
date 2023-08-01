using Microsoft.VisualStudio.TestPlatform.TestHost;
using ProgrammingExercise;

namespace UnitTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestSubrecipientsInFile1()
        {
            string path = Directory.GetCurrentDirectory();
            string subAwardBudgetFiles = Path.Combine(path, "SubAwardBudgetFiles");
            string[] excelFilePaths = Directory.GetFiles(subAwardBudgetFiles, "*.xlsx");
            string testFilePath = excelFilePaths[0];
            List<string> expectedSubrecipients = new List<string> { "Indiana", "Mayo", "Purdue", "Florida" };
            Program program = new();
            List<SubrecipientData> result = ReadExcelFile.ReadSpecificTable(testFilePath);
            Assert.IsNotNull(result);
            Assert.AreEqual(expectedSubrecipients.Count, result.Count);

            foreach (var subrecipient in expectedSubrecipients)
            {
                Assert.IsTrue(result.Any(data => data.SubrecipientName == subrecipient));
            }
        }

    }
}