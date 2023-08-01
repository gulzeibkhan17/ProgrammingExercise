using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ProgrammingExercise;

string path = Directory.GetCurrentDirectory();
string subAwardBudgetFiles = Path.Combine(path, "SubAwardBudgetFiles");
string[] excelFilePaths = Directory.GetFiles(subAwardBudgetFiles, "*.xlsx");
int fileNumb = 0;
List<SubrecipientData> list2 = new();
foreach (string filePath in excelFilePaths)
{
    List<SubrecipientData> list = new();
    fileNumb++;
    Console.WriteLine("\nSummary for file:" + fileNumb);
    list = ReadExcelFile.ReadSpecificTable(filePath);

    foreach (var item in list)
    {
        list2.Add(item);
    }
}
Console.WriteLine("\n /////////////////////////////////");
Console.WriteLine("\n Total Subaward Summary:");

Dictionary<string, double> kv = new Dictionary<string, double>();

foreach (var data in list2)
{
    string subrecipientName = data.SubrecipientName;
    double amount = data.Amount;

    if (kv.ContainsKey(subrecipientName))
    {
        kv[subrecipientName] += amount;
    }
    else
    {
        kv.Add(subrecipientName, amount);
    }
}
Console.WriteLine("\nDistinct Subaward recipients and their Total Subaward Amounts:");
foreach (var entry in kv)
{
    if (string.IsNullOrWhiteSpace(entry.Key))
    {
        continue;
    }
    Console.WriteLine($"{entry.Key}: Total Subaward Amount: {entry.Value}");
}
//foreach (var item in list2)
//{

//    Console.WriteLine($"{item.SubrecipientName}: Total Subaward Amount: {item.Amount}");
//}
Console.ReadKey();


