using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace ProgrammingExercise
{
    public class ReadExcelFile
    {
        public static List<SubrecipientData> ReadSpecificTable(string filePath)
        {
            List<SubrecipientData> lst = new();
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                var workbook = new XSSFWorkbook(file);
                var sheet = workbook.GetSheetAt(0);
                int startRowIndex = FindStartRowIndex(sheet);
                if (startRowIndex != -1)
                {
                    var distinctSubrecipients = ReadTable(sheet, startRowIndex, filePath);
                    foreach (var data in distinctSubrecipients)
                    {
                        lst.Add(data);
                        if (string.IsNullOrWhiteSpace(data.SubrecipientName)){
                            continue;
                        }
                        Console.WriteLine($"{data.SubrecipientName}: Subaward Amount: {data.Amount}");
                    }
                }
                else
                {
                    Console.WriteLine("No data found.");
                }
                return lst;
            }
        }

        public static int FindStartRowIndex(ISheet sheet)
        {
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                var cellG = row.GetCell(0);
                var cellOtherDirectCost = row.GetCell(1);
                if (cellG != null && cellOtherDirectCost != null)
                {
                    var cell1Value = cellG.StringCellValue;
                    var cell2Value = cellOtherDirectCost.StringCellValue;

                    if (cell1Value.Equals("G.", StringComparison.OrdinalIgnoreCase) && cell2Value.Equals("Other Direct Costs", StringComparison.OrdinalIgnoreCase))
                    {
                        return rowIndex + 1;
                    }
                }
            }

            return -1;
        }

        public static List<SubrecipientData> ReadTable(ISheet sheet, int startRowIndex, string fileName)
        {
            var subrecipientData = new List<SubrecipientData>();
            for (int rowIndex = startRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row != null)
                {
                    for (int columnIndex = 0; columnIndex < row.LastCellNum; columnIndex++)
                    {
                        var cell = row.GetCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        if (cell.CellType == CellType.String && cell.StringCellValue.StartsWith("Subaward:"))
                        {
                            string subrecipientName = string.Empty;
                            double amount = 0;
                            if (fileName.EndsWith("1.xlsx"))
                            {
                                subrecipientName = row.GetCell(columnIndex + 1).StringCellValue;
                                amount = row.GetCell(columnIndex + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue;
                                subrecipientData.Add(new SubrecipientData(subrecipientName, amount));
                            }
                            else if (fileName.EndsWith("2.xlsx"))
                            {
                                subrecipientName = cell.StringCellValue.Replace("Subaward: ", "");
                                amount = row.GetCell(columnIndex + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue;
                                subrecipientData.Add(new SubrecipientData(subrecipientName, amount));
                            }
                            else
                            {
                                subrecipientName = row.GetCell(columnIndex + 1).StringCellValue;
                                amount = row.GetCell(columnIndex + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue;
                                subrecipientData.Add(new SubrecipientData(subrecipientName, amount));
                            }
                        }
                    }
                }
            }
            return subrecipientData;
        }
    }
}
