using ClosedXML.Excel;

namespace PostClosedXML
{
  class Program
  {
    static void Main(string[] args)
    {
      using (var workbook = new XLWorkbook())
      {
        var worksheet = workbook.Worksheets.Add("Sample Sheet");
        worksheet.Cell("A1").Value = "Hello World!";
        worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
        workbook.SaveAs("HelloWorld.xlsx");
      }
    }
  }
}
