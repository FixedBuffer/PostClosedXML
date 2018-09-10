using ClosedXML.Excel;
using System.IO;
using System.Linq;

namespace PostClosedXML
{
  class Program
  {
    static void Main(string[] args)
    {
      //Si no existe el fichero xlsx, pasamos a la primera parte, y sino a la segunda
      if (!File.Exists("HelloWorld.xlsx"))
        PrimeraParte();
      else
        SegundaParte();

    }

    static void PrimeraParte()
    {
      using (var workbook = new XLWorkbook())
      {
        var worksheet = workbook.Worksheets.Add("Sample Sheet");
        worksheet.Cell("A1").Value = "Hello World!";
        worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
        workbook.SaveAs("HelloWorld.xlsx");
      }
    }

    static void SegundaParte()
    {
      using (var workbook = new XLWorkbook("HelloWorld.xlsx"))
      {
        //Buscamos con LinQ la hohja que nos interesa copiar
        var SampleSheet = workbook.Worksheets.Where(x => x.Name == "Sample Sheet").First();
        //Añadimos una hoja nuevo
        var worksheet = workbook.Worksheets.Add("FixedBuffer");
        //Copiamos los valores
        worksheet.Cell("A1").Value = SampleSheet.Cell("A1").GetString().ToUpper();
        worksheet.Cell("A2").FormulaA1 = SampleSheet.Cell("A2").FormulaA1;
        //Guardamos el libro
        workbook.Save();
      }
    }
  }
}
