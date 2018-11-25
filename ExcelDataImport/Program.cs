using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json.Converters;
using System.Configuration;
using System.Data.SqlClient;

namespace ExcelDataImport
{
    class ExcelStartingCells
    {
        public string FileName { get; set; }
        public Dictionary<string,string> StartingCells { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            var json = File.ReadAllText("StartingCells.json");
            List<ExcelStartingCells> excelStartingCells = JsonConvert.DeserializeObject<List<ExcelStartingCells>>(json, new KeyValuePairConverter()); // OK
            string connectionString = ConfigurationManager.ConnectionStrings["default"].ConnectionString;

            foreach (var excelStartingCell in excelStartingCells)
            {
                string path = Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, string.Format(@"Data\{0}", excelStartingCell.FileName));
                string[] sheetNames = excelStartingCell.StartingCells.Select(kvp => kvp.Key).ToArray();

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, false))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                        foreach (var sheetStartingCell in excelStartingCell.StartingCells)
                        {
                            var sheetId = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(w => w.Name == sheetStartingCell.Key)?.Id;
                            if (sheetId != null && sheetId.HasValue)
                            {
                                WorksheetPart worksheetPart = workbookPart.GetPartById(sheetId.Value) as WorksheetPart;
                                bool startingCellFound = false;
                                Console.WriteLine(string.Format("Traitement de la feuille {0} ...", sheetStartingCell.Key));
                                foreach (SheetData sheetData in worksheetPart.Worksheet.Elements<SheetData>())
                                {
                                    string text;
                                    int rowCount = sheetData.Elements<Row>().Count();

                                    foreach (Row r in sheetData.Elements<Row>())
                                    {
                                        foreach (Cell c in r.Elements<Cell>())
                                        {
                                            if (c.CellReference == sheetStartingCell.Value)
                                                startingCellFound = true;
                                            if (startingCellFound)
                                            {
                                                text = c.CellValue.InnerText;
                                                if (!string.IsNullOrWhiteSpace(text))
                                                {
                                                    if (text.Length > 150 || text.Contains("insert into"))
                                                    {
                                                        SqlCommand command = new SqlCommand(text, conn);
                                                        try
                                                        {
                                                            command.ExecuteNonQuery();
                                                        }
                                                        catch (Exception exc)
                                                        {
                                                            Console.WriteLine(string.Format("Une erreur s'est produite. Cellule {0}, feuille {1}. {2}",
                                                                c.CellReference, sheetStartingCell.Key, exc.Message));
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    
                                }
                            }
                        }
                    }
                    conn.Close();
                }
            }



            Console.WriteLine("Traitement terminé.");
            Console.ReadKey();
        }
    }
}
