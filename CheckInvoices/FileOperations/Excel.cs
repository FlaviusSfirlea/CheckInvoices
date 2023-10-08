using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CheckInvoices.FileOperations
{
    public static class Excel
    {
        public static List<string[]> ReadExcel(string excelFilePath)
        {
            try
            { 
            // Check if the Excel file exists
            if (File.Exists(excelFilePath))
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    if (worksheet != null)
                    {
                        int lastRow = worksheet.Dimension.End.Row;

                        var data = new List<string[]>();

                        for (int row = 2; row <= lastRow; row++)
                        {
                            var rowData = worksheet.Cells[row, 1, row, 6].Select(cell => cell.Text).ToArray();
                            data.Add(rowData);
                        }

                        return data;
                    }
                    else
                    {
                        Console.WriteLine("Worksheet not found in the Excel file.");
                    }
                }
            }
            else
            {
                Console.WriteLine("Excel file not found at the specified path.");
            }
            }
            catch(Exception ex)
            {
                return null; // Return null if there was an error
            }

            return null; // Return null if there was an error
        }

        public static List<Dictionary<string, List<object>>> GenerateExcelDictionary(string excelFilePath)
        {
            try 
            { 
                List<Dictionary<string, List<object>>> myExcelContainer = new List<Dictionary<string, List<object>>>();

                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    if (worksheet != null)
                    {
                        var columnMax = worksheet.Dimension.End.Column;
                        var rowMax = worksheet.Dimension.End.Row;

                        for (int column = 1; column <= columnMax; column++)
                        {
                            var columnName = worksheet.Cells[1, column].Text.Replace(" ", string.Empty);
                            var columnValues = new List<object>();

                            for (int row = 2; row <= rowMax; row++)
                            {
                                var cellValue = worksheet.Cells[row, column].Text;
                                columnValues.Add(cellValue);
                            }

                            var columnData = new Dictionary<string, List<object>>();
                            columnData.Add(columnName, columnValues);
                            myExcelContainer.Add(columnData);
                        }
                    }
                }
                return myExcelContainer;
            }
            catch(Exception ex)
            {
                return null;
            }
        }
        public static List<string> GetColumnData(List<string[]> allData, int columnIndex)
        {
            var columnData = new List<string>();

            foreach (var rowData in allData)
            {
                if (rowData.Length > columnIndex)
                {
                    columnData.Add(rowData[columnIndex]);
                }
                else
                {
                    // If the column index is out of range for this row, add an empty string or handle accordingly
                    columnData.Add("");
                }
            }

            return columnData;
        }
        public static int GetColumnIndexByName(string columnName, List<Dictionary<string, List<object>>> columnNames)
        {
            for (int columnIndex = 0; columnIndex < columnNames.Count; columnIndex++)
            {
                if (columnNames[columnIndex].ContainsKey(columnName))
                {
                    return columnIndex;
                }
            }
            return -1; // Return -1 if the column name is not found
        }
        public static void WriteToExcelByColumnByRow(string textToWrite, int rowNumber, int columnIndex, string filePath)
        {
            try
            {
                using (var excelPackage = new ExcelPackage(new System.IO.FileInfo(filePath)))
                {
                    var worksheet = excelPackage.Workbook.Worksheets[0]; 

                    // Check if the row exists (create if necessary)
                    var row = worksheet.Cells[rowNumber, columnIndex + 1].Start.Row;
                    if (worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Any(c => !string.IsNullOrEmpty(c.Text)))
                    {
                        row++;
                    }
                    worksheet.Cells[row, columnIndex + 1].Value = textToWrite;
                    excelPackage.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing to Excel: {ex.Message}");
            }
        }
        public static Baza_Clienti Baza_ClientyByRowNumber(int rowNumber, string filePath)
        {
            try
            {
                using (var excelPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = excelPackage.Workbook.Worksheets[0]; 

                    var bazaClienti = new Baza_Clienti
                    {
                        Nume_client = worksheet.Cells[rowNumber, 1].Text,
                        CUI = worksheet.Cells[rowNumber, 2].Text,
                        Nr_factura = worksheet.Cells[rowNumber, 3].Text,
                        Data_factura = (worksheet.Cells[rowNumber, 4].Text),
                        Cod_Produs = (worksheet.Cells[rowNumber, 5].Text),
                        Valoare_factura = (worksheet.Cells[rowNumber, 6].Text),
                    };

                    return bazaClienti;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading from Excel: {filePath} {ex.Message}");
            }

            return null; // Return null if data is not found or cannot be parsed
        }
        public static void CreateRezultatVerificari(string fileName)
        {
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    string[] headers = {
                    "Nume_client",
                    "Denumire_client",
                    "CUI",
                    "Nr_factura",
                    "Data_factura",
                    "Nr_factura_data",
                    "Cod_Produs",
                    "Valoare_factura",
                    "Observatii"
                };

                    for (int i = 0; i < headers.Length; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = headers[i];
                    }


                    File.WriteAllBytes(fileName, package.GetAsByteArray());

                    Console.WriteLine($"Excel file created at: {fileName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating Excel file: {fileName} {ex.Message}");
            }
        }
    }
}
