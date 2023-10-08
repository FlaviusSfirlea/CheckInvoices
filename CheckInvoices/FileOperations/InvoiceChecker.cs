using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CheckInvoices.FileOperations
{
    public static class InvoiceChecker
    {
        public static void CheckingInvoices(string invoicesFolder, string bazaClientiFolder)
        {
            if (!Directory.Exists(invoicesFolder))
            { 
                Directory.CreateDirectory(invoicesFolder);
                Console.WriteLine("Directory " + invoicesFolder + " created. Please copy invoices to the folder.");
            }
            if (!Directory.Exists(bazaClientiFolder))
            {
                Directory.CreateDirectory(bazaClientiFolder);
                Console.WriteLine("Directory " + bazaClientiFolder + " created. Please copy invoices to the folder.");
            }

            string bazaClientiFile = Path.Combine(bazaClientiFolder, "Baza_clienti.xlsx");

            string rezultatFileName = Path.Combine(invoicesFolder, "Rezultat_VerificareFacturi_" + DateTime.Now.ToString("dd.MM.yy") + ".xlsx");

            //delete Rezultat_VerificareFacturi if already existing
            if (File.Exists(rezultatFileName))
                File.Delete(rezultatFileName);

            var columnNames = Excel.GenerateExcelDictionary(bazaClientiFile);
            var allData = Excel.ReadExcel(bazaClientiFile);

            int keyCUI = Excel.GetColumnIndexByName("CUI", columnNames);

            var CUIValues = Excel.GetColumnData(allData, keyCUI);

            for (int i = 0; i < CUIValues.Count; i++)
            {
                string PDFfile = Path.Combine(invoicesFolder, CUIValues[i] + ".pdf");
                if (File.Exists(PDFfile))
                {
                    PDFExtractedData pDFExtractedData = PDF.ExtractPDFData(PDFfile);
                    //write into Baza_clienti.xlsx if codProduse is not empty
                    if(!String.IsNullOrEmpty(pDFExtractedData.codProduse.Trim()))
                    {
                        int Keycod_produs = Excel.GetColumnIndexByName("Cod_Produs", columnNames);
                        Excel.WriteToExcelByColumnByRow(pDFExtractedData.codProduse, i + 1, Keycod_produs, bazaClientiFile);
                    }
                    //get data per row number from Baza_Clienti
                    Baza_Clienti bazaClienti = Excel.Baza_ClientyByRowNumber(i + 2, bazaClientiFile);

                    //comparing PDF and Baza_Clienti data and writing to Rezultat_Verificare
                    CompareExcelToPDF(pDFExtractedData, bazaClienti, rezultatFileName, i+1);
                }
                else
                {
                    Console.Out.WriteLine("PDF file " + PDFfile + " does not exist");
                }
            }
        }
        public static void CompareExcelToPDF(PDFExtractedData pDFExtractedData, Baza_Clienti bazaClienti, string rezultatFileName, int rowIndex)
        {
            if(!File.Exists(rezultatFileName))
                Excel.CreateRezultatVerificari(rezultatFileName);

            var rezultatKeys = Excel.GenerateExcelDictionary(rezultatFileName);

            Console.Out.WriteLine($"Writing into {rezultatFileName}");

            //write Nume_client
            int keyNume_Client = Excel.GetColumnIndexByName("Nume_client", rezultatKeys);
            Excel.WriteToExcelByColumnByRow(bazaClienti.Nume_client, rowIndex, keyNume_Client, rezultatFileName);
            //write Denumire_client
            int keyDenumire_Client = Excel.GetColumnIndexByName("Denumire_client", rezultatKeys);
            Excel.WriteToExcelByColumnByRow(pDFExtractedData.Client, rowIndex, keyDenumire_Client, rezultatFileName);
            //write CUI
            int keyCUI = Excel.GetColumnIndexByName("CUI", rezultatKeys);
            Excel.WriteToExcelByColumnByRow(bazaClienti.CUI, rowIndex, keyCUI, rezultatFileName);
            //write Nr_factura
            int keyNr_factura = Excel.GetColumnIndexByName("Nr_factura", rezultatKeys);
            Excel.WriteToExcelByColumnByRow(bazaClienti.Nr_factura, rowIndex, keyNr_factura, rezultatFileName);
            //write Data_factura
            int keyData_factura = Excel.GetColumnIndexByName("Data_factura", rezultatKeys);
            Excel.WriteToExcelByColumnByRow(bazaClienti.Data_factura, rowIndex, keyData_factura, rezultatFileName);
            //write Nr_factura_data
            int keyNr_factura_data = Excel.GetColumnIndexByName("Nr_factura_data", rezultatKeys);
            Excel.WriteToExcelByColumnByRow(pDFExtractedData.numberDate, rowIndex, keyNr_factura_data, rezultatFileName);
            //write Cod_Produs
            int keyCod_Produs = Excel.GetColumnIndexByName("Cod_Produs", rezultatKeys);
            Excel.WriteToExcelByColumnByRow(pDFExtractedData.codProduse, rowIndex, keyCod_Produs, rezultatFileName);
            //write Valoare_Factura
            int keyValoare_Factura = Excel.GetColumnIndexByName("Valoare_factura", rezultatKeys);
            Excel.WriteToExcelByColumnByRow(bazaClienti.Valoare_factura, rowIndex, keyValoare_Factura, rezultatFileName);

            //write Observatii
            string observatii = String.Empty;

            if (bazaClienti.Nume_client.Trim() != pDFExtractedData.Client.Trim())
                observatii += "Nume_client diferit pe factura";
            //check client
            if(bazaClienti.Nr_factura.Trim() != pDFExtractedData.Number.Trim() || bazaClienti.Data_factura.Trim() != pDFExtractedData.Date.Trim())
            {
                if (!String.IsNullOrEmpty(observatii))
                    observatii += ";";

                observatii += "nr si/ sau data diferita pe factura";
            }
            //check total
            if(bazaClienti.Valoare_factura.Trim() != pDFExtractedData.TotalPayment.Trim())
            {
                if (!String.IsNullOrEmpty(observatii))
                    observatii += ";";

                observatii += "Total de plata diferit pe factura";
            }
            //check if it is signed
            if(!pDFExtractedData.isSigned)
            {
                if (!String.IsNullOrEmpty(observatii))
                    observatii += ";";

                observatii += "Factura nesemnata";
            }
            
            //write Observatii if not Empty
            if(!String.IsNullOrEmpty(observatii))
            {
                int keyObservatii = Excel.GetColumnIndexByName("Observatii", rezultatKeys);
                Excel.WriteToExcelByColumnByRow(observatii, rowIndex, keyObservatii, rezultatFileName);
            }
        }

    }
}
