using System;
using System.Collections.Generic;
using System.Text;

namespace CheckInvoices.FileOperations
{
    public class PDFExtractedData
    {
        public string Client { get; set; }
        public string Number { get; set; }
        public string Date { get; set; }
        public string numberDate { get; set; }
        public string TotalPayment { get; set; }
        public string codProduse { get; set; }
        public bool isSigned { get; set; }
    }
}
