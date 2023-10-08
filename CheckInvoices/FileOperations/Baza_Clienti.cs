using System;
using System.Collections.Generic;
using System.Text;

namespace CheckInvoices.FileOperations
{
    public class Baza_Clienti
    {
        public string Nume_client { get; set; }
        public string CUI { get; set; }
        public string Nr_factura { get; set; }
        public string Data_factura { get; set; }
        public string Cod_Produs { get; set; }
        public string Valoare_factura { get; set; }
    }
}
