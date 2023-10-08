using System;
using System.Collections.Generic;
using System.Text;

namespace CheckInvoices
{
    class AppSettings
    {
        public string InvoicesFolder { get; set; }
        public string BazaClientiFolder { get; set; }
        public string Email { get; set; }
        public string EmailPassword { get; set; }
    }
}
