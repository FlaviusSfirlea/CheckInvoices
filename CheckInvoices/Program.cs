using System;
using System.Threading.Tasks;
using CheckInvoices.FileOperations;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using MimeKit;
using System.Threading;
using System.Linq;

namespace CheckInvoices
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            IConfiguration configuration = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json")
            .Build();

            var appSettings = new AppSettings();
            configuration.GetSection("AppSettings").Bind(appSettings);

            string invoicesFolder = appSettings.InvoicesFolder;
            string bazaClientiFolder = appSettings.BazaClientiFolder;

            string gmailImapServer = "imap.gmail.com";
            int port = 993;
            string email = "verificarefacturi@gmail.com";
            string password = "eofi miwk trem ozds";

            using (var client = new ImapClient())
            {
                client.Connect(gmailImapServer, port, SecureSocketOptions.SslOnConnect);
                client.Authenticate(email, password);

                var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadOnly);

                // Create a new thread to periodically check for new email notifications
                var listenerThread = new Thread(() =>
                {
                    while (true)
                    {
                        // Check for new messages every 60 seconds.
                        Thread.Sleep(60000);

                        var uids = inbox.Search(SearchQuery.All);

                        uids = uids.OrderByDescending(uid => inbox.GetMessage(uid).Date).ToList();

                        int count = uids.Count;

                        var mostRecentUid = uids.FirstOrDefault();

                        if (mostRecentUid != null)
                        {
                            var message = inbox.GetMessage(mostRecentUid);
                            var subject = message.Subject;
                            var sender = message.From.ToString();

                            if (subject == "Verificare Facturi")
                            {
                                InvoiceChecker.CheckingInvoices(invoicesFolder, bazaClientiFolder);
                                EmailSender.SendEmail(sender);
                            }
                        }
                    }
                });

                listenerThread.Start();

                Console.WriteLine("Email listener started. Press Enter to exit.");
                Console.ReadLine();

                listenerThread.Join();
                client.Disconnect(true);
            }
        }
    }
}
