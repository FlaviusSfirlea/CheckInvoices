using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace CheckInvoices
{
    public static class EmailSender
    {
        public static void SendEmail(string toEmail)
        {
            try
            {
                string smtpServer = "smtp.gmail.com";
                int smtpPort = 587;
                string smtpUsername = "verificarefacturi@gmail.com";
                string smtpPassword = "eofi miwk trem ozds";

                // Create a new SmtpClient instance
                using (var client = new SmtpClient(smtpServer, smtpPort))
                {
                    client.EnableSsl = true;
                    client.Credentials = new NetworkCredential(smtpUsername, smtpPassword);

                    // Create a new MailMessage
                    using (var message = new MailMessage())
                    {
                        message.From = new MailAddress(smtpUsername);
                        message.To.Add(toEmail);
                        message.Subject = "Verificare facturi";
                        message.Body = "Buna ziua,\n\nProcesul a fost finalizat cu succes. Se poate verifica rezultatul accesand sharefolderul dedicat.\n\nSpor,\n";
                        message.IsBodyHtml = false; // Change to true if you want HTML body

                        // Send the email
                        client.Send(message);
                    }
                }

                Console.WriteLine($"Email sent to {toEmail} successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending email: {ex.Message}");
            }
        }
    }
}
