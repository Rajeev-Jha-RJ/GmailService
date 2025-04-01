using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;

namespace ImapOAuth2EmailSender
{
    class Program
    {
        // Gmail API scopes for IMAP and SMTP access
        static string[] Scopes = { "https://mail.google.com/" };
        static string ApplicationName = "Gmail IMAP OAuth2 Email Sender";

        static async Task Main(string[] args)
        {
            try
            {
                // Get OAuth 2.0 token
                var oauth2Token = await GetOAuth2TokenAsync();

                // Example attachments list - replace with your actual files
                List<string> attachments = new List<string>
                {
                    "file1.pdf",
                    "image.jpg",
                    "document.docx"
                };
                
                // Send email using SMTP with OAuth 2.0 and attachments
                await SendEmailWithOAuth2Async(oauth2Token, attachments);

                // Optionally, save a copy to the Sent folder via IMAP
                await SaveToSentFolderAsync(oauth2Token, attachments);

                Console.WriteLine("Email sent successfully and saved to Sent folder!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
                }
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static async Task<string> GetOAuth2TokenAsync()
        {
            UserCredential credential;

            // Load client secrets
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time
                string credPath = "token.json";
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true));

                Console.WriteLine($"Credential file saved to: {credPath}");
            }

            // Get the access token
            var token = await credential.GetAccessTokenForRequestAsync();
            return token;
        }

        private static async Task SendEmailWithOAuth2Async(string oauth2Token, List<string> attachmentPaths = null)
        {
            // Create a new message
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Your Name", "your.email@gmail.com"));
            message.To.Add(new MailboxAddress("Recipient Name", "recipient@example.com"));
            message.Subject = "Test email from IMAP with OAuth 2.0";

            // Create a multipart message body
            var multipart = new Multipart("mixed");
            
            // Add the plain text part
            var textPart = new TextPart("plain")
            {
                Text = "This is a test email sent using IMAP with OAuth 2.0 in .NET"
            };
            multipart.Add(textPart);
            
            // Add attachments if provided
            if (attachmentPaths != null && attachmentPaths.Count > 0)
            {
                foreach (var attachmentPath in attachmentPaths)
                {
                    if (File.Exists(attachmentPath))
                    {
                        // Create the attachment
                        var attachment = new MimePart()
                        {
                            Content = new MimeContent(File.OpenRead(attachmentPath)),
                            ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                            ContentTransferEncoding = ContentEncoding.Base64,
                            FileName = Path.GetFileName(attachmentPath)
                        };
                        
                        // Determine content type based on file extension
                        attachment.ContentType.MimeType = GetMimeType(Path.GetExtension(attachmentPath));
                        
                        // Add to multipart
                        multipart.Add(attachment);
                    }
                }
            }
            
            // Set the message body to our multipart
            message.Body = multipart;

            // Send the message
            using (var client = new SmtpClient())
            {
                // Connect to Gmail's SMTP server
                await client.ConnectAsync("smtp.gmail.com", 587, SecureSocketOptions.StartTls);

                // Authenticate using OAuth 2.0
                var oauth2 = new SaslMechanismOAuth2("your.email@gmail.com", oauth2Token);
                await client.AuthenticateAsync(oauth2);

                // Send the message
                await client.SendAsync(message);

                // Disconnect
                await client.DisconnectAsync(true);
            }
        }
        
        private static string GetMimeType(string extension)
        {
            // A simple mapping of common file extensions to MIME types
            switch (extension.ToLower())
            {
                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";
                case ".png":
                    return "image/png";
                case ".gif":
                    return "image/gif";
                case ".pdf":
                    return "application/pdf";
                case ".doc":
                case ".docx":
                    return "application/msword";
                case ".xls":
                case ".xlsx":
                    return "application/vnd.ms-excel";
                case ".txt":
                    return "text/plain";
                case ".zip":
                    return "application/zip";
                default:
                    return "application/octet-stream";
            }
        }

        private static async Task SaveToSentFolderAsync(string oauth2Token, List<string> attachmentPaths = null)
        {
            // Create a copy of the message for the Sent folder
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Your Name", "your.email@gmail.com"));
            message.To.Add(new MailboxAddress("Recipient Name", "recipient@example.com"));
            message.Subject = "Test email from IMAP with OAuth 2.0";
            
            // Create a multipart message body
            var multipart = new Multipart("mixed");
            
            // Add the plain text part
            var textPart = new TextPart("plain")
            {
                Text = "This is a test email sent using IMAP with OAuth 2.0 in .NET"
            };
            multipart.Add(textPart);
            
            // Add attachments if provided
            if (attachmentPaths != null && attachmentPaths.Count > 0)
            {
                foreach (var attachmentPath in attachmentPaths)
                {
                    if (File.Exists(attachmentPath))
                    {
                        // Create the attachment
                        var attachment = new MimePart()
                        {
                            Content = new MimeContent(File.OpenRead(attachmentPath)),
                            ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                            ContentTransferEncoding = ContentEncoding.Base64,
                            FileName = Path.GetFileName(attachmentPath)
                        };
                        
                        // Determine content type based on file extension
                        attachment.ContentType.MimeType = GetMimeType(Path.GetExtension(attachmentPath));
                        
                        // Add to multipart
                        multipart.Add(attachment);
                    }
                }
            }
            
            // Set the message body to our multipart
            message.Body = multipart;

            using (var client = new ImapClient())
            {
                // Connect to Gmail's IMAP server
                await client.ConnectAsync("imap.gmail.com", 993, SecureSocketOptions.SslOnConnect);

                // Authenticate using OAuth 2.0
                var oauth2 = new SaslMechanismOAuth2("your.email@gmail.com", oauth2Token);
                await client.AuthenticateAsync(oauth2);

                // Get the Sent folder
                var sent = client.GetFolder(SpecialFolder.Sent);
                await sent.OpenAsync(FolderAccess.ReadWrite);

                // Append the message to the Sent folder
                await sent.AppendAsync(message);

                // Disconnect
                await client.DisconnectAsync(true);
            }
        }
    }
}