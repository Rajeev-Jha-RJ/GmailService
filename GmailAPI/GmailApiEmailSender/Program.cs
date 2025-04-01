using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Text;

namespace GmailApiEmailSender
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        static string[] Scopes = { GmailService.Scope.GmailSend };
        static string ApplicationName = "Gmail API .NET Email Sender";

        static async Task Main(string[] args)
        {
            try
            {
                // Create Gmail API service
                var service = await CreateGmailServiceAsync();

                // Define the email
                string to = "recipient@example.com";
                string subject = "Test email from Gmail API";
                string body = "This is a test email sent using the Gmail API in .NET";

                // Example attachments list - replace with your actual files
                List<string> attachments = new List<string>
                {
                    "file1.pdf",
                    "image.jpg",
                    "document.docx"
                };
                
                // Send the email with attachments
                await SendEmailAsync(service, to, subject, body, attachments);

                Console.WriteLine("Email with attachments sent successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static async Task<GmailService> CreateGmailServiceAsync()
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

            // Create Gmail API service
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        private static async Task SendEmailAsync(GmailService service, string to, string subject, string body, List<string> attachmentPaths = null)
        {
            // Create a new message
            var email = new Message();
            email.Raw = Base64UrlEncode(CreateMessage(to, subject, body, attachmentPaths));

            // Send the message
            await service.Users.Messages.Send(email, "me").ExecuteAsync();
        }

        private static string CreateMessage(string to, string subject, string body, List<string> attachmentPaths = null)
        {
            // Generate a boundary for the multipart message
            string boundary = $"==Boundary_${Guid.NewGuid().ToString("N")}";
            
            // Create the message headers
            var headers = new StringBuilder();
            headers.AppendLine($"To: {to}");
            headers.AppendLine($"Subject: {subject}");
            headers.AppendLine("MIME-Version: 1.0");
            headers.AppendLine($"Content-Type: multipart/mixed; boundary=\"{boundary}\"");
            headers.AppendLine();
            
            // Start the multipart message
            var messageBody = new StringBuilder();
            messageBody.AppendLine($"--{boundary}");
            messageBody.AppendLine("Content-Type: text/plain; charset=utf-8");
            messageBody.AppendLine("Content-Transfer-Encoding: 7bit");
            messageBody.AppendLine();
            messageBody.AppendLine(body);
            
            // Add attachments if provided
            if (attachmentPaths != null && attachmentPaths.Count > 0)
            {
                foreach (var attachmentPath in attachmentPaths)
                {
                    if (File.Exists(attachmentPath))
                    {
                        // Read the file content
                        byte[] fileBytes = File.ReadAllBytes(attachmentPath);
                        string fileBase64 = Convert.ToBase64String(fileBytes);
                        
                        // Get the file name
                        string fileName = Path.GetFileName(attachmentPath);
                        
                        // Determine content type based on file extension (simplified)
                        string contentType = GetMimeType(Path.GetExtension(attachmentPath));
                        
                        // Add the attachment part
                        messageBody.AppendLine($"--{boundary}");
                        messageBody.AppendLine($"Content-Type: {contentType}; name=\"{fileName}\"");
                        messageBody.AppendLine("Content-Transfer-Encoding: base64");
                        messageBody.AppendLine($"Content-Disposition: attachment; filename=\"{fileName}\"");
                        messageBody.AppendLine();
                        
                        // Split the base64 string into 76-character lines
                        for (int i = 0; i < fileBase64.Length; i += 76)
                        {
                            int length = Math.Min(76, fileBase64.Length - i);
                            messageBody.AppendLine(fileBase64.Substring(i, length));
                        }
                    }
                }
            }
            
            // End the multipart message
            messageBody.AppendLine($"--{boundary}--");
            
            return headers.ToString() + messageBody.ToString();
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

        private static string Base64UrlEncode(string input)
        {
            var inputBytes = Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(inputBytes)
                .Replace('+', '-')
                .Replace('/', '_')
                .Replace("=", "");
        }
    }
}