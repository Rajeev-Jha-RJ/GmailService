using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Requests;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GmailApiEmailSender
{
    class Program
    {
        // Gmail API scopes
        static string[] Scopes = { GmailService.Scope.GmailSend };
        static string ApplicationName = "Gmail API .NET Email Sender";

        static async Task Main(string[] args)
        {
            try
            {
                // Create Gmail API service
                var service = await CreateGmailServiceAsync();

                // Define the email
                string primaryEmail = "dhruv@domain.com"; // The authenticated user's email
                string sendAsEmail = "james@domain.com";  // Email to send as/on behalf of
                string to = "rajeev@isgesolutions.com";
                string subject = "Test email from Gmail API";
                string body = "This is a test email sent using the Gmail API in .NET";

                // Example attachments list - replace with your actual files
                List<string> attachments = new List<string>
                {
                    "file1.pdf",
                    "image.jpg",
                    "document.docx"
                };
                
                // Choose sending method:
                // 1. Send as the primary authenticated user
                await SendEmailAsync(service, to, subject, body, attachments);
                
                // 2. Send on behalf of another user (shows "dhruv@domain.com on behalf of james@domain.com")
                // await SendEmailOnBehalfOfAsync(service, primaryEmail, sendAsEmail, to, subject, body, attachments);
                
                // 3. Send as another user completely (shows only "james@domain.com" as sender)
                // await SendEmailAsAsync(service, sendAsEmail, to, subject, body, attachments);

                Console.WriteLine("Email with attachments sent successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                }
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
                
                // Create a custom flow
                var flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
                {
                    ClientSecrets = GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes = Scopes,
                    DataStore = new FileDataStore(credPath, true)
                });
                
                // Create authorization code request URL
                var codeReceiver = new LoopbackCodeReceiver();
                
                // Authorize using that flow
                credential = new AuthorizationCodeInstalledApp(flow, codeReceiver).AuthorizeAsync(
                    "user", CancellationToken.None).Result;

                Console.WriteLine($"Credential file saved to: {credPath}");
                Console.WriteLine("Using fixed port 8080 for OAuth callback");
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
        
        private static async Task SendEmailOnBehalfOfAsync(GmailService service, string primaryEmail, string sendAsEmail, 
            string to, string subject, string body, List<string> attachmentPaths = null)
        {
            // Create a new message
            var email = new Message();
            email.Raw = Base64UrlEncode(CreateMessageOnBehalfOf(primaryEmail, sendAsEmail, to, subject, body, attachmentPaths));

            // Send the message
            await service.Users.Messages.Send(email, "me").ExecuteAsync();
        }
        
        private static async Task SendEmailAsAsync(GmailService service, string sendAsEmail, 
            string to, string subject, string body, List<string> attachmentPaths = null)
        {
            // Create a new message
            var email = new Message();
            email.Raw = Base64UrlEncode(CreateMessageAs(sendAsEmail, to, subject, body, attachmentPaths));

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
        
        private static string CreateMessageOnBehalfOf(string primaryEmail, string sendAsEmail, string to, 
            string subject, string body, List<string> attachmentPaths = null)
        {
            // Generate a boundary for the multipart message
            string boundary = $"==Boundary_${Guid.NewGuid().ToString("N")}";
            
            // Create the message headers
            var headers = new StringBuilder();
            
            // Set the From header to indicate "on behalf of"
            headers.AppendLine($"From: {primaryEmail}");
            headers.AppendLine($"Sender: {sendAsEmail}");
            
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
        
        private static string CreateMessageAs(string sendAsEmail, string to, 
            string subject, string body, List<string> attachmentPaths = null)
        {
            // Generate a boundary for the multipart message
            string boundary = $"==Boundary_${Guid.NewGuid().ToString("N")}";
            
            // Create the message headers - this fully impersonates the sendAsEmail address
            var headers = new StringBuilder();
            headers.AppendLine($"From: {sendAsEmail}");
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

    // Custom code receiver with fixed port
    public class LoopbackCodeReceiver : ICodeReceiver
    {
        private readonly string _redirectUri;
        private readonly HttpListener _listener;

        public LoopbackCodeReceiver()
        {
            _redirectUri = "http://localhost:8080/";
            _listener = new HttpListener();
            _listener.Prefixes.Add(_redirectUri);
        }

        public string RedirectUri => _redirectUri;

        public async Task<AuthorizationCodeResponseUrl> ReceiveCodeAsync(AuthorizationCodeRequestUrl url, 
            CancellationToken taskCancellationToken)
        {
            var authorizationUrl = url.Build().ToString();
            
            // Open the browser
            try
            {
                OpenBrowser(authorizationUrl);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to open browser: {ex.Message}");
                Console.WriteLine($"Please manually open the following URL: {authorizationUrl}");
            }

            // Start the listener
            _listener.Start();
            
            try
            {
                Console.WriteLine($"Waiting for authorization response on {_redirectUri}...");
                
                // Wait for the callback
                var context = await _listener.GetContextAsync();
                var queryString = context.Request.Url.Query;
                
                // Send a response to the browser
                using (var response = context.Response)
                {
                    string responseHtml = "<html><head><title>Authentication Complete</title></head>" +
                                         "<body>Authentication complete. You can close this window now.</body></html>";
                    byte[] buffer = System.Text.Encoding.UTF8.GetBytes(responseHtml);
                    response.ContentLength64 = buffer.Length;
                    response.StatusCode = 200;
                    response.ContentType = "text/html";
                    await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                }
                
                // Extract the authorization code
                var authorizationResponse = new AuthorizationCodeResponseUrl();
                if (!string.IsNullOrEmpty(queryString))
                {
                    if (queryString.StartsWith("?"))
                        queryString = queryString.Substring(1);
                    
                    foreach (var pair in queryString.Split('&'))
                    {
                        var parts = pair.Split('=');
                        if (parts.Length == 2)
                        {
                            var key = parts[0];
                            var value = Uri.UnescapeDataString(parts[1]);
                            
                            if (key == "code")
                                authorizationResponse.Code = value;
                            else if (key == "error")
                                authorizationResponse.Error = value;
                        }
                    }
                }
                
                return authorizationResponse;
            }
            finally
            {
                _listener.Stop();
            }
        }
        
        private void OpenBrowser(string url)
        {
            // Cross-platform browser opening
            if (OperatingSystem.IsWindows())
            {
                url = url.Replace("&", "^&");
                Process.Start(new ProcessStartInfo("cmd", $"/c start {url}") { CreateNoWindow = true });
            }
            else if (OperatingSystem.IsMacOS())
            {
                Process.Start("open", url);
            }
            else if (OperatingSystem.IsLinux())
            {
                Process.Start("xdg-open", url);
            }
            else
            {
                throw new PlatformNotSupportedException("Automatic browser opening not supported on this platform.");
            }
        }
    }
}