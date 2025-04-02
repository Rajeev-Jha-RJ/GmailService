using MailKit;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using System;
using System.IO;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Requests;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Util.Store;
using System.Collections.Generic;

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

namespace SmtpOAuth2EmailSender
{
    class Program
    {
        // Gmail API scopes for SMTP access
        static string[] Scopes = { "https://mail.google.com/" };
        static string ApplicationName = "Gmail SMTP OAuth2 Email Sender";

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
                
                // Define the email addresses
                string primaryEmail = "dhruv@domain.com"; // The authenticated user's email
                string sendAsEmail = "james@domain.com";  // Email to send as/on behalf of
                string to = "recipient@example.com";

                // Choose sending method:
                // 1. Send as the primary authenticated user
                // await SendEmailWithOAuth2Async(oauth2Token, primaryEmail, null, to, attachments);
                
                // 2. Send on behalf of another user
                await SendEmailWithOAuth2Async(oauth2Token, primaryEmail, sendAsEmail, to, attachments);

                Console.WriteLine("Email with attachments sent successfully!");
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

            // Get the access token
            var token = await credential.GetAccessTokenForRequestAsync();
            return token;
        }

        private static async Task SendEmailWithOAuth2Async(string oauth2Token, string primaryEmail, string sendAsEmail = null, 
            string to = "recipient@example.com", List<string> attachmentPaths = null)
        {
            // Create a new message
            var message = new MimeMessage();
            
            // Set up the From header based on whether we're sending as or on behalf of
            if (string.IsNullOrEmpty(sendAsEmail))
            {
                // Just send as the primary email
                message.From.Add(new MailboxAddress("Primary User", primaryEmail));
            }
            else
            {
                // Send on behalf of sendAsEmail
                message.From.Add(new MailboxAddress("Send-As User", sendAsEmail));
                
                // Add the primary email as the Sender
                message.Sender = new MailboxAddress("Primary User", primaryEmail);
            }
            
            message.To.Add(new MailboxAddress("Recipient Name", to));
            message.Subject = "Test email with Send-As functionality";

            // Create a multipart message body
            var multipart = new Multipart("mixed");
            
            // Add the plain text part
            var textPart = new TextPart("plain")
            {
                Text = "This is a test email that demonstrates send-as/send-on-behalf-of functionality in .NET"
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
                var oauth2 = new SaslMechanismOAuth2(primaryEmail, oauth2Token);
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
    }
}