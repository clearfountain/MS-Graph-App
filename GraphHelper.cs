using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.Chats.GetAllMessages;
using Microsoft.Graph.Users.Item.SendMail;
using System.IO;
using System.Text;

namespace MSGraphAppWithAppPermissions;
public class GraphHelper
{
    // Settings object
    private static Settings? _settings;
    // App-ony auth token credential
    private static ClientSecretCredential? _clientSecretCredential;
    // Client configured with app-only authentication
    private static GraphServiceClient? _appClient;

    public static void InitializeGraphForAppOnlyAuth(Settings settings)
    {
        _settings = settings;

        // Ensure settings isn't null
        _ = settings ??
            throw new System.NullReferenceException("Settings cannot be null");

        _settings = settings;

        if (_clientSecretCredential == null)
        {
            _clientSecretCredential = new ClientSecretCredential(
                _settings.TenantId, _settings.ClientId, _settings.ClientSecret);
        }

        if (_appClient == null)
        {
            _appClient = new GraphServiceClient(_clientSecretCredential,
                // Use the default scope, which will request the scopes
                // configured on the app registration
                new[] { "https://graph.microsoft.com/.default" });
        }
    }

    public static async Task<string> GetAppOnlyTokenAsync()
    {
        // Ensure credential isn't null
        _ = _clientSecretCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        // Request token with given scopes
        var context = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
        var response = await _clientSecretCredential.GetTokenAsync(context);
        return response.Token;
    }

    public static Task<UserCollectionResponse> GetUsersAsync()
    {
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        var users = _appClient.Users.GetAsync();

        return users;

        
    }

    public async static Task<(DriveItemCollectionResponse, Drive)> GetFilesOfUser(string userId)
    {
        try
        {
            var userDrive = await _appClient.Users[userId].Drive.GetAsync();

            var userDriveId = userDrive.Id;

            var rootFolder = await _appClient.Drives[userDriveId].Root.GetAsync();

            var rootFolderId = rootFolder.Id;

            var children = await _appClient.Drives[userDriveId].Items[rootFolderId].Children.GetAsync();

            return (children, userDrive);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            return (new DriveItemCollectionResponse(), new Drive());
        }
    }

    internal static async Task SendFilefromDriveToEmails(User user, UserCollectionResponse userCollectionResponse, DriveItem driveItem, Drive drive)
    {

        // Ensure client isn't null
        _ = _appClient ?? throw new System.NullReferenceException("Graph has not been initialized for user auth");

        try
        {
            var fileStream = await _appClient.Drives[drive.Id].Items[driveItem.Id].Content.GetAsync();

            byte[] bytes;

            using (var memoryStream = new MemoryStream())
            {
                fileStream.CopyTo(memoryStream);
                bytes = memoryStream.ToArray();
            }

            string base64 = Convert.ToBase64String(bytes);
            var requestBody = new SendMailPostRequestBody
            {
                Message = new Message
                {
                    Subject = "Interesting file",
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = "Howdy team! check out the attached file.",
                    },
                    ToRecipients = userCollectionResponse.Value.Select(u => new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = u.Mail
                        }
                    }).ToList(),
                    Attachments = new List<Attachment>
                    {
                        new Attachment
                        {
                            OdataType = "#microsoft.graph.fileAttachment",
                            Name = "sharedfile.png",
                            ContentType = "text/plain",
                            AdditionalData = new Dictionary<string, object>
                            {
                                {
                                    "contentBytes" , base64
                                },
                            },
                        },
                    },
                },
            };

            await _appClient.Users[user.Id].SendMail.PostAsync(requestBody);
        }
        catch(Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
}
