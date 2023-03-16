using Microsoft.Graph.Models;
using MSGraphAppWithAppPermissions;

Console.WriteLine("Welcome to FileCaster!\n");

var settings = Settings.LoadSettings();

// Initialize Graph
InitializeGraph(settings);

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("****Please choose 1 to begin");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Send file from user drive to emails");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch (choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            await SendFilefromDriveToEmailsAndChannels();
            break;
        default:
            Console.WriteLine("\n****Invalid choice! Please try again.");
            break;
    }
}

void InitializeGraph(Settings settings)
{
    GraphHelper.InitializeGraphForAppOnlyAuth(settings);
}

async Task SendFilefromDriveToEmailsAndChannels()
{
    try
    {
        //User section
        int choice = 0;

        var users = await GraphHelper.GetUsersAsync();

        if (users != null)
        {

            Console.WriteLine("\n****Select user to view their drive files : ");

            int i = 0;

            foreach (var user in users.Value)
            {
                Console.WriteLine($"    {++i}   User: {user.DisplayName ?? "NO NAME"} {user.Mail ?? "NO EMAIL"} ");
            }

            choice = int.Parse(Console.ReadLine() ?? string.Empty);

            while (choice < 1 || choice > users.Value.Count)
            {
                choice = int.Parse(Console.ReadLine() ?? string.Empty);
                Console.WriteLine("\n****Invalid choice! Please try again.");
            }

            var selectedUser = users.Value[choice - 1];
            var selectedUserId = selectedUser.Id;


            //User drive and files section
            var (filesOfUser, userDrive) = await GraphHelper.GetFilesOfUser(selectedUserId);

            Console.WriteLine("\n****Select file to send : ");

            i = 0;
            choice = 0;
            foreach (var file in filesOfUser.Value)
            {
                Console.WriteLine($"    {++i}   User: {file.Name ?? "NO NAME"} ");
            }

            choice = int.Parse(Console.ReadLine() ?? string.Empty);

            while (choice < 1 || choice > users.Value.Count)
            {
                choice = int.Parse(Console.ReadLine() ?? string.Empty);
                Console.WriteLine("\n****Invalid choice! Please try again.");
            }

            var selectedFile = filesOfUser.Value[choice - 1];
            var selectedFileLink = filesOfUser.Value[choice - 1].WebUrl;

            if(selectedFile != null && selectedFileLink != null)
            {
                await GraphHelper.SendFilefromDriveToEmails(selectedUser, users, selectedFile, userDrive);

                Console.WriteLine("\n****File sent as attachment to all users****\n\n");
            }
            else
            {
                Console.WriteLine("\n****The File was not sent any user****\n\n");
            }

            
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error while excecuting. Error : {ex}");
    }
}