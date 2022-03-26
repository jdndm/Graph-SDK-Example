using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;

namespace GraphTutorial
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(".NET Core Graph Tutorial\n");

            var appConfig = LoadAppSettings();

            if (appConfig == null)
            {
                Console.WriteLine("Missing or invalid appsettings.json...exiting");
                return;
            }

            var appId = appConfig["appId"];
            var scopesString = appConfig["scopes"];
            var scopes = scopesString.Split(';');

            // Initialize Graph client
            GraphHelper.Initialize(appId, scopes, (code, cancellation) => {
                Console.WriteLine(code.Message);
                return Task.FromResult(0);
            });

            var accessToken = GraphHelper.GetAccessTokenAsync(scopes).Result;

            int choice = -1;

            while (choice != 0)
            {
                Console.WriteLine("Please choose one of the following options:");
                Console.WriteLine("0. Exit");
                Console.WriteLine("1. Display access token");
                Console.WriteLine("2. Fetch latests Messages");
                Console.WriteLine("3. Add an event");

                try
                {
                    choice = int.Parse(Console.ReadLine());
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
                        // Display access token
                        Console.WriteLine($"Access token: {accessToken}\n");
                        break;
                    case 2:
                        var messages = GraphHelper.GetMessagesAsync().Result;
                        //Console.WriteLine($"Welcome {user.DisplayName}!\n");
                        int i = 0;
                        int j = 0;
                        while (i < 1)
                        {
                            if (messages[j].Subject == "UniSuper - Forgotten Username")
                            {
                                Console.WriteLine(messages[j].Body.Content);
                                var expectedString = "Your username is: ";
                                if (messages[j].Body.Content.Contains(expectedString))
                                {
                                    Console.WriteLine("Forgot Useranme email is recieved");
                                }
                                else
                                {
                                    Console.WriteLine("Email not recieved.");
                                }
                                i++;
                            }
                            j++;
                        }

                        
                        
                        //Console.WriteLine(messages[0].Body.Content);
                        //Console.WriteLine(messages);
                        break;
                    case 3:
                        // Create a new event
                        break;
                    default:
                        Console.WriteLine("Invalid choice! Please try again.");
                        break;
                }
            }
        }
        static IConfigurationRoot LoadAppSettings()
        {
            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();

            // Check for required settings
            if (string.IsNullOrEmpty(appConfig["appId"]) ||
                string.IsNullOrEmpty(appConfig["scopes"]))
            {
                return null;
            }

            return appConfig;
        }
    }

}