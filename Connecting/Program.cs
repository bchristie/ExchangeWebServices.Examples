using System;
using System.Configuration;
using Microsoft.Exchange.WebServices.Data;

namespace Connecting
{
    class Program
    {
        static void Main(string[] args)
        {
            // Make sure to configure the App.config in the solution directory
            String username = ConfigurationManager.AppSettings["username"];
            String password = ConfigurationManager.AppSettings["password"];
            String ewsUrl = ConfigurationManager.AppSettings["ews"];

            // Connect to the service
            ExchangeService service = new ExchangeService();
            service.Credentials = new WebCredentials(username, password);
            // Either supply the ews endpoint, or discover it
            if (!String.IsNullOrEmpty(ewsUrl))
            {
                service.Url = new Uri(ewsUrl);
            }
            else
            {
                service.AutodiscoverUrl(username, url =>
                {
                    // here you validate the URL to make sure
                    // it's acceptable but we're just going to
                    // assume it's ok.
                    return true;
                });
            }
            Console.WriteLine("Connected!");

            try
            {
                // For show, let's just grab the inbox and the top 5 messages
                ItemView itemView = new ItemView(5);
                FindItemsResults<Item> items = service.FindItems(WellKnownFolderName.Inbox, itemView);
                foreach (var item in items)
                {
                    Console.WriteLine(item.Subject);
                }
            }
            catch (ServiceRequestException sex)
            {
                Console.WriteLine("Server failure.");
                Console.WriteLine(sex.Message);
                Console.WriteLine("(did you fill out the appSettings.config?)");
            }
            catch (Exception ex)
            {
                Console.WriteLine("General failure.");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
