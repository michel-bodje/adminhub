using System;
using System.IO;
using Newtonsoft.Json.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

class Program
{
    static void Main(string[] args)
    {
        string BASE_PATH;

        if (args.Length > 0 && Directory.Exists(args[0]))
            BASE_PATH = args[0];
        else
            BASE_PATH = AppDomain.CurrentDomain.BaseDirectory;

        try
        {
            // === Load and parse JSON ===
            string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..");
            string jsonPath = Path.Combine(baseDir, "app", "data.json");
            string jsonText = File.ReadAllText(jsonPath);
            JObject json = JObject.Parse(jsonText);

            string clientEmail = (string)json["form"]["clientEmail"];
            string clientLanguage = (string)json["form"]["clientLanguage"];

            // === Load the appropriate HTML template ===
            string templateFile = clientLanguage == "English" ? "en/Review.html" : "fr/Review.html";
            string templatePath = Path.Combine(baseDir, "app", "templates", templateFile);
            string htmlBody = File.ReadAllText(templatePath);

            // === Start Outlook and create mail using Interop ===
            var outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            var mail = (Microsoft.Office.Interop.Outlook.MailItem)outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

            mail.To = clientEmail;
            mail.Subject = clientLanguage == "English" ? "Google Review - Allen Madelin" : "Commentaires Google - Allen Madelin";
            mail.HTMLBody = htmlBody;
            mail.Display();
            WindowFocus.ShowWithFocus(mail.GetInspector);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}