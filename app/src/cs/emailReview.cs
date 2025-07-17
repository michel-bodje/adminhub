using System;
using System.IO;
using Newtonsoft.Json.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

class Program
{
    static void Main(string[] args)
    {
        try
        {
            // === Load JSON ===
            JObject json = Util.ReadJsonFromStdin();

            string clientEmail = (string)json["form"]["clientEmail"];
            string clientLanguage = (string)json["form"]["clientLanguage"];

            // === Load the appropriate HTML template ===
            string lang = clientLanguage == "Français" ? "fr" : "en";

            string templatePath = Path.Combine(Util.TemplateDir, lang, "Review.html");

            string htmlBody = File.ReadAllText(templatePath);

            // === Start Outlook and create mail using Interop ===
            var outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            var mail = (Microsoft.Office.Interop.Outlook.MailItem)outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

            mail.To = clientEmail;
            mail.Subject = clientLanguage == "English"
                ? "Google Review - Allen Madelin"
                : "Commentaires Google - Allen Madelin";
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