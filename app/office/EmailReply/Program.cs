namespace LawhubOffice.EmailReply;

using System;
using System.IO;
using Newtonsoft.Json.Linq;

class Program
{
    static void Main()
    {
        try
        {
            // === Load and parse JSON ===
            string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..");
            string jsonPath = Path.Combine(baseDir, "app" , "data.json");
            string jsonText = File.ReadAllText(jsonPath);
            JObject json = JObject.Parse(jsonText);

            string clientEmail = (string)json["form"]["clientEmail"];
            string clientLanguage = (string)json["form"]["clientLanguage"];
            string lawyerName = (string)json["lawyer"]["name"];

            // === Load the appropriate HTML template ===
            string templateFile = clientLanguage == "English" ? "en/Reply.html" : "fr/Reply.html";
            string templatePath = Path.Combine(baseDir, "templates", templateFile);
            string htmlBody = File.ReadAllText(templatePath);

            // === Replace placeholder ===
            htmlBody = htmlBody.Replace("{{lawyerName}}", lawyerName);

            // === Start Outlook and create mail ===
            Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
            dynamic outlookApp = Activator.CreateInstance(outlookType);
            const int olMailItem = 0;
            dynamic mail = outlookApp.CreateItem(olMailItem);

            mail.To = clientEmail;
            mail.Subject = clientLanguage == "English" ? "Reply - Allen Madelin" : "Réponse - Allen Madelin";
            mail.HTMLBody = htmlBody;
            mail.Display();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
