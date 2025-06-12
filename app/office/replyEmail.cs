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
            string jsonPath = @"\\AMNAS\amlex\Admin\Scripts\lawhub\app\data.json";
            string jsonText = File.ReadAllText(jsonPath);
            JObject json = JObject.Parse(jsonText);

            string clientEmail = json["form"]["clientEmail"].ToString();
            string clientLanguage = json["form"]["clientLanguage"].ToString();
            string lawyerName = json["lawyer"]["name"].ToString();

            // === Load the appropriate HTML template ===
            string templateFile = clientLanguage == "English" ? "en/Reply.html" : "fr/Reply.html";
            string templatePath = @"\\AMNAS\amlex\Admin\Scripts\lawhub\templates\" + templateFile;
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
