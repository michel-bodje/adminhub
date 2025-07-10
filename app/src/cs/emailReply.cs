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
            string jsonText = File.ReadAllText(Util.JsonPath);
            JObject json = JObject.Parse(jsonText);

            string clientEmail = (string)json["form"]["clientEmail"];
            string clientLanguage = (string)json["form"]["clientLanguage"];
            string lawyerName = (string)json["lawyer"]["name"];
            string lawyerId = (string)json["lawyer"]["id"];

            // === Load the appropriate HTML template ===
            string templateFile = clientLanguage == "English" ? "en/Reply.html" : "fr/Reply.html";
            string templatePath = Path.Combine(Util.RootDir, "app", "templates", templateFile);
            string htmlBody = File.ReadAllText(templatePath);

            // === Replace placeholder ===
            string lawyerString = Util.GetLawyerString(lawyerName, lawyerId);

            htmlBody = htmlBody.Replace("{{lawyerName}}", lawyerString);

            // === Start Outlook and create mail using Interop ===
            var outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            var mail = (Microsoft.Office.Interop.Outlook.MailItem)outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

            mail.To = clientEmail;
            mail.Subject = clientLanguage == "English" ? "Reply - Allen Madelin" : "Réponse - Allen Madelin";
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
