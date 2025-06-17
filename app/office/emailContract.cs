using System;
using System.IO;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

class ContractEmail
{
    static void Main()
    {
        try
        {
            string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..");
            string jsonPath = Path.Combine(baseDir, "app", "data.json");
            string templateFolder = Path.Combine(baseDir, "templates");

            // === Read JSON ===
            JObject json = JObject.Parse(File.ReadAllText(jsonPath));

            string clientEmail = (string)json["form"]["clientEmail"];
            string clientLanguage = (string)json["form"]["clientLanguage"];
            double depositAmount = (double)json["form"]["depositAmount"];
            string lawyerName = (string)json["lawyer"]["name"];
            string lawyerId = (string)json["lawyer"]["id"];

            string lang = clientLanguage == "Français" ? "fr" : "en";
            string templatePath = Path.Combine(templateFolder, lang, "Contract.html");

            string htmlBody = File.ReadAllText(templatePath);

            double totalAmount = Util.AddTaxes(depositAmount, true);

            string lawyerString = Util.GetLawyerString(lawyerName, lawyerId);

            // === Replace placeholders ===
            htmlBody = htmlBody
                .Replace("{{depositAmount}}", depositAmount.ToString("F0"))
                .Replace("{{totalAmount}}", totalAmount.ToString("F2"))
                .Replace("{{lawyerName}}", lawyerString);

            string subject = lang == "fr" ? "Contrat de services - Allen Madelin" : "Contract of services - Allen Madelin";

            // === Create email using Outlook Interop ===
            var outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            var mail = (Microsoft.Office.Interop.Outlook.MailItem)outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

            mail.To = clientEmail;
            mail.Subject = subject;
            mail.HTMLBody = htmlBody;

            string projectPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..");
            string pathTempFile = Path.Combine(projectPath, "app", "latest_contract_path.txt");

            if (File.Exists(pathTempFile))
            {
                string pdfPath = File.ReadAllText(pathTempFile).Trim();

                if (File.Exists(pdfPath))
                {
                    mail.Attachments.Add(pdfPath);
                    Console.WriteLine("Attached PDF contract: " + pdfPath);
                }
                else
                {
                    Console.WriteLine("PDF path recorded, but file no longer exists: " + pdfPath);
                }
            }
            else
            {
                Console.WriteLine("No PDF was generated, skipping attachment.");
            }

            mail.Display(); // Opens the draft
            WindowFocus.ShowWithFocus(mail.GetInspector); // Force focus
            
            // === Clean up ===
            if (File.Exists(pathTempFile))
                File.Delete(pathTempFile);

            Console.WriteLine("Draft contract email created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
