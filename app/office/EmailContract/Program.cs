namespace LawhubOffice.EmailContract;

using System;
using System.IO;
using System.Globalization;
using Newtonsoft.Json.Linq;

class Program
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

            string clientEmail = json["form"]["clientEmail"].ToString();
            string clientLanguage = json["form"]["clientLanguage"].ToString();
            double depositAmount = json["form"]["depositAmount"].ToObject<double>();
            string lawyerName = json["lawyer"]["name"].ToString();

            string lang = clientLanguage == "Français" ? "fr" : "en";
            string templatePath = Path.Combine(templateFolder, lang, "Contract.html");

            string htmlBody = File.ReadAllText(templatePath);

            // === Replace placeholders ===
            double totalAmount = Util.AddTaxes(depositAmount, true);

            htmlBody = htmlBody
                .Replace("{{depositAmount}}", depositAmount.ToString("F0"))
                .Replace("{{totalAmount}}", totalAmount.ToString("F2"))
                .Replace("{{lawyerName}}", lawyerName);

            string subject = lang == "fr" ? "Contrat de services - Allen Madelin" : "Contract of services - Allen Madelin";

            // === Create email using late binding (COM) ===
            Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
            dynamic outlook = Activator.CreateInstance(outlookType);
            dynamic mail = outlook.CreateItem(0); // 0 = olMailItem

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
