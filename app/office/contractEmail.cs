using System;
using System.IO;
using System.Globalization;
using Newtonsoft.Json.Linq;

class ContractEmail
{
    static void Main()
    {
        string jsonPath = @"\\AMNAS\amlex\Admin\Scripts\lawhub\app\data.json";
        string templateFolder = @"\\AMNAS\amlex\Admin\Scripts\lawhub\templates\";

        // === Read JSON ===
        JObject json = JObject.Parse(File.ReadAllText(jsonPath));

        string clientEmail = (string)json["form"]["clientEmail"];
        string clientLanguage = (string)json["form"]["clientLanguage"];
        double depositAmount = (double)json["form"]["depositAmount"];
        string lawyerName = (string)json["lawyer"]["name"];

        string lang = clientLanguage == "Français" ? "fr" : "en";
        string templatePath = Path.Combine(templateFolder, lang, "Contract.html");

        string htmlBody = File.ReadAllText(templatePath);

        // === Replace placeholders ===
        double totalAmount = AddTaxes(depositAmount, true);

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
        mail.Display(); // Opens the draft
        
        Console.WriteLine("Draft contract email created successfully.");
    }

    static double AddTaxes(double amount, bool addFOF = false) {
        if (double.IsNaN(amount))
            throw new ArgumentException("Amount is not a valid number.");

        double total = amount * (1 + 0.05 + 0.09975);
        if (addFOF) total += 100;

        return total;
    }
}
