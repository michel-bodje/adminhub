using System;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;

class Program
{
    static void Main()
    {
        try
        {
            // === Load JSON form state ===
            string jsonPath = @"\\AMNAS\amlex\Admin\Scripts\lawhub\app\data.json";
            string jsonText = File.ReadAllText(jsonPath);
            JObject json = JObject.Parse(jsonText);

            string clientEmail = json["form"]["clientEmail"]?.ToString();
            string clientLanguage = json["form"]["clientLanguage"]?.ToString() ?? "English";
            string location = json["form"]["location"]?.ToString() ?? "";
            bool isFirstConsult = json["form"]["isFirstConsultation"]?.ToObject<bool>() ?? true;
            string appointmentDate = json["form"]["appointmentDate"]?.ToString();
            string appointmentTime = json["form"]["appointmentTime"]?.ToString();
            string lawyerName = json["lawyer"]?["name"]?.ToString() ?? "";

            // === Validate required fields ===
            if (string.IsNullOrEmpty(clientEmail) || !clientEmail.Contains("@"))
                throw new Exception("Missing or invalid client email.");
            if (string.IsNullOrEmpty(location))
                throw new Exception("Missing location (used for template).");

            string lang = (clientLanguage == "Français") ? "fr" : "en";
            string templateName = char.ToUpper(location[0]) + location.Substring(1).ToLower() + ".html";
            string templatePath = @"\\AMNAS\amlex\Admin\Scripts\lawhub\templates\{lang}\{templateName}";

            if (!File.Exists(templatePath))
                throw new FileNotFoundException("Template not found.", templatePath);

            string htmlBody = File.ReadAllText(templatePath);

            // === Insert confirmation-specific tokens ===
            if (!string.IsNullOrEmpty(appointmentDate) && !string.IsNullOrEmpty(appointmentTime))
            {
                DateTime slot = DateTime.ParseExact(
                    appointmentDate + " " + appointmentTime,
                    "yyyy-MM-dd HH:mm",
                    CultureInfo.InvariantCulture
                );

                string locale = (lang == "fr") ? "fr-CA" : "en-US";
                string formattedDate = slot.ToString("D", CultureInfo.CreateSpecificCulture(locale));
                string formattedTime = slot.ToString("t", CultureInfo.CreateSpecificCulture(locale));

                int baseRate = isFirstConsult ? 125 : 350;
                double totalRate = AddTaxes(baseRate);

                htmlBody = htmlBody
                    .Replace("{{date}}", formattedDate)
                    .Replace("{{time}}", formattedTime)
                    .Replace("{{rates}}", baseRate.ToString())
                    .Replace("{{totalRates}}", totalRate.ToString("0.00"));
            }

            htmlBody = htmlBody.Replace("{{lawyerName}}", lawyerName);

            // === Generate Outlook draft email ===
            dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application"));
            dynamic mail = outlook.CreateItem(0); // 0 = olMailItem

            mail.To = clientEmail;
            mail.Subject = (lang == "fr")
                ? "Confirmation de rendez-vous - Allen Madelin"
                : "Appointment Confirmation - Allen Madelin";
            mail.HTMLBody = htmlBody;
            mail.Display(true);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    static double AddTaxes(double baseAmount)
    {
        return baseAmount * 1.14975; // GST + QST
    }
}
