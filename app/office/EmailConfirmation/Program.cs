namespace LawhubOffice.EmailConfirmation;

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
            string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..");
            string jsonPath = Path.Combine(baseDir, "app", "data.json");
            string jsonText = File.ReadAllText(jsonPath);

            JObject json = JObject.Parse(jsonText);

            string clientEmail = (string)json["form"]["clientEmail"];
            string clientLanguage = (string)json["form"]["clientLanguage"];
            string location = (string)json["form"]["location"];
            bool isFirstConsult = (bool)json["form"]["isFirstConsultation"];
            string appointmentDate = (string)json["form"]["appointmentDate"];
            string appointmentTime = (string)json["form"]["appointmentTime"];
            string lawyerName = (string)json["lawyer"]["name"];

            // === Validate required fields ===
            if (string.IsNullOrEmpty(clientEmail) || !clientEmail.Contains("@"))
                throw new Exception("Missing or invalid client email.");
            if (string.IsNullOrEmpty(location))
                throw new Exception("Missing location (used for template).");

            string lang = (clientLanguage == "Français") ? "fr" : "en";
            string templateName = char.ToUpper(location[0]) + location.Substring(1).ToLower() + ".html";
            string templatePath = Path.Combine(baseDir, "templates", lang, templateName);

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
                // Format date with ordinal suffix for English
                string formattedDate;
                if (lang == "en")
                {
                    int day = slot.Day;
                    string suffix = (day % 10 == 1 && day != 11) ? "st"
                        : (day % 10 == 2 && day != 12) ? "nd"
                        : (day % 10 == 3 && day != 13) ? "rd"
                        : "th";
                    formattedDate = string.Format("{0:dddd, MMMM} {1}{2}, {0:yyyy}", slot, day, suffix);
                }
                else
                {
                    formattedDate = slot.ToString("D", CultureInfo.CreateSpecificCulture(locale));
                }
                // Format time as 2:00 PM in English, 14:00 in French
                string formattedTime = slot.ToString(
                    (lang == "en") ? "h:mm tt" : "HH:mm",
                    CultureInfo.CreateSpecificCulture(locale)
                );

                int baseRate = isFirstConsult ? 125 : 350;
                double totalRate = Util.AddTaxes(baseRate);

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
                : "Confirmation of appointment - Allen Madelin";
            mail.HTMLBody = htmlBody;
            mail.Display(true);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
