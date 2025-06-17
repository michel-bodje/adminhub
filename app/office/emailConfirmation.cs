using System;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

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
            bool isRefBarreau = (bool)json["form"]["isRefBarreau"];
            bool isFirstConsult = (bool)json["form"]["isFirstConsultation"];
            string appointmentDate = (string)json["form"]["appointmentDate"];
            string appointmentTime = (string)json["form"]["appointmentTime"];
            string lawyerName = (string)json["lawyer"]["name"];
            string lawyerId = (string)json["lawyer"]["id"];

            // === Validate required fields ===
            if (string.IsNullOrEmpty(clientEmail) || !Util.IsValidEmail(clientEmail))
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

                int baseRate = isRefBarreau ? 60 : isFirstConsult ? 125 : 350;
                double totalRate = Util.AddTaxes(baseRate);

                htmlBody = htmlBody
                    .Replace("{{date}}", formattedDate)
                    .Replace("{{time}}", formattedTime)
                    .Replace("{{rates}}", baseRate.ToString())
                    .Replace("{{totalRates}}", totalRate.ToString("0.00"));
            }

            string lawyerString = Util.GetLawyerString(lawyerName, lawyerId);
            
            htmlBody = htmlBody.Replace("{{lawyerName}}", lawyerString);

            // === Generate Outlook draft email ===
            Outlook.Application outlookApp = new Outlook.Application();
            Outlook.MailItem mail = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

            mail.To = clientEmail;
            mail.Subject = (lang == "fr")
                ? "Confirmation de rendez-vous - Allen Madelin"
                : "Confirmation of appointment - Allen Madelin";
            mail.HTMLBody = htmlBody;
            mail.Display();
            WindowFocus.ShowWithFocus(mail.GetInspector); // Force focus
            Console.WriteLine("Email draft created in Outlook.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
