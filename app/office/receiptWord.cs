using System;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;

class Program
{
    [STAThread] // Required for SaveFileDialog
    static void Main()
    {
        try
        {
            // === Load JSON form state ===
            string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..");
            string jsonPath = Path.Combine(baseDir, "app", "data.json");
            string jsonText = File.ReadAllText(jsonPath);
            JObject json = JObject.Parse(jsonText);

            // === Extract form values ===
            string clientName = (string)json["form"]["clientName"];
            string clientLanguage = (string)json["form"]["clientLanguage"];
            string paymentMethod = (string)json["form"]["paymentMethod"];
            double depositAmount = (double)json["form"]["depositAmount"];
            string lawyerName = (string)json["lawyer"]["name"];

            if (string.IsNullOrEmpty(clientName) || string.IsNullOrEmpty(lawyerName))
                throw new Exception("Missing required input.");

            string lang = clientLanguage == "Français" ? "fr" : "en";
            string locale = lang == "fr" ? "fr-CA" : "en-US";
            string templatePath = Path.Combine(baseDir, "templates", lang, "Receipt.docx");

            if (!File.Exists(templatePath))
                throw new FileNotFoundException("Receipt template not found.", templatePath);

            // === Create temp copy of template ===
            string tempDocPath = Path.Combine(Path.GetTempPath(), "Receipt_" + Guid.NewGuid() + ".docx");
            File.Copy(templatePath, tempDocPath, overwrite: true);

            // === Prepare replacement values ===
            string formattedDate = DateTime.Today.ToString("D", CultureInfo.CreateSpecificCulture(locale));
            string formattedAmount = depositAmount.ToString("F2");

            // === Open and fill Word document ===
            Type wordType = Type.GetTypeFromProgID("Word.Application");
            dynamic wordApp = Activator.CreateInstance(wordType);
            dynamic doc = wordApp.Documents.Open(tempDocPath);

            Util.ReplaceWordText(doc, "{user}", "Michel Assi-Bodje");
            Util.ReplaceWordText(doc, "{reason}", "{}");
            Util.ReplaceWordText(doc, "{clientName}", clientName);
            Util.ReplaceWordText(doc, "{paymentMethod}", paymentMethod);
            Util.ReplaceWordText(doc, "{depositAmount}", formattedAmount);
            Util.ReplaceWordText(doc, "{lawyerName}", lawyerName);
            Util.ReplaceWordText(doc, "{date}", formattedDate);

            wordApp.Visible = true;

            // Show the print dialog for the user to choose settings and print
            wordApp.Dialogs[88].Show(); // 88 = File > Print dialog


            // === Ask user to export PDF ===
            string pdfPath = null;
            var saveDialog = new SaveFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                Title = "Export Receipt as PDF",
                FileName = DateTime.Now.ToString("yyyy-MM-dd") + "_" + clientName.Replace(" ", "-") + "_" + ".pdf"
            };

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                pdfPath = saveDialog.FileName;
                doc.ExportAsFixedFormat(pdfPath, 17); // 17 = wdExportFormatPDF
            }

            // === Cleanup ===
            doc.Close(false);
            wordApp.Quit();
            File.Delete(tempDocPath);

            Console.WriteLine("Receipt saved to: " + pdfPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
