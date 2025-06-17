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
            string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..");
            string jsonPath = Path.Combine(baseDir, "app", "data.json");
            string jsonText = File.ReadAllText(jsonPath);
            JObject json = JObject.Parse(jsonText);

            // === Extract form values ===
            string clientName = (string)json["form"]["clientName"];
            string clientLanguage = (string)json["form"]["clientLanguage"];
            string paymentMethod = (string)json["form"]["paymentMethod"];
            double depositAmount = (double)json["form"]["depositAmount"];
            string receiptReason = (string)json["form"]["receiptReason"];
            string lawyerName = (string)json["lawyer"]["name"];

            if (string.IsNullOrEmpty(clientName) || string.IsNullOrEmpty(lawyerName))
                throw new Exception("Missing required input.");

            string lang = clientLanguage == "Français" ? "fr" : "en";
            string locale = lang == "fr" ? "fr-CA" : "en-US";
            string templatePath = Path.Combine(baseDir, "app", "templates", lang, "Receipt.docx");

            if (!File.Exists(templatePath))
                throw new FileNotFoundException("Receipt template not found.", templatePath);

            // === Create temp copy of template ===
            string tempDocPath = Path.Combine(Path.GetTempPath(), "Receipt_" + Guid.NewGuid() + ".docx");
            File.Copy(templatePath, tempDocPath, true);

            // === Prepare replacement values ===
            string formattedDate = DateTime.Today.ToString("D", CultureInfo.CreateSpecificCulture(locale));
            string formattedAmount = depositAmount.ToString("F2");

            // === Open and fill Word document ===
            Type wordType = Type.GetTypeFromProgID("Word.Application");
            dynamic wordApp = Activator.CreateInstance(wordType);
            dynamic doc = wordApp.Documents.Open(tempDocPath);

            // === Prepare reason text ===
            string reasonText;
            if (receiptReason == "consultation")
                reasonText = lang == "fr" ? "une consultation juridique" : "a legal consultation";
            else if (receiptReason == "trust")
                reasonText = lang == "fr" ? "un paiement en fidéicommis" : "a trust payment";
            else
                reasonText = receiptReason; // Use the raw value for custom/other reasons
            
            // === Replace placeholders ===
            Util.WordReplaceText(doc, "{user}", "Michel Assi-Bodje");
            Util.WordReplaceText(doc, "{reason}", reasonText);
            Util.WordReplaceText(doc, "{clientName}", clientName);
            Util.WordReplaceText(doc, "{paymentMethod}", paymentMethod);
            Util.WordReplaceText(doc, "{depositAmount}", formattedAmount);
            Util.WordReplaceText(doc, "{lawyerName}", lawyerName);
            Util.WordReplaceText(doc, "{date}", formattedDate);

            wordApp.Visible = true;
            WindowFocus.ShowWithFocus(doc.ActiveWindow);

            // Show the print dialog for the user to choose settings and print
            wordApp.Dialogs[97].Show(); // 97 = wdDialogFilePrint (Ctrl+P dialog)

            // === Ask user to export PDF ===
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "PDF files (*.pdf)|*.pdf";
                dialog.Title = "Save Receipt as PDF";
                dialog.FileName = DateTime.Now.ToString("yyyy-MM-dd") + "_" + clientName.Replace(" ", "-") + "_" + ".pdf";

                string projectPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..");

                string tempFilePath = Path.Combine(projectPath, "app", "latest_receipt_path.txt");
    
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string pdfPath = dialog.FileName;
                    doc.ExportAsFixedFormat(pdfPath, 17); // 17 = wdExportFormatPDF

                    // Save the PDF path to temp file
                    File.WriteAllText(tempFilePath, pdfPath);

                    Console.WriteLine("PDF saved to: " + pdfPath);
                    Console.WriteLine("PDF path saved to: " + tempFilePath);
                    
                    // === Cleanup ===
                    doc.Close(false);
                    wordApp.Quit();
                    File.Delete(tempDocPath);
    
                    Console.WriteLine("Receipt successfully generated.");
                }
                else
                {
                    Util.CallWordCleaner(tempDocPath, tempFilePath);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
