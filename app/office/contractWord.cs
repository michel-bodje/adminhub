using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

class Program
{
    [STAThread] // Required for SaveFileDialog
    static void Main()
    {
        dynamic wordApp = null;
        dynamic doc = null;
        string tempDocPath = null;

        try
        {
            // === Load JSON ===
            string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..");
            string jsonPath = Path.Combine(baseDir, "app", "data.json");
            JObject json = JObject.Parse(File.ReadAllText(jsonPath));

            string clientName = (string)json["form"]["clientName"];
            string clientEmail = (string)json["form"]["clientEmail"];
            string clientLanguage = (string)json["form"]["clientLanguage"];
            string contractTitle = (string)json["form"]["contractTitle"];
            double depositAmount = (double)json["form"]["depositAmount"];

            string lang = clientLanguage == "Français" ? "fr" : "en";
            string locale = lang == "fr" ? "fr-CA" : "en-US";
            string templatePath = Path.Combine(baseDir, "templates", lang, "Contract.docx");

            if (!File.Exists(templatePath))
                throw new FileNotFoundException("Contract template not found.", templatePath);

            // === Copy template to temp location ===
            tempDocPath = Path.Combine(Path.GetTempPath(), "Contract_" + Guid.NewGuid() + ".docx");
            File.Copy(templatePath, tempDocPath, true);

            // === Start Word and open the document ===
            Type wordType = Type.GetTypeFromProgID("Word.Application");
            wordApp = Activator.CreateInstance(wordType);
            doc = wordApp.Documents.Open(tempDocPath);

            // === Replace placeholders ===
            double totalAmount = Util.AddTaxes(depositAmount, true);
            string today = DateTime.Today.ToString("D", CultureInfo.CreateSpecificCulture(locale));

            Util.ReplaceWordText(doc, "{clientName}", clientName);
            Util.ReplaceWordText(doc, "{contractTitle}", contractTitle);
            Util.ReplaceWordText(doc, "{depositAmount}", depositAmount.ToString("F0"));
            Util.ReplaceWordText(doc, "{totalAmount}", totalAmount.ToString("F2"));
            Util.ReplaceWordText(doc, "{date}", today);
            Util.HyperlinkEmail(doc, clientEmail);
            
            wordApp.Visible = true;

            // === Prompt user for PDF path ===
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "PDF files (*.pdf)|*.pdf";
                dialog.Title = "Save Contract as PDF";
                dialog.FileName = lang == "fr"
                    ? "Contrat de services_" + clientName.Replace(" ", "-") + "_" + DateTime.Today.ToString("yyyy-MM-dd") + ".pdf"
                    : "Contract of services_" + clientName.Replace(" ", "-") + "_" + DateTime.Today.ToString("yyyy-MM-dd") + ".pdf";

                string projectPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..");
                string pathTempFile = Path.Combine(projectPath, "app", "latest_contract_path.txt");

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string pdfPath = dialog.FileName;
                    doc.ExportAsFixedFormat(pdfPath, 17); // 17 = wdExportFormatPDF

                    // Save the PDF path to temp file
                    File.WriteAllText(pathTempFile, pdfPath);

                    Console.WriteLine("PDF saved to: " + pdfPath);
                    Console.WriteLine("PDF path saved to: " + pathTempFile);

                    // === Cleanup ===
                    doc.Close(false);
                    wordApp.Quit();
                    File.Delete(tempDocPath);

                    Console.WriteLine("Contract successfully generated.");
                }
                else
                {
                    Console.WriteLine("User skipped PDF export. Scheduling temp cleanup...");
                    Console.WriteLine("Word document remains open for editing.");
                    // If user cancels, delete the temp file later
                    // Schedule cleanup using a separate executable

                    string cleanerExe = Path.Combine(baseDir, "cleanTempDoc.exe");
                    if (File.Exists(cleanerExe))
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = cleanerExe,
                            Arguments = "\"" + tempDocPath + "\"",
                            UseShellExecute = false,
                            CreateNoWindow = true
                        });
                    }
                    else
                    {
                        Console.WriteLine("Cleanup exe not found.");
                    }
                    
                    // Delete path record if user cancels
                    if (File.Exists(pathTempFile))
                        File.Delete(pathTempFile);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            // Cleanup COM
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
