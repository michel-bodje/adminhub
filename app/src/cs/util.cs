using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;

public static class Util
{
    // ---- Client Utilities ----

    public static string RootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\");
    public static string TemplateDir = Path.Combine(RootDir, "app", "templates");

    public static JObject ReadJsonFromStdin()
    {
        string input = Console.In.ReadToEnd();
        if (string.IsNullOrWhiteSpace(input))
            throw new Exception("No input received via stdin.");
        return JObject.Parse(input);
    }

    // ---- Contract Utilities ----

    /// <summary>
    /// Adds GST and QST taxes to the specified amount, with an optional file opening fee.
    /// </summary>
    /// <param name="amount">The amount to which taxes will be added.</param>
    /// <param name="addFOF">If set to <c>true</c>, includes a file opening fee of $100.</param>
    /// <returns>The total amount including taxes and optional fees.</returns>
    /// <exception cref="ArgumentException">Thrown when the specified amount is not a valid number.</exception> 
    public static double AddTaxes(double amount, bool addFOF = false)
    {
        if (double.IsNaN(amount))
            throw new ArgumentException("Amount is not a valid number.");

        double total = amount * (1 + 0.05 + 0.09975); // GST + QST
        if (addFOF) total += 100; // File Opening Fee
        return total;
    }

    /// <summary>
    /// Formats the lawyer's name with a title prefix based on their ID.
    /// </summary>
    /// <param name="name">The name of the lawyer.</param>
    /// <param name="id">The ID of the lawyer.</param>
    /// <returns>
    /// The formatted lawyer's name prefixed with "Me" if the ID is not one of the non-lawyer staff; otherwise, returns the name as is.
    /// </returns>
    public static string GetLawyerString(string name, string id)
    {
        if (id != "AR" && id != "MG" && id != "PM")
        {
            return "Me " + name;
        }
        return name;
    }

    // ---- Word Document Utilities ----

    /// <summary>
    /// Replaces all occurrences of a placeholder string with a replacement string
    /// in the specified Word document.
    /// </summary>
    /// <param name="doc">The Word document to modify.</param>
    /// <param name="placeholder">The placeholder string to replace.</param>
    /// <param name="replacement">The replacement string.</param>
    public static void WordReplaceText(dynamic doc, string placeholder, string replacement)
    {
        dynamic range = doc.Content;
        dynamic find = range.Find;
        find.Text = placeholder;
        find.Replacement.Text = replacement;
        find.Forward = true;
        find.Wrap = 1; // wdFindContinue
        find.Format = false;
        find.MatchCase = false;
        find.MatchWholeWord = false;
        find.MatchWildcards = false;
        find.MatchSoundsLike = false;
        find.MatchAllWordForms = false;
        find.Execute(Replace: 2); // wdReplaceAll
    }


    /// <summary>
    /// Replaces all occurrences of a placeholder string in the specified Word document
    /// with a hyperlink to the specified email address.
    /// </summary>
    /// <param name="doc">The Word document to modify.</param>
    /// <param name="placeholder">The placeholder string to replace.</param>
    /// <param name="clientEmail">The client email address to hyperlink to.</param>
    public static void WordHyperlinkEmail(dynamic doc, string placeholder, string clientEmail)
    {
        dynamic range = doc.Content;
        dynamic found = range.Find;
        found.Text = placeholder;
        if (found.Execute())
        {
            range.Hyperlinks.Add(range, "mailto:" + clientEmail, Type.Missing, Type.Missing, clientEmail);
        }
    }

    public static void CallWordCleaner(string tempDocPath, string tempFilePath = "")
    {
        Console.WriteLine("User skipped PDF export. Scheduling temp cleanup...");
        Console.WriteLine("Word document remains open for editing.");
        // If user cancels, delete the temp file later
        // Schedule cleanup using a separate executable
        string cleanerExe = Path.Combine(RootDir, "app", "bin", "cleanTempDoc.exe");
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
        if (File.Exists(tempFilePath))
            File.Delete(tempFilePath);
    }
}

public static class WindowFocus
{
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private const int SW_RESTORE = 9;

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("00000114-0000-0000-C000-000000000046")]
    private interface IOleWindow
    {
        void GetWindow(out IntPtr hwnd);
        void ContextSensitiveHelp(bool fEnterMode);
    }

    public static void ShowWithFocus(object comWindow)
    {
        if (comWindow == null) return;

        try
        {
            var oleWindow = (IOleWindow)comWindow;
            IntPtr hwnd;
            oleWindow.GetWindow(out hwnd);
            if (hwnd != IntPtr.Zero)
            {
                ShowWindow(hwnd, SW_RESTORE);
                SetForegroundWindow(hwnd);
            }
        }
        catch
        {
            // fallback or ignore
        }
    }
}
