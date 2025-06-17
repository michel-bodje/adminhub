using System;
using System.IO;
using System.Runtime.InteropServices;

public static class Util
{
    // ---- Client Utilities ----

    /// <summary>
    /// Utility function to validate an email address.
    /// </summary>
    /// <param name="email">The email address to validate.</param>
    /// <returns>True if valid, false otherwise.</returns>
    public static bool IsValidEmail(string email)
    {
        if (string.IsNullOrWhiteSpace(email))
            return false;

        // Simple regex for email validation
        var emailRegex = new System.Text.RegularExpressions.Regex(@"^[^\s@]+@[^\s@]+\.[^\s@]+$");
        return emailRegex.IsMatch(email.ToLowerInvariant());
    }

    /// <summary>
    /// Utility function to validate a phone number (with optional region).<para/>
    /// 
    /// Only supports Canadian/US numbers (1-10 digits) or international numbers (8-15 digits).
    /// </summary>
    /// <param name="number">The phone number to validate.</param>
    /// <param name="region">The region to validate the number against (default: "CA").</param>
    /// <returns>True if valid, false otherwise.</returns>
    public static bool IsValidPhoneNumber(string number, string region = "CA")
    {
        if (string.IsNullOrWhiteSpace(number))
            return false;

        // Remove all non-digit characters
        string digitsOnly = System.Text.RegularExpressions.Regex.Replace(number, @"\D", "");

        // If the number starts with 1 and has 10+ digits, keep it
        if (digitsOnly.Length == 10)
        {
            // Canadian/US numbers are usually 10 digits (without the country code)
            digitsOnly = "1" + digitsOnly;
        }

        // Now validate: must start with non-zero and have 8–15 digits total
        return System.Text.RegularExpressions.Regex.IsMatch(digitsOnly, @"^[1-9]\d{7,14}$");
    }

    /// <summary>
    /// Formats a phone number for display.<para/>
    ///
    /// Tries to format the number intelligently based on its length and region.
    /// </summary>
    /// <param name="number">The phone number to format.</param>
    /// <param name="region">The region to format the number for (default: "CA").</param>
    /// <returns>The formatted phone number.</returns>
    public static string FormatPhoneNumber(string number, string region = "CA")
    {
        if (string.IsNullOrWhiteSpace(number))
            return number;

        // Remove all non-digit characters
        string digitsOnly = System.Text.RegularExpressions.Regex.Replace(number, @"\D", "");

        // Assume Canadian if only 10 digits
        if (digitsOnly.Length == 10)
        {
            return string.Format("{0}-{1}-{2}", digitsOnly.Substring(0, 3), digitsOnly.Substring(3, 3), digitsOnly.Substring(6, 4));
        }
        // If it starts with 1 and has 11 digits (Canadian with country code)
        else if (digitsOnly.Length == 11 && digitsOnly.StartsWith("1"))
        {
            return "+1 " + digitsOnly.Substring(1, 3) + "-" + digitsOnly.Substring(4, 3) + "-" + digitsOnly.Substring(7, 4);
        }
        // If it starts with another country code
        else if (digitsOnly.Length >= 11)
        {
            return "+" + digitsOnly;
        }

        // If invalid, return as-is
        return number;
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
        string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..");
        Console.WriteLine("User skipped PDF export. Scheduling temp cleanup...");
        Console.WriteLine("Word document remains open for editing.");
        // If user cancels, delete the temp file later
        // Schedule cleanup using a separate executable
        string cleanerExe = Path.Combine(baseDir, "app", "office", "bin", "cleanTempDoc.exe");
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
