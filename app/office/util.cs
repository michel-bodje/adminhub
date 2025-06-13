namespace LawhubOffice.Util;

using System;
using System.Runtime;
using PhoneNumbers;
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
    /// Utility function to validate an international phone number using libphonenumber.
    /// </summary>
    /// <param name="phone">The phone number to validate.</param>
    /// <returns>True if valid, false otherwise.</returns>
    public static bool IsValidPhoneNumber(string phone)
    {
        if (string.IsNullOrWhiteSpace(phone))
            return false;

        try
        {
            var phoneUtil = PhoneNumberUtil.GetInstance();
            // null region means parse expects a leading + (E.164)
            var number = phoneUtil.Parse(phone, null);
            return phoneUtil.IsValidNumber(number);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Utility function to format a phone number for display using libphonenumber.
    /// </summary>
    /// <param name="phone">The phone number to format.</param>
    /// <returns>The formatted phone number, or the original input if invalid.</returns>
    public static string FormatPhoneNumber(string phone)
    {
        if (string.IsNullOrWhiteSpace(phone))
            return phone;

        try
        {
            var phoneUtil = PhoneNumberUtil.GetInstance();
            var number = phoneUtil.Parse(phone, null);
            // Format as INTERNATIONAL (e.g., +1 555-555-5555)
            return phoneUtil.Format(number, PhoneNumberFormat.INTERNATIONAL);
        }
        catch
        {
            return phone;
        }
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
    public static void WordHyperlinkEmail(dynamic doc, string placeholder , string clientEmail)
    {
        dynamic range = doc.Content;
        dynamic found = range.Find;
        found.Text = placeholder;
        if (found.Execute())
        {
            range.Hyperlinks.Add(range, "mailto:" + clientEmail, Type.Missing, Type.Missing, clientEmail);
        }
    }
}
