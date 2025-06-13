using System;
public static class Util
{
    public static double AddTaxes(double amount, bool addFOF = false)
    {
        if (double.IsNaN(amount))
            throw new ArgumentException("Amount is not a valid number.");

        double total = amount * (1 + 0.05 + 0.09975); // GST + QST
        if (addFOF) total += 100; // File Opening Fee
        return total;
    }

    public static void ReplaceWordText(dynamic doc, string placeholder, string replacement)
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
    
    public static void HyperlinkEmail(dynamic doc, string clientEmail)
    {
        dynamic range = doc.Content;
        dynamic found = range.Find;
        found.Text = "{clientEmail}";
        if (found.Execute())
        {
            range.Hyperlinks.Add(range, "mailto:" + clientEmail, Type.Missing, Type.Missing, clientEmail);
        }
    }
}
