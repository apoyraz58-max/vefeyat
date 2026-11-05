using System;
using Vefeyat;

// No Word process or document is created by these tests. Word integration still
// needs the Windows acceptance checks documented in KULLANIM.md.
internal static class SelectionTests
{
    private static int assertions;

    public static int Main()
    {
        try
        {
            NamesAndDates();
            CapturedSelections();
            DuplicateDates();
            Console.WriteLine("Selection checks passed: " + assertions);
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            return 1;
        }
    }

    private static void NamesAndDates()
    {
        SelectionRules.ValidateName("İbn Sînâ");
        SelectionRules.ValidateName("أبو حامد الغزالي");
        SelectionRules.ValidateName("Abaka \u00a0");
        Reject(delegate { SelectionRules.ValidateName(""); }, "Empty selection");
        Reject(delegate { SelectionRules.ValidateName(" \u00a0"); }, "Whitespace selection");
        Reject(delegate { SelectionRules.ValidateName("Abaka\r"); }, "Paragraph selected");
        Reject(delegate { SelectionRules.ValidateName("Abaka\a"); }, "Table cell marker selected");
        Reject(delegate { SelectionRules.ValidateName("Abaka\tAhmed"); }, "Tab-separated names");
        Reject(delegate { SelectionRules.ValidateName(new String('a', 201)); }, "Excessively long name");

        string[] historicalDates = { "(ö. 680/1282)", "(m.ö. 384-322)", "(gaybeti 260/874)", "(ö. 965 [?])" };
        foreach (string value in historicalDates)
            Check(SelectionRules.ValidateDate(value) == value, "Historical date must remain unchanged: " + value);
        Reject(delegate { SelectionRules.ValidateDate(null); }, "Missing date");
        Reject(delegate { SelectionRules.ValidateDate("(ö. 680)\rMetin"); }, "Multiline date");
        Reject(delegate { SelectionRules.ValidateDate("(ö. 680)\a"); }, "Date cannot insert a cell marker");
        Check(SelectionRules.TrailingSpaces("Abaka \u00a0") == 2, "Insert before Word's selected trailing spaces");
        Check(SelectionRules.TrailingSpaces("Abaka") == 0, "Do not move insertion into the name");
    }

    private static void CapturedSelections()
    {
        Check(SelectionRules.SameSelection("Abaka", "Abaka", "Abaka", 7, 12, 7, 12, 1, 1), "Unchanged selection");
        Check(!SelectionRules.SameSelection("Abaka", "Abaka", "Abaka", 7, 12, 20, 25, 1, 1),
            "An identical name at another location must not become the target");
        Check(!SelectionRules.SameSelection("Abaka", "Abaka", "Abaka", 7, 12, 7, 12, 1, 2),
            "An identical range in another story must not become the target");
        Check(!SelectionRules.SameSelection("Abaka", "Ahmed", "Ahmed", 7, 12, 7, 12, 1, 1),
            "A tracked range whose contents changed must be rejected");
        Check(!SelectionRules.SameSelection("Abaka", "Abaka", "abaka", 7, 12, 7, 12, 1, 1),
            "Selection comparison must preserve exact document text");
        Check(SelectionRules.SameSelection("Abaka", "Abaka", "Abaka", 80, 85, 80, 85, 1, 1),
            "A tracked range may shift when surrounding document text changes");
    }

    private static void DuplicateDates()
    {
        const string date = "(ö. 680/1282)";
        Check(SelectionRules.HasDate("Abaka", " " + date + " hükümdar", date), "Date immediately after name");
        Check(SelectionRules.HasDate("Abaka", "\u00a0 (ö.   680/1282)", date), "Whitespace variations in existing date");
        Check(SelectionRules.HasDate("Abaka " + date, "", date), "Date already inside selection");
        Check(!SelectionRules.HasDate("Abaka", " hükümdar; " + date, date), "Unrelated later date must not block insertion");
        Check(!SelectionRules.HasDate("Abaka", " (ö. 681/1283)", date), "Different date must remain distinguishable");
        Check(!SelectionRules.HasDate("Abaka", null, date), "End of story with no following date");
    }

    private static void Check(bool condition, string message)
    {
        assertions++;
        if (!condition) throw new Exception("FAIL: " + message);
    }

    private static void Reject(Action action, string message)
    {
        try { action(); }
        catch (InvalidOperationException) { assertions++; return; }
        throw new Exception("FAIL: " + message + " was accepted");
    }
}
