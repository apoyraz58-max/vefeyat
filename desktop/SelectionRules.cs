using System;
using System.Text.RegularExpressions;

namespace Vefeyat
{
    internal static class SelectionRules
    {
        public static void ValidateName(string text)
        {
            if (String.IsNullOrWhiteSpace(text) || text.Trim().Length > 200 || Regex.IsMatch(text, @"[\r\n\t\a\v\f]"))
                throw new InvalidOperationException("Belgede yalnız şahıs adını seçin (en fazla 200 karakter).");
        }

        public static string ValidateDate(string text)
        {
            if (String.IsNullOrWhiteSpace(text) || text.Length > 500 || Regex.IsMatch(text, @"[\r\n\t\a\v\f]"))
                throw new InvalidOperationException("Bu kayıtta eklenebilir tek satırlık bir tarih bulunmuyor.");
            return text.Trim();
        }

        public static bool SameSelection(string original, string tracked, string selected,
            int start, int end, int selectedStart, int selectedEnd, int story, int selectedStory)
        {
            return original == tracked && original == selected && start == selectedStart && end == selectedEnd && story == selectedStory;
        }

        public static int TrailingSpaces(string text)
        {
            return text.Length - text.TrimEnd(' ', '\u00a0').Length;
        }

        public static bool HasDate(string selection, string following, string date)
        {
            string normalizedDate = Regex.Replace(date, @"\s+", " ").Trim();
            return Regex.Replace(selection, @"\s+", " ").Trim().EndsWith(normalizedDate, StringComparison.Ordinal) ||
                Regex.Replace(following ?? "", @"\s+", " ").TrimStart().StartsWith(normalizedDate, StringComparison.Ordinal);
        }
    }
}
