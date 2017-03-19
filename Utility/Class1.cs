using System;
using System.Diagnostics;
using System.Linq;
using System.Net.Mime;
using Word = Microsoft.Office.Interop.Word;

namespace Utility
{
    public static class MainTimer
    {
        public static void onElapsed (Word.Application state, System.Timers.ElapsedEventArgs e)
        {
            if (state.Documents.Count > 0)
            {
                Debug.WriteLine("Woof!");
                var doc = state.ActiveDocument;
                Word.InlineShapes shps;
                Word.Paragraphs pars;
                try
                {
                    pars = doc.Paragraphs;
                }
                catch (Exception)
                {
                    return;
                }
                var pars2 = pars.Cast<Word.Paragraph>()
                    .Where(p => p.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText)
                    .Select(p => p) // do stuff with the selected parragraphs...
                    .ToList();
            }
        }
    }
}
