using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Timers;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Utility;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void MainTimer_Elapsed (object sender, System.Timers.ElapsedEventArgs e)
        {
            var scrnUpdating = Application.ScreenUpdating;
            Application.ScreenUpdating = false;

            if (Application.Documents.Count > 0)
            {
                var doc = Application.ActiveDocument;
                Word.InlineShapes shps;
                Word.Paragraphs pars;
                try
                {
                    pars = doc.Paragraphs;
                }
                catch (Exception)
                {
                    if (scrnUpdating)
                        Application.ScreenUpdating = true;
                    return;
                }
                var pars2 = pars.Cast<Word.Paragraph>().ToList();
                foreach (var obj in pars2)
                {
                    if (obj.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText)//PROBLEM HERE
                    {

                    };
                }
            }
            if (scrnUpdating)
                Application.ScreenUpdating = true;
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Timer timer = new Timer(2000);
            timer.Elapsed += (s, t) =>
            {
                var scrnUpdating = Application.ScreenUpdating;
                Application.ScreenUpdating = false;
                MainTimer.onElapsed(Application, t);
                if (scrnUpdating)
                    Application.ScreenUpdating = true;
            };
            timer.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
