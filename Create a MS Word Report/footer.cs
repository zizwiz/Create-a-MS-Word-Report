using System;
using Word = Microsoft.Office.Interop.Word;


namespace Create_a_MS_Word_Report
{
    public partial class Form1
    {
        //There are more items than I show here but this will get you started.
        private void CreateFooter(Word._Document word_doc) //Header colours are not the colour you choose unless you are in the header
        {
            // choose the font colour
            Word.WdColorIndex[] footerFontColour = {Word.WdColorIndex.wdBlack, Word.WdColorIndex.wdBlue, Word.WdColorIndex.wdBrightGreen,
                Word.WdColorIndex.wdDarkBlue, Word.WdColorIndex.wdDarkRed, Word.WdColorIndex.wdDarkYellow, Word.WdColorIndex.wdGray25,
                Word.WdColorIndex.wdGray50, Word.WdColorIndex.wdGreen, Word.WdColorIndex.wdPink, Word.WdColorIndex.wdRed, Word.WdColorIndex.wdTeal,
                Word.WdColorIndex.wdTurquoise, Word.WdColorIndex.wdViolet, Word.WdColorIndex.wdWhite, Word.WdColorIndex.wdYellow};

            // choose the background and/or foreground colour
            Word.WdColor[] footerBackgroundPatternColor = { Word.WdColor.wdColorAqua, Word.WdColor.wdColorAutomatic,
                Word.WdColor.wdColorBlack, Word.WdColor.wdColorBlue, Word.WdColor.wdColorBlueGray, Word.WdColor.wdColorBrightGreen,
                Word.WdColor.wdColorBrown, Word.WdColor.wdColorDarkBlue, Word.WdColor.wdColorDarkGreen, Word.WdColor.wdColorDarkRed,
                Word.WdColor.wdColorDarkTeal, Word.WdColor.wdColorDarkYellow, Word.WdColor.wdColorGold, Word.WdColor.wdColorGray05,
                Word.WdColor.wdColorGray10, Word.WdColor.wdColorGray125, Word.WdColor.wdColorGray15, Word.WdColor.wdColorGray20,
                Word.WdColor.wdColorGray25, Word.WdColor.wdColorGray30, Word.WdColor.wdColorGray35, Word.WdColor.wdColorGray375,
                Word.WdColor.wdColorGray40, Word.WdColor.wdColorGray45, Word.WdColor.wdColorGray50, Word.WdColor.wdColorGray55,
                Word.WdColor.wdColorGray60, Word.WdColor.wdColorGray625, Word.WdColor.wdColorGray65, Word.WdColor.wdColorGray70,
                Word.WdColor.wdColorGray75, Word.WdColor.wdColorGray80, Word.WdColor.wdColorGray85, Word.WdColor.wdColorGray875,
                Word.WdColor.wdColorGray90, Word.WdColor.wdColorGray95, Word.WdColor.wdColorGreen, Word.WdColor.wdColorIndigo,
                Word.WdColor.wdColorLavender, Word.WdColor.wdColorLightBlue, Word.WdColor.wdColorLightGreen,
                Word.WdColor.wdColorLightOrange, Word.WdColor.wdColorLightTurquoise, Word.WdColor.wdColorLightYellow,
                Word.WdColor.wdColorLime, Word.WdColor.wdColorOliveGreen, Word.WdColor.wdColorOrange, Word.WdColor.wdColorPaleBlue,
                Word.WdColor.wdColorPink, Word.WdColor.wdColorPlum, Word.WdColor.wdColorRed, Word.WdColor.wdColorRose,
                Word.WdColor.wdColorSeaGreen, Word.WdColor.wdColorSkyBlue, Word.WdColor.wdColorTeal, Word.WdColor.wdColorTurquoise,
                Word.WdColor.wdColorViolet, Word.WdColor.wdColorWhite, Word.WdColor.wdColorYellow };

            //choose the underlining style
            Word.WdUnderline[] footerUnderlineStyle = { Word.WdUnderline.wdUnderlineDash, Word.WdUnderline.wdUnderlineDashHeavy,
                Word.WdUnderline.wdUnderlineDashLong, Word.WdUnderline.wdUnderlineDashLongHeavy,
                Word.WdUnderline.wdUnderlineDotDash, Word.WdUnderline.wdUnderlineDotDashHeavy,
                Word.WdUnderline.wdUnderlineDotDotDash, Word.WdUnderline.wdUnderlineDotDotDashHeavy,
                Word.WdUnderline.wdUnderlineDotted, Word.WdUnderline.wdUnderlineDottedHeavy,
                Word.WdUnderline.wdUnderlineDouble, Word.WdUnderline.wdUnderlineNone,
                Word.WdUnderline.wdUnderlineSingle, Word.WdUnderline.wdUnderlineThick,
                Word.WdUnderline.wdUnderlineWavy, Word.WdUnderline.wdUnderlineWavyDouble,
                Word.WdUnderline.wdUnderlineWavyHeavy, Word.WdUnderline.wdUnderlineWords };

            //Add the footers into the document  
            foreach (Word.Section wordSection in word_doc.Sections)
            {
                ////Get the footer range and add the footer details.  
                //Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //footerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkRed;
                //footerRange.Font.Size = 10;
                //footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                // footerRange.Text = "Footer text goes here";

                WinWord.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
                WinWord.Selection.TypeParagraph();

                WinWord.Selection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                WinWord.ActiveWindow.Selection.Font.Name = "Arial";
                WinWord.ActiveWindow.Selection.Font.Size = 8;

                WinWord.ActiveWindow.Selection.TypeText("Page ");
                Object CurrentPage = Word.WdFieldType.wdFieldPage;
                WinWord.ActiveWindow.Selection.Fields.Add(WinWord.Selection.Range, ref CurrentPage, ref oMissing, ref oMissing);
                WinWord.ActiveWindow.Selection.TypeText(" of ");
                Object TotalPages = Word.WdFieldType.wdFieldNumPages;
                WinWord.ActiveWindow.Selection.Fields.Add(WinWord.Selection.Range, ref TotalPages, ref oMissing, ref oMissing);

                WinWord.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            }
        }
    }
}