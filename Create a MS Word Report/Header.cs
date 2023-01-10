using Word = Microsoft.Office.Interop.Word;


namespace Create_a_MS_Word_Report
{
    public partial class Form1
    {

        private void CreateHeader(Word._Document word_doc) //Header colours are not the colour you choose unless you are in the header
        {
            Word.WdColorIndex headerFontColour;

            switch (cmbobx_header_colour.SelectedIndex)
            {
                case 0:
                    headerFontColour = Word.WdColorIndex.wdBlack;
                    break;
                case 1:
                    headerFontColour = Word.WdColorIndex.wdBlue;
                    break;
                case 2:
                    headerFontColour = Word.WdColorIndex.wdBrightGreen;
                    break;
                case 3:
                    headerFontColour = Word.WdColorIndex.wdBrightGreen;
                    break;
                case 4:
                    headerFontColour = Word.WdColorIndex.wdDarkBlue;
                    break;
                case 5:
                    headerFontColour = Word.WdColorIndex.wdDarkRed;
                    break;
                case 6:
                    headerFontColour = Word.WdColorIndex.wdDarkYellow;
                    break;
                case 7:
                    headerFontColour = Word.WdColorIndex.wdGray25;
                    break;
                case 8:
                    headerFontColour = Word.WdColorIndex.wdGray50;
                    break;
                case 9:
                    headerFontColour = Word.WdColorIndex.wdGreen;
                    break;
                case 10:
                    headerFontColour = Word.WdColorIndex.wdPink;
                    break;
                case 11:
                    headerFontColour = Word.WdColorIndex.wdRed;
                    break;
                case 12:
                    headerFontColour = Word.WdColorIndex.wdTeal;
                    break;
                case 13:
                    headerFontColour = Word.WdColorIndex.wdTurquoise;
                    break;
                case 14:
                    headerFontColour = Word.WdColorIndex.wdViolet;
                    break;
                case 15:
                    headerFontColour = Word.WdColorIndex.wdWhite;
                    break;
                case 16:
                    headerFontColour = Word.WdColorIndex.wdYellow;
                    break;
                default:
                    headerFontColour = Word.WdColorIndex.wdBlack;
                    break;
            }


            //Add header into the document  
            foreach (Word.Section section in word_doc.Sections)
            {
                //Get the header range and add the header details.  
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = headerFontColour; //font colour
                headerRange.Font.Size = float.Parse(cmbobx_header_fontsize.SelectedItem.ToString()); //size of font
                headerRange.Font.Bold = (int)Word.WdConstants.wdToggle; // Toggle the title to a Bold Font
                headerRange.Text = "Project: " + txtbx_proj_name.Text + " Report";
            }
        }






    }
}