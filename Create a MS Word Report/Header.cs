using System;
using Word = Microsoft.Office.Interop.Word;


namespace Create_a_MS_Word_Report
{
    public partial class Form1
    {

        private void CreateHeader(Word._Document word_doc) //Header colours are not the colour you choose unless you are in the header
        {
           // choose the font colour
            Word.WdColorIndex[] headerFontColour = {Word.WdColorIndex.wdBlack, Word.WdColorIndex.wdBlue, Word.WdColorIndex.wdBrightGreen,
                Word.WdColorIndex.wdDarkBlue, Word.WdColorIndex.wdDarkRed, Word.WdColorIndex.wdDarkYellow, Word.WdColorIndex.wdGray25,
                Word.WdColorIndex.wdGray50, Word.WdColorIndex.wdGreen, Word.WdColorIndex.wdPink, Word.WdColorIndex.wdRed, Word.WdColorIndex.wdTeal,
                Word.WdColorIndex.wdTurquoise, Word.WdColorIndex.wdViolet, Word.WdColorIndex.wdWhite, Word.WdColorIndex.wdYellow};
            
            //choose teh underlining style
            Word.WdUnderline[] headerUnderlineStyle = { Word.WdUnderline.wdUnderlineDash, Word.WdUnderline.wdUnderlineDashHeavy, 
                Word.WdUnderline.wdUnderlineDashLong, Word.WdUnderline.wdUnderlineDashLongHeavy, 
                Word.WdUnderline.wdUnderlineDotDash, Word.WdUnderline.wdUnderlineDotDashHeavy, 
                Word.WdUnderline.wdUnderlineDotDotDash, Word.WdUnderline.wdUnderlineDotDotDashHeavy, 
                Word.WdUnderline.wdUnderlineDotted, Word.WdUnderline.wdUnderlineDottedHeavy, 
                Word.WdUnderline.wdUnderlineDouble, Word.WdUnderline.wdUnderlineNone, 
                Word.WdUnderline.wdUnderlineSingle, Word.WdUnderline.wdUnderlineThick, 
                Word.WdUnderline.wdUnderlineWavy, Word.WdUnderline.wdUnderlineWavyDouble, 
                Word.WdUnderline.wdUnderlineWavyHeavy, Word.WdUnderline.wdUnderlineWords };

            

            //Add header into the document  
            foreach (Word.Section section in word_doc.Sections)
            {
                //Get the header range and add the header details.  
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = headerFontColour[cmbobx_header_colour.SelectedIndex];    //font colour
                headerRange.Font.Name =  cmbobx_header_fontname.SelectedItem.ToString(); //font name
                headerRange.Font.Size = float.Parse(cmbobx_header_fontsize.SelectedItem.ToString()); //size of font
                
                headerRange.Font.Bold = chkbx_header_bold.Checked? 1:0 ; 
                headerRange.Font.Italic = chkbx_header_italic.Checked? 1:0; //(int)Word.WdConstants.wdToggle;
                headerRange.Font.StrikeThrough = (int)Word.WdConstants.wdToggle;
                headerRange.Font.AllCaps = 0;
                headerRange.Font.DoubleStrikeThrough = 0;
                headerRange.Font.BoldBi = 0;
                headerRange.Font.Emboss = 0;
                headerRange.Font.Engrave = 0;
                headerRange.Font.ItalicBi = 0;
                headerRange.Font.Outline = 0;
                headerRange.Font.Shadow = 0;
               

                headerRange.Font.Underline = headerUnderlineStyle[cmbobx_header_underline_style.SelectedIndex]; //choose type of underlining



                headerRange.Text = "Project: " + txtbx_proj_name.Text + " Report";
            }
        }






    }
}