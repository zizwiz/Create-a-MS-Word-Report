using Word = Microsoft.Office.Interop.Word;


namespace Create_a_MS_Word_Report
{
    public partial class Form1
    {
        //There are more items than I show here but this will get you started.
        private void CreateHeader(Word._Document word_doc) //Header colours are not the colour you choose unless you are in the header
        {
           // choose the font colour
            Word.WdColorIndex[] headerFontColour = {Word.WdColorIndex.wdBlack, Word.WdColorIndex.wdBlue, Word.WdColorIndex.wdBrightGreen,
                Word.WdColorIndex.wdDarkBlue, Word.WdColorIndex.wdDarkRed, Word.WdColorIndex.wdDarkYellow, Word.WdColorIndex.wdGray25,
                Word.WdColorIndex.wdGray50, Word.WdColorIndex.wdGreen, Word.WdColorIndex.wdPink, Word.WdColorIndex.wdRed, Word.WdColorIndex.wdTeal,
                Word.WdColorIndex.wdTurquoise, Word.WdColorIndex.wdViolet, Word.WdColorIndex.wdWhite, Word.WdColorIndex.wdYellow};

            // choose the background and/or foreground colour
            Word.WdColor[] headerBackgroundPatternColor = { Word.WdColor.wdColorAqua, Word.WdColor.wdColorAutomatic, 
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
                headerRange.Font.ColorIndex = headerFontColour[cmbobx_header_font_colour.SelectedIndex];    //font colour
                
                headerRange.Shading.BackgroundPatternColor = headerBackgroundPatternColor[cmbobx_header_background_colour.SelectedIndex];
                headerRange.Shading.ForegroundPatternColor = headerBackgroundPatternColor[cmbobx_header_background_colour.SelectedIndex];
                
                headerRange.Font.Name =  cmbobx_header_fontname.SelectedItem.ToString(); //font name
                headerRange.Font.Size = float.Parse(cmbobx_header_fontsize.SelectedItem.ToString()); //size of font
                
                headerRange.Font.Bold = chkbx_header_bold.Checked? 1:0 ; 
                headerRange.Font.Italic = chkbx_header_italic.Checked? 1:0; //(int)Word.WdConstants.wdToggle;
                headerRange.Font.StrikeThrough = rdobtn_header_single_strikethrough.Checked? 1:0;
                headerRange.Font.DoubleStrikeThrough = rdobtn_header_double_strikethrough.Checked? 1:0;
                headerRange.Font.AllCaps = chkbx_header_all_caps.Checked? 1:0;
                headerRange.Font.Emboss = chkbx_header_emboss.Checked? 1:0;
                headerRange.Font.Engrave = chkbx_header_engrave.Checked? 1:0;
                headerRange.Font.Outline = chkbx_header_outline.Checked? 1:0;
                headerRange.Font.Shadow = chkbx_header_shadow.Checked? 1:0;
               
                headerRange.Font.Underline = headerUnderlineStyle[cmbobx_header_underline_style.SelectedIndex]; //choose type of underlining
                
                headerRange.Text = "Project: " + txtbx_proj_name.Text + " Report";
            }
        }






    }
}