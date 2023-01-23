using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word; // now got to ref and in properties set "Embed Interop types" to false



// Get Tools from:
//Programs and Features --> Select Visual Studio > Change
//Choose Modify
//Windows and Webdevelopment --> Tick/select "Office Developer Tools for Visual Studio"
//Start Update
//After this go to Project | Add Reference | Assemblies | Extensions and now add | Microsoft.Office.Tools.Word



namespace Create_a_MS_Word_Report
{
    public partial class Form1 : Form
    {
        private object oMissing = Missing.Value;
        private object oTrue = true;
        private object oFalse = false;
        private object oTemplatePath = "C:\\Users\\itobo\\source\\repos\\Create-a-MS-Word-Report\\Create a MS Word Report\\bin\\Debug\\My text.com";

        Word.Application WinWord = new Word.Application(); // create a word object and show it.
        Word.Document word_doc = new Word.Document();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SetUp();

            Text += " : v" + Assembly.GetExecutingAssembly().GetName().Version; // put in the version number


        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btn_create_Click(object sender, EventArgs e)
        {
            // object oMissing = Missing.Value;

            //This uses the styles from your default-default word doc template
            //change them below if you have your own template



            WinWord.Visible = true; //Set status for word application is to be visible or not.
            WinWord.ShowAnimation = false; //Set animation status for word application

            // Create the Word document, choose your template here or you will get the default-default one.
            // Add(Template, New Template, DocType, Visible).
            word_doc = WinWord.Documents.Add(
                ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);


            if (chkbx_page_header.Checked) //Only create if we ticked to say so
            {
                CreateHeader(word_doc);
            }

            if (chkbx_page_footer.Checked)
            {
                CreateFooter(word_doc);
            }




           // // Report Title.
           // Word.Paragraph para0 = word_doc.Paragraphs.Add(oMissing);
           // para0.Range.Text = "Project: " + txtbx_proj_name.Text + " Report";
           // para0.Range.set_Style("Title");
           // para0.Range.InsertParagraphAfter();


           // Word.Paragraph paraA = word_doc.Paragraphs.Add(oMissing);
           // //Create bookmark
           //// paraA.Range.Bookmarks.Add("bookmark1", oMissing);
           // paraA.Range.Text = "Before replacement";
           //paraA.Range.InsertParagraphAfter();





            ReplaceBookmarkText(word_doc, "bookmark1", "Hello");
            ReplaceBookmarkText(word_doc, "bookmark2", "Bottom one");
            ReplaceBookmarkText(word_doc, "bookmark3", "Top Banana");

            string[] items = new string[] {};
            foreach (Word.Bookmark bookmark in word_doc.Bookmarks)
            {
                items = new List<string>(items) { bookmark.Name }.ToArray();
            }


            int count = word_doc.Bookmarks.Count;
            for (int i = 1; i < count + 1; i++)
            {
                object oRange = word_doc.Bookmarks[i].Range;
                object saveWithDocument = true;
                object missing = Type.Missing;
                string pictureName =
                    @"C:\\Users\\itobo\\source\repos\\Create-a-MS-Word-Report\\Create a MS Word Report\\bin\\Debug\\plane.png";

                if (items[i-1] == "Picture1")
                {
                    cleanBookmark("Picture1"); //
                    word_doc.InlineShapes.AddPicture(pictureName, ref missing, ref saveWithDocument, ref oRange);
                }
            }

            //Word.Bookmarks books = word_doc.Bookmarks;
            //Word.Range rngImage = books["Picture1"].Range;
            //rngImage.InlineShapes.AddPicture(@"C:\\Users\\itobo\\source\repos\\Create-a-MS-Word-Report\\Create a MS Word Report\\bin\\Debug\\plane.png");
            //rngImage.InlineShapes[2].Delete();


            /*

            // Create a header.
            Word.Paragraph para1 = word_doc.Paragraphs.Add(oMissing);
            para1.Range.set_Style("Heading 1");
            para1.Range.Text = "Income";
            para1.Range.InsertParagraphAfter();

            //add text to paragraph not I am also using CR and LF
            Word.Paragraph para2 = word_doc.Paragraphs.Add(oMissing);
            para2.Range.set_Style("Normal");
            para2.Range.Text = "Loads and Loads of text More much more";
            para2.Range.InsertParagraphAfter();

            WinWord.Visible = false; //Make invisible so we do not need to keep redrawing the whole document.
            //Create a 5X5 table and insert some dummy record 
            Word.Paragraph para3 = word_doc.Paragraphs.Add(oMissing);
            Word.Table myTable = word_doc.Tables.Add(para3.Range, 5, 5, oMissing, oMissing);

            myTable.Borders.Enable = 1;
            foreach (Word.Row row in myTable.Rows)
            {
                foreach (Word.Cell cell in row.Cells)
                {
                    //Header row  
                    if (cell.RowIndex == 1)
                    {
                        cell.Range.Text = "Column " + cell.ColumnIndex;
                        cell.Range.Font.Bold = 1;
                        //other format properties goes here  
                        cell.Range.Font.Name = "verdana";
                        cell.Range.Font.Size = 10;
                        cell.Range.Font.ColorIndex = Word.WdColorIndex.wdBrightGreen;
                        cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
                        //Center alignment for the Header cells  
                        cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    //Data row  
                    else
                    {
                        cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                    }
                }
            }

            myTable.Columns[1].Width = WinWord.InchesToPoints(1); //Change width of columns 1 & 2
            myTable.Columns[2].Width = WinWord.InchesToPoints(1);
            para3.Range.InsertParagraphAfter();

            WinWord.Visible = true;
            Word.Paragraph para5 = word_doc.Paragraphs.Add(oMissing);
            string picture_file =
                @"C:\\Users\\itobo\\source\repos\\Create-a-MS-Word-Report\\Create a MS Word Report\\bin\\Debug\\mypic.png";

            // Add the picture to the paragraph.
            Word.InlineShape inline_shape = para5.Range.InlineShapes.AddPicture(
                picture_file, oMissing, oMissing, oMissing);

            para5.Range.InsertParagraphAfter();

            

            //Insert a chart.
            object oEndOfDoc = "\\endofdoc"; \\endofdoc is a predefined bookmark 
            object oRng = word_doc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            Word.Table oTable;
            Word.Range wrdRng = word_doc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            Word.Paragraph para6 = word_doc.Paragraphs.Add(Type.Missing);
            Word.InlineShape oShape;
            object oClassType = "MSGraph.Chart.8";
            wrdRng = word_doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing);

            //Demonstrate use of late bound oChart and oChartApp objects to
            //manipulate the chart object with MSGraph.
            object oChart;
            object oChartApp;
            oChart = oShape.OLEFormat.Object;
            oChartApp = oChart.GetType().InvokeMember("Application",
                BindingFlags.GetProperty, null, oChart, null);

            //Change the chart type to Line.
            object[] Parameters = new Object[1];
            Parameters[0] = 4; //xlLine = 4; pie = 5;
            oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty,
                null, oChart, Parameters);

            //Update the chart image and quit MSGraph.
            oChartApp.GetType().InvokeMember("Update",
                BindingFlags.InvokeMethod, null, oChartApp, null);
            oChartApp.GetType().InvokeMember("Quit",
                BindingFlags.InvokeMethod, null, oChartApp, null);
            //... If desired, you can proceed from here using the Microsoft Graph 
            //Object model on the oChart and oChartApp objects to make additional
            //changes to the chart.

            //Set the width of the chart.
            oShape.Width = WinWord.InchesToPoints(6.25f);
            oShape.Height = WinWord.InchesToPoints(3.57f);

            para6.Range.InsertParagraphAfter();






            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PDF |*.pdf";
            saveFileDialog.Title = "Save Document";
            saveFileDialog.DefaultExt = ".pdf";
            saveFileDialog.FileName = "Sample.pdf";

            DialogResult result = saveFileDialog.ShowDialog();
            saveFileDialog.RestoreDirectory = true;

            if (result == DialogResult.OK && saveFileDialog.FileName != "")
            {
                try
                {
                    word_doc.SaveAs(saveFileDialog.FileName, Word.WdSaveFormat.wdFormatPDF,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }*/
        }




        public void cleanBookmark(string bookmark)
        {
            var start = word_doc.Bookmarks[bookmark].Start;
            var end = word_doc.Bookmarks[bookmark].End;
            Word.Range range = word_doc.Range(start, end);
            range.Delete();
            //The Delete() only deletes text so if you got tables in the doc it leaves the tables empty. 
            //The following removes the tables in the current range.
            if (range.Tables.Count != 0)
            {
                for (int i = 1; i <= range.Tables.Count; i++)
                {
                    range.Tables[i].Delete();
                }
            }
            word_doc.Bookmarks.Add(bookmark, range);
        }



        //Choose teh template we will be using
        private void btn_choose_doc_template_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                Title = "Open Template Doc.",
                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "dotx",
                Filter = "Template files (*.dotx)|*.dotx",
                FilterIndex = 2,
                RestoreDirectory = true,

            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //The location of the template we will be using
                oTemplatePath = openFileDialog1.FileName;

            }
        }
    }
}
