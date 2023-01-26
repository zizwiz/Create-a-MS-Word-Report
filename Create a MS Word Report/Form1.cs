using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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

            // find all the bookmarks in the doc and list them out.
            int bkmk_count = 0;
            int bkmk_num = 0;
           // string[] items = new string[] { };
            List<Word.Bookmark> bmarks = new List<Word.Bookmark>();
            foreach (Word.Bookmark bookmark in word_doc.Bookmarks)
            {
                ////add a label to the screen
                //Label namelabel = new Label();
                //namelabel.Location = new Point(100, 100+bkmk_count);
                //namelabel.Text = bookmark.Name;
                //tab_bookmark_update.Controls.Add(namelabel);

                // items = new List<string>(items) { bookmark.Name }.ToArray(); // Add to list
                bmarks.Add(bookmark);
                
            }





            // Re-sort list in order of appearance
            bmarks = bmarks.OrderBy(b => b.Start).ToList(); // LINQ

            List<string> slBMarks = new List<string>();
            foreach (Word.Bookmark b in bmarks)
            {
                //add a label to the screen
                Label namelabel = new Label();
                namelabel.Location = new Point(100, 100 + bkmk_count);
                namelabel.Text = b.Name;
                tab_bookmark_update.Controls.Add(namelabel);
                bkmk_count += 40;
                bkmk_num++;

                slBMarks.Add(b.Name); // Accumulate bookmark names
            }


            cleanBookmark("bookmark1"); // Remove everything at this bookmark so we can replace it
            ReplaceBookmarkText(word_doc, "bookmark1", "Hello");
            cleanBookmark("bookmark2"); // Remove everything at this bookmark so we can replace it
            ReplaceBookmarkText(word_doc, "bookmark2", "Bottom one");
            cleanBookmark("bookmark3"); // Remove everything at this bookmark so we can replace it
            ReplaceBookmarkText(word_doc, "bookmark3", "Top Banana");

            

            //int count = word_doc.Bookmarks.Count;
            //for (int i = 1; i < count + 1; i++)
            //{
            //    object oRange = word_doc.Bookmarks[i].Range;
            //    object saveWithDocument = true;
            //    object missing = Type.Missing;
            //    string pictureName =
            //        @"C:\\Users\\itobo\\source\repos\\Create-a-MS-Word-Report\\Create a MS Word Report\\bin\\Debug\\plane.png";

            //    if (items[i-1] == "Picture1")
            //    {
            //        cleanBookmark("Picture1"); // Remove everything at this bookmark so we can replace it
            //        word_doc.InlineShapes.AddPicture(pictureName, ref missing, ref saveWithDocument, ref oRange);
            //    }
            //}

           



            /*

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


        //Here we remove the bookmark,text, pictures and tables and then replace the bookmark
        //back to it original place ready for you to put in the new items

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



        //Choose the template we will be using
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
