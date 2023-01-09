using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word; // now got to ref and in properties set "Embed Interop types" to false

namespace Create_a_MS_Word_Report
{
    public partial class Form1 : Form
    {
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
            //This uses the styles from your default-default word doc template
            //change them below if you have your own template

            Word._Application wordDoc = new Word.ApplicationClass(); // create a word object and show it.
            wordDoc.Visible = true; //Set status for word application is to be visible or not.
            wordDoc.ShowAnimation = false; //Set animation status for word application

            // Create the Word document, choose your template here or you will get the default-default one.
            // Add(Template, New Template, DocType, Visible).
            Word._Document word_doc = wordDoc.Documents.Add(
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Word.Paragraph para = word_doc.Paragraphs.Add(Type.Missing);

            if (chkbx_page_header.Checked) //Only create if we ticked to say so
            {
                CreateHeader(word_doc);
            }

            if (chkbx_page_footer.Checked)
            {
                //Add the footers into the document  
                foreach (Word.Section wordSection in word_doc.Sections)
                {
                    //Get the footer range and add the footer details.  
                    Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkRed;
                    footerRange.Font.Size = 10;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    footerRange.Text = "Footer text goes here";
                }
            }




            // Report Title.
            para.Range.Text = "Project: " + txtbx_proj_name.Text + " Report";
            para.Range.set_Style("Title");
            para.Range.InsertParagraphAfter();

            // Create a header.
            para.Range.set_Style("Heading 1");
            para.Range.Text = "Income";
            para.Range.InsertParagraphAfter();

            //add text to paragraph not I am also using CR and LF
            para.Range.set_Style("Normal");
            para.Range.Text = "Loads and Loads of text\rMore\rmuch more\r\r";

            string picture_file =
                @"C:\\Users\\itobo\\source\repos\\Create-a-MS-Word-Report\\Create a MS Word Report\\bin\\Debug\\mypic.png";

            // Add the picture to the paragraph.
            Word.InlineShape inline_shape = para.Range.InlineShapes.AddPicture(
                picture_file, Type.Missing, Type.Missing, Type.Missing);


        }

        
    }
}
