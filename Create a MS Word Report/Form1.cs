using System;
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

            // create a word object and show it.
            Word._Application word_app = new Word.ApplicationClass();
            word_app.Visible = true;

            // Create the Word document, choose your template here or you will get the default-default one.
            // Add(Template, New Template, DocType, Visible).
            Word._Document word_doc = word_app.Documents.Add(
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Report Header.
            Word.Paragraph para = word_doc.Paragraphs.Add(Type.Missing);
            para.Range.Text = "Project: " +txtbx_proj_name.Text + " Report";
            para.Range.set_Style("Title");
            para.Range.InsertParagraphAfter();

            // Create a header.
            para.Range.set_Style("Heading 1");
            para.Range.Text = "Income";
            para.Range.InsertParagraphAfter();

            //add text to paragraph not I am also using CR and LF
            para.Range.set_Style("Normal");
            para.Range.Text = "Loads and Loads of text\rMore\rmuch more\r\r";


            



        }
    }
}
