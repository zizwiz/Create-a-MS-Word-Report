using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
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

        Word.Application WinWord = new Word.Application(); // open a word app in windows.
        Word.Document word_doc = new Word.Document();      // open word doc in app.


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SetUp();

            Text += " : v" + Assembly.GetExecutingAssembly().GetName().Version; // put in the version number

            btn_create.Visible = false; //only allow button when we have an open doc.
        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btn_create_Click(object sender, EventArgs e)
        {
            string myType = "";
            object oMissing = Missing.Value;

            //This uses the styles from your default-default word doc template
            //change them below if you have your own template
            WinWord.Visible = false; //needs to be here to stop RPC server issues. No idea why.
            WinWord.Visible = true; //Set status for word application is to be visible or not.
            WinWord.ShowAnimation = false; //Set animation status for word application

            List<Word.Bookmark> bmarks = FindBookmarks(2); //Get a list of bookmarks in the document


            if (chkbx_page_header.Checked) //Only create if we ticked to say so
            {
                CreateHeader(word_doc);
            }

            if (chkbx_page_footer.Checked)
            {
                CreateFooter(word_doc);
            }

            foreach (Word.Bookmark b in bmarks)
            {
                //Text is prefixed txt_ and images by img_
                myType = b.Name.Split('_')[0];

                if (myType == "txt")
                {
                    cleanBookmark(b.Name); // Remove everything at this bookmark so we can replace it
                    ReplaceBookmarkText(word_doc, b.Name, ((TextBox)tab_bookmark_update.Controls["txtbx_" + b.Name]).Text);
                }
                else if (myType == "img")
                {
                    //change a picture
                    string bookmarkname = b.Name;
                    string pictureName = ((TextBox)tab_bookmark_update.Controls["txtbx_" + bookmarkname]).Text;

                    //change a picture at this bookmarkname for the picture named one.
                    ChangePicture(bmarks, pictureName, bookmarkname);
                }
                else
                {
                    cleanBookmark(b.Name);
                }
            }


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



        private void ChangePicture(List<Word.Bookmark> bmarks, string pictureName, string bookmarkname)
        {
            //change a picture
            int range_count = 1;

            foreach (Word.Bookmark b in bmarks)
            {
                object oRange = word_doc.Bookmarks[range_count].Range;
                object saveWithDocument = true;
                object missing = Type.Missing;

                if (b.Name == bookmarkname)
                {
                    cleanBookmark(bookmarkname); // Remove everything at this bookmark so we can replace it
                    word_doc.InlineShapes.AddPicture(pictureName, ref missing, ref saveWithDocument, ref oRange);
                }

                range_count++;
            }
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

                // Create the Word document, choose your template here or you will get the default-default one.
                // Add(Template, New Template, DocType, Visible).
                word_doc = WinWord.Documents.Add(
                    ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                btn_create.Visible = true;
            }

            List<Word.Bookmark> myBookmarks = FindBookmarks(1); //we ignore return value
        }

        private List<Word.Bookmark> FindBookmarks(int type)
        {
            //type 1 = get list and write to GUI
            //type 2 = get list only

            // find all the bookmarks in the doc and list them out.
            int bkmk_count = 0;
            int bkmk_num = 0;

            List<Word.Bookmark> bmarks = new List<Word.Bookmark>();
            foreach (Word.Bookmark bookmark in word_doc.Bookmarks)
            {
                bmarks.Add(bookmark); //this adds in alphabetical order
            }

            if (type == 1) //write to GUI 
            {
                // Re-sort list in order of appearance
                bmarks = bmarks.OrderBy(b => b.Start).ToList(); // LINQ

                //Add the bookmark names to the GUI in labels
                List<string> slBMarks = new List<string>();
                foreach (Word.Bookmark b in bmarks)
                {
                    //add a label to the screen
                    Label label = new Label();
                    label.Name = "lbl_" + b.Name;
                    label.Location = new Point(40, 100 + bkmk_count);
                    label.Text = b.Name;
                    tab_bookmark_update.Controls.Add(label);

                    //add a checkbox to the screen
                    //we check it later if it is an image to be replaced.
                    CheckBox checkBox = new CheckBox();
                    checkBox.Name = "chkbx_" + b.Name;
                    checkBox.Location = new Point(10, 100 + bkmk_count);
                    //get part before the underscore
                    if ((b.Name.Split('_')[0]) == "img") checkBox.Checked = true;
                    tab_bookmark_update.Controls.Add(checkBox);

                    //add a textbox to the screen
                    TextBox textBox = new TextBox();
                    textBox.Name = "txtbx_" + b.Name;
                    textBox.Location = new Point(150, 100 + bkmk_count);
                    textBox.Width = 430;
                    textBox.Text = label.Name + " : " + textBox.Name;
                    tab_bookmark_update.Controls.Add(textBox);

                    //add a button to the screen
                    Button button = new Button();
                    button.Name = "btn_" + b.Name;
                    button.Text = "Get Data";
                    button.Location = new Point(600, 100 + bkmk_count);
                    button.Size = new Size(100, 20);
                    tab_bookmark_update.Controls.Add(button);

                    // add click event to the button.
                    button.Click += new EventHandler(MyButton_Click);

                    bkmk_count += 40;
                    bkmk_num++;

                    slBMarks.Add(b.Name); // Accumulate bookmark names
                }
            }

            return bmarks;
        }

        // Event for all the bookmark buttons so we can get data from them.
        private void MyButton_Click(object sender, EventArgs e)
        {
            string myType = "";
            string myName = "";
            
            //Get the name of the button that wa just pressed
            Button btn = (Button)sender;
            string senderBtnType = btn.Name.Split('_')[1];

            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                
                Title = "Choose File.",
                Filter = "Image files (*.png)|*.png|Text files (*.txt)|*.txt|All files (*.*)|*.*",
                    
                CheckFileExists = true,
                CheckPathExists = true,
                FilterIndex = 1,
                RestoreDirectory = true,

            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //get a list of the controls and if a textbox check its name
                //if the name is same as the one we are looking for then we add the file name
                //to the text box
                foreach (Control c in tab_bookmark_update.Controls)
                {
                    //Make sure we have the correct items and if so write the file name in it.
                    myType = btn.Name.Split('_')[1] + "_" + btn.Name.Split('_')[2];
                    myName = c.Name.Split('_')[1] + "_" + c.Name.Split('_')[2];

                    if ((c is TextBox) && (myType == myName) && (senderBtnType == "img"))
                    {
                        try
                        {
                            c.Text = openFileDialog1.FileName;
                            myType = "";
                            break;
                        }
                        catch (Exception exception)
                        {
                            MessageBox.Show("not a image file");
                        }

                    }
                    else if ((c is TextBox) && (myType == myName) && (senderBtnType == "txt"))
                    {
                        try
                        {
                            c.Text = File.ReadAllText(openFileDialog1.FileName);
                            myType = "";
                            break;
                        }
                        catch (Exception exception)
                        {
                            MessageBox.Show("not a text file");
                        }
                    }

                }

            }

        }
    }
}
