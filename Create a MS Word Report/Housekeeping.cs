using System;
using System.Drawing;
using System.Windows.Forms;

namespace Create_a_MS_Word_Report
{
    public partial class Form1
    {
        private void SetUp()
        {
            //Set tabs to see and not see
            RemoveTab(tab_header);



            //Populate things in Header Tab
            foreach (FontFamily oneFontFamily in FontFamily.Families)
            {
                cmbobx_header_fontname.Items.Add(oneFontFamily.Name);
            }

            cmbobx_header_fontname.SelectedIndex = 5;
            cmbobx_header_fontsize.SelectedIndex = 4;
            cmbobx_header_style.SelectedIndex = 2;
            cmbobx_header_colour.SelectedIndex = 0;


        }

        private void chkbx_page_header_CheckedChanged(object sender, EventArgs e)
        {
            if (chkbx_page_header.Checked)
            {
                ShowTab(1, tab_header);
            }
            else
            {
                RemoveTab(tab_header);
            }
        }








        private void RemoveTab(TabPage tab)
        {
            tabControl1.TabPages.Remove(tab);
        }

        private void ShowTab(int pos, TabPage tab)
        {
            tabControl1.TabPages.Insert(pos, tab);
        }
    }
}