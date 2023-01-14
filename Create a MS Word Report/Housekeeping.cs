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
            RemoveTab(tab_footer);


            //Populate things in Header Tab
            foreach (FontFamily oneFontFamily in FontFamily.Families)
            {
                cmbobx_header_fontname.Items.Add(oneFontFamily.Name);
                cmbobx_footer_fontname.Items.Add(oneFontFamily.Name);
            }

            cmbobx_header_fontname.SelectedIndex = 5;
            cmbobx_footer_fontname.SelectedIndex = 5;
            cmbobx_header_fontsize.SelectedIndex = 4;
            cmbobx_footer_fontsize.SelectedIndex = 4;
            cmbobx_header_font_colour.SelectedIndex = 0;
            cmbobx_footer_font_colour.SelectedIndex = 0;
            cmbobx_header_background_colour.SelectedIndex = 1;
            cmbobx_footer_background_colour.SelectedIndex = 1;
            cmbobx_header_foreground_colour.SelectedIndex = 1;
            cmbobx_footer_foreground_colour.SelectedIndex = 1;
            cmbobx_header_underline_style.SelectedIndex = 11;
            cmbobx_footer_underline_style.SelectedIndex = 11;

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

        private void chkbx_page_footer_CheckedChanged(object sender, EventArgs e)
        {
            if (chkbx_page_footer.Checked)
            {
                ShowTab(2, tab_footer);
            }
            else
            {
                RemoveTab(tab_footer);
            }
        }






        private void RemoveTab(TabPage tab)
        {
            TabControl1.TabPages.Remove(tab);
        }

        private void ShowTab(int pos, TabPage tab)
        {
           TabControl1.TabPages.Insert(TabControl1.TabPages.Count, tab);
        }
    }
}