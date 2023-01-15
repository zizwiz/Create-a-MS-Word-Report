

using System;

namespace Create_a_MS_Word_Report
{
    public partial class Form1
    {
        /// <summary>
        /// Call the common method ReplaceBookmarkText to replace bookmark StudentName with actual value
        /// i.e studentName
        /// 
        /// Usage = ReplaceBookmarkText(doc, "StudentName", studentName);
        /// 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="bookmarkName"></param>
        /// <param name="text"></param>
        private void ReplaceBookmarkText(Microsoft.Office.Interop.Word.Document doc, string bookmarkName, string text)
        {
            if (doc.Bookmarks.Exists(bookmarkName))
            {
                Object name = bookmarkName;
                Microsoft.Office.Interop.Word.Range range = doc.Bookmarks.get_Item(ref name).Range;

                range.Text = text; //replaces text
                object newRange = range;

                doc.Bookmarks.Add(bookmarkName, ref newRange);
            }
        }


	}
}