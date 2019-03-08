using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace miniWordPad
{
    public partial class frmFindReplace : Form
    {
        private RichTextBox rtxt = new RichTextBox();
        private frmWordPad frm = new frmWordPad();
        public frmFindReplace(RichTextBox richText, frmWordPad frmW)
        {
            InitializeComponent();
            rtxt = richText;
            frm = frmW;
        }

        private int Find(RichTextBox richText, string text, bool matchCase, bool matchWholeWord, bool upDirection)
        {
            RichTextBoxFinds options = new RichTextBoxFinds();
            int index = new int();
            if (matchCase)
                options |= RichTextBoxFinds.MatchCase;
            if (matchWholeWord)
                options |= RichTextBoxFinds.WholeWord;
            if (upDirection)
                options |= RichTextBoxFinds.Reverse;
            if (upDirection)
                index = richText.Find(text, 0, richText.SelectionStart, options);
            else
                index = richText.Find(text, richText.SelectionStart + richText.SelectionLength, options);
            if (index >= 0)
            {
                richText.SelectionStart = index;
                richText.SelectionLength = text.Length;
                frm.Focus();
            }
            else // text not found
            {
                MessageBox.Show("Mini WordPad không tìm thấy chuỗi", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return index;
        }

        private void btnFindNext_Click(object sender, EventArgs e)
        {
            if (txtFind.Text == "")
                MessageBox.Show("Nhập từ cần tìm", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
                Find(rtxt, txtFind.Text, checkCase.Checked, checkWholeWord.Checked, rBtnUp.Checked);
        }

        private void btnReplace_Click(object sender, EventArgs e)
        {
            if (txtFind.Text == "")
                MessageBox.Show("Nhập từ cần tìm", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            if (txtReplace.Text == "")
                MessageBox.Show("Nhập từ bạn muốn thay thành", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            if (rtxt.SelectedText != "")
                rtxt.SelectedText = txtReplace.Text;
        }

        private void btnRelaceAll_Click(object sender, EventArgs e)
        {
            if (txtFind.Text == "")
                MessageBox.Show("Nhập từ cần tìm", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            if (txtReplace.Text == "")
                MessageBox.Show("Nhập từ bạn muốn thay thành", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            if (Find(rtxt, txtFind.Text, checkCase.Checked, checkWholeWord.Checked, rBtnUp.Checked) != 0)
                rtxt.Text = rtxt.Text.Replace(txtFind.Text, txtReplace.Text);
        }
    }
}
