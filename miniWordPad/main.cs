using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace miniWordPad
{
    public partial class frmWordPad : Form
    {
        public frmWordPad()
        {
            InitializeComponent();
            tsmiFontStyle.Visible = false;
            tsBtnFontColor.Font = new Font(tsBtnFontColor.Font, FontStyle.Bold);
            tsBtnHightLight.Font = new Font(tsBtnHightLight.Font, FontStyle.Bold);
            foreach (FontFamily Font in FontFamily.Families)
            {
                tsCbFontStyle.Items.Add(Font.Name.ToString());
            }
            tsCbFontStyle.Text = rtxtWordPad.SelectionFont.Name.ToString();
            tsCbFontSize.Text = rtxtWordPad.SelectionFont.Size.ToString();
            tsBtnFontColor.ForeColor = rtxtWordPad.SelectionColor;
        }

        private bool isSaved = true;

        private void rtxtWordPad_TextChanged(object sender, EventArgs e) => isSaved = false;

        private void frmWordPad_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (isSaved == false)
            {
                MessageBoxManager.Yes = "Thoát";
                MessageBoxManager.No = "Lưu lại";
                MessageBoxManager.Cancel = "Trở lại";
                MessageBoxManager.Register();
                DialogResult dR = MessageBox.Show("Dữ liệu chưa được lưu sẽ bị mất\r\nBạn có chắc muốn thoát?", "Cảnh báo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (dR == DialogResult.Cancel)
                    e.Cancel = true;
                else if (dR == DialogResult.No)
                {
                    this.saveFile(false, "Save Document Files");
                    e.Cancel = true;
                }
                MessageBoxManager.Unregister();
            }
        }

        private void rtxtWordPad_KeyDown(object sender, KeyEventArgs e)
        {
            RichTextBox rtb = (RichTextBox)sender;
            if (e.KeyCode == Keys.Space || e.KeyCode == Keys.Tab)
            {
                this.SuspendLayout();
                rtb.Undo();
                rtb.Redo();
                this.ResumeLayout();
            }
        }

        //File menu
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isSaved == false)
            {
                MessageBoxManager.Yes = "Thoát";
                MessageBoxManager.No = "Lưu lại";
                MessageBoxManager.Cancel = "Trở lại";
                MessageBoxManager.Register();
                DialogResult dR = MessageBox.Show("Dữ liệu chưa được lưu sẽ bị mất\r\nBạn có chắc muốn thoát?", "Cảnh báo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (dR == DialogResult.Yes)
                    this.Close();
                else if (dR == DialogResult.No)
                {
                    this.saveFile(false, "Save Document Files");
                }
                MessageBoxManager.Unregister();
            }
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isSaved == false)
            {
                MessageBoxManager.Yes = "Tiếp tục";
                MessageBoxManager.No = "Lưu lại";
                MessageBoxManager.Cancel = "Hủy bỏ";
                MessageBoxManager.Register();
                DialogResult dR = MessageBox.Show("Dữ liệu chưa được lưu sẽ bị mất\r\nBạn có muốn tiếp tục?", "Cảnh báo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                
                if (dR == DialogResult.Yes)
                    rtxtWordPad.Clear();
                else if (dR == DialogResult.No)
                    this.saveFile(false, "Save Document Files");
                MessageBoxManager.Unregister();
            }
        }

        private void saveFile(bool check, string title)
        {
            saveFDWordPad.InitialDirectory = @"C:\";
            saveFDWordPad.Title = title;
            saveFDWordPad.CheckFileExists = check;
            saveFDWordPad.CheckPathExists = true;
            saveFDWordPad.DefaultExt = "*.doc";
            saveFDWordPad.Filter = "Document Files(*.doc)|*.doc|All files (*.*)|*.*";
            saveFDWordPad.FilterIndex = 2;
            saveFDWordPad.RestoreDirectory = true;
            if (saveFDWordPad.ShowDialog() == DialogResult.OK && saveFDWordPad.FileName.Length > 0)
            {
                rtxtWordPad.SaveFile(saveFDWordPad.FileName, RichTextBoxStreamType.RichText);
                isSaved = true;
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFile(false, "Save Document Files");
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFile(true, "Save Document Files As");
        }
       
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFDWordPad.InitialDirectory = @"C:\";
            openFDWordPad.Title = "Open Document Files";
            openFDWordPad.CheckFileExists = true;
            openFDWordPad.CheckPathExists = true;
            openFDWordPad.DefaultExt = "*.doc";
            openFDWordPad.Filter = "Document Files(*.doc)|*.doc|All files (*.*)|*.*";
            openFDWordPad.FilterIndex = 2;
            if (openFDWordPad.ShowDialog() == DialogResult.OK && openFDWordPad.FileName.Length > 0)
            {
                rtxtWordPad.LoadFile(openFDWordPad.FileName, RichTextBoxStreamType.RichText);
                isSaved = true;
            }
        }

        //fontStyle Menu
        private void boldToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Font newFont;
            if (rtxtWordPad.SelectionFont.Style != (FontStyle.Bold | rtxtWordPad.SelectionFont.Style))
                newFont = new Font(rtxtWordPad.SelectionFont.FontFamily.Name, rtxtWordPad.SelectionFont.Size, rtxtWordPad.SelectionFont.Style | FontStyle.Bold);
            else
                newFont = new Font(rtxtWordPad.SelectionFont.FontFamily.Name, rtxtWordPad.SelectionFont.Size, rtxtWordPad.SelectionFont.Style & ~FontStyle.Bold);
            rtxtWordPad.SelectionFont = newFont;
        }

        private void italicToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Font newFont;
            if (rtxtWordPad.SelectionFont.Style != (FontStyle.Italic | rtxtWordPad.SelectionFont.Style))
                newFont = new Font(rtxtWordPad.SelectionFont.FontFamily.Name, rtxtWordPad.SelectionFont.Size, rtxtWordPad.SelectionFont.Style | FontStyle.Italic);
            else
                newFont = new Font(rtxtWordPad.SelectionFont.FontFamily.Name, rtxtWordPad.SelectionFont.Size, rtxtWordPad.SelectionFont.Style & ~FontStyle.Italic);
            rtxtWordPad.SelectionFont = newFont;
        }

        private void underlineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Font newFont;
            if (rtxtWordPad.SelectionFont.Style != (FontStyle.Underline | rtxtWordPad.SelectionFont.Style))
                newFont = new Font(rtxtWordPad.SelectionFont.FontFamily.Name, rtxtWordPad.SelectionFont.Size, rtxtWordPad.SelectionFont.Style | FontStyle.Underline);
            else
                newFont = new Font(rtxtWordPad.SelectionFont.FontFamily.Name, rtxtWordPad.SelectionFont.Size, rtxtWordPad.SelectionFont.Style & ~FontStyle.Underline);
            rtxtWordPad.SelectionFont = newFont;
        }

        private void leftToolStripMenuItem_Click(object sender, EventArgs e)
        {
                rtxtWordPad.SelectionAlignment = HorizontalAlignment.Left;
        }

        private void centerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtxtWordPad.SelectionAlignment = HorizontalAlignment.Center;
        }

        private void rightToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtxtWordPad.SelectionAlignment = HorizontalAlignment.Right;
        }

        private void tsCbFontStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            rtxtWordPad.SelectionFont = new Font(tsCbFontStyle.Text, rtxtWordPad.SelectionFont.Size);
        }

        private void tsCbFontSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            rtxtWordPad.SelectionFont = new Font(rtxtWordPad.SelectionFont.FontFamily, float.Parse(tsCbFontSize.Text));
        }


        private void changeFont()
        {
            if (tsCbFontSize.Text.Length < 1)
            {
                tsCbFontSize.Text = "10";
            }
            rtxtWordPad.SelectionFont = new Font(rtxtWordPad.SelectionFont.FontFamily, float.Parse(tsCbFontSize.Text));
        }

        private void tsCbFontSize_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyValue < 48 || e.KeyValue > 57) && e.KeyCode != Keys.Back && e.KeyCode != Keys.Delete)
            {
                e.SuppressKeyPress = true;
            }
            if (e.KeyCode == Keys.Enter)
            {
                changeFont();
                rtxtWordPad.Focus();
            }
        }

        private void tsCbFontSize_Leave(object sender, EventArgs e)
        {
            changeFont();
        }

        private void tsBtnFontColor_Click(object sender, EventArgs e)
        {
            ColorDialog colorDiag = new ColorDialog();
            if (colorDiag.ShowDialog() == DialogResult.OK)
            {
                rtxtWordPad.SelectionColor = colorDiag.Color;
            }
        }

        private void tsBtnHightlight_Click(object sender, EventArgs e)
        {
            ColorDialog colorDiag = new ColorDialog();
            if (colorDiag.ShowDialog() == DialogResult.OK)
            {
                rtxtWordPad.SelectionBackColor = colorDiag.Color;
            }
        }

        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FontDialog fontDiag = new FontDialog();
            if (fontDiag.ShowDialog() == DialogResult.OK)
                if (rtxtWordPad.SelectedText != "")
                    rtxtWordPad.Font = fontDiag.Font;
        }

        //About menu
        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Trường Đại học Sư phạm Thành phố Hồ Chí Minh\r\nKhoa Công Nghệ Thông Tin\r\nTrần Vĩnh Phước\r\n43.01.104.136", "Thông tin", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //Undo & Redo
        private void tsBtnUndo_Click(object sender, EventArgs e)
        {
            rtxtWordPad.Undo();
        }

        private void tsBtnRedo_Click(object sender, EventArgs e)
        {
            rtxtWordPad.Redo();
        }

        //Edit menu
        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e) => rtxtWordPad.SelectAll();

        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmFindReplace frm = new frmFindReplace(rtxtWordPad, this);
            frm.Show();
        }
    }
}