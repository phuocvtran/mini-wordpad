namespace miniWordPad
{
    partial class frmWordPad
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmWordPad));
            this.menuStripWordPad = new System.Windows.Forms.MenuStrip();
            this.tsmiFile = new System.Windows.Forms.ToolStripMenuItem();
            this.newToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiFormat = new System.Windows.Forms.ToolStripMenuItem();
            this.fontToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiEdit = new System.Windows.Forms.ToolStripMenuItem();
            this.findToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            this.selectAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiAbout = new System.Windows.Forms.ToolStripMenuItem();
            this.infoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiFontStyle = new System.Windows.Forms.ToolStripMenuItem();
            this.boldToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.italicToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.underlineToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.leftToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.centerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.rightToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.justifyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.rtxtWordPad = new System.Windows.Forms.RichTextBox();
            this.tsWordPadBottom = new System.Windows.Forms.ToolStrip();
            this.tsDropDownBtn = new System.Windows.Forms.ToolStripDropDownButton();
            this.newToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.newToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.tsBtnSave = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.tsBtnUndo = new System.Windows.Forms.ToolStripButton();
            this.tsBtnRedo = new System.Windows.Forms.ToolStripButton();
            this.ttWordPad = new System.Windows.Forms.ToolTip(this.components);
            this.tsWordPadTop = new System.Windows.Forms.ToolStrip();
            this.tsCbFontStyle = new System.Windows.Forms.ToolStripComboBox();
            this.tsCbFontSize = new System.Windows.Forms.ToolStripComboBox();
            this.tsBtnFontColor = new System.Windows.Forms.ToolStripButton();
            this.tsBtnHightLight = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.tsBtnBold = new System.Windows.Forms.ToolStripButton();
            this.tsBtnItalic = new System.Windows.Forms.ToolStripButton();
            this.tsBtnUnderline = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.tsBtnLeft = new System.Windows.Forms.ToolStripButton();
            this.tsBtnCenter = new System.Windows.Forms.ToolStripButton();
            this.tsBtnRight = new System.Windows.Forms.ToolStripButton();
            this.saveFDWordPad = new System.Windows.Forms.SaveFileDialog();
            this.openFDWordPad = new System.Windows.Forms.OpenFileDialog();
            this.menuStripWordPad.SuspendLayout();
            this.tsWordPadBottom.SuspendLayout();
            this.tsWordPadTop.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStripWordPad
            // 
            this.menuStripWordPad.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.menuStripWordPad.Dock = System.Windows.Forms.DockStyle.None;
            this.menuStripWordPad.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiFile,
            this.tsmiFormat,
            this.tsmiEdit,
            this.tsmiAbout,
            this.tsmiFontStyle});
            this.menuStripWordPad.Location = new System.Drawing.Point(0, 0);
            this.menuStripWordPad.Name = "menuStripWordPad";
            this.menuStripWordPad.Size = new System.Drawing.Size(259, 24);
            this.menuStripWordPad.TabIndex = 0;
            this.menuStripWordPad.Text = "menuStrip1";
            // 
            // tsmiFile
            // 
            this.tsmiFile.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newToolStripMenuItem,
            this.openToolStripMenuItem,
            this.saveToolStripMenuItem,
            this.saveAsToolStripMenuItem,
            this.toolStripMenuItem1,
            this.exitToolStripMenuItem});
            this.tsmiFile.Name = "tsmiFile";
            this.tsmiFile.Size = new System.Drawing.Size(37, 20);
            this.tsmiFile.Text = "File";
            // 
            // newToolStripMenuItem
            // 
            this.newToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("newToolStripMenuItem.Image")));
            this.newToolStripMenuItem.Name = "newToolStripMenuItem";
            this.newToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N)));
            this.newToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.newToolStripMenuItem.Text = "New";
            this.newToolStripMenuItem.Click += new System.EventHandler(this.newToolStripMenuItem_Click);
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("openToolStripMenuItem.Image")));
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.openToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.openToolStripMenuItem.Text = "Open";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("saveToolStripMenuItem.Image")));
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.saveToolStripMenuItem.Text = "Save";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
            // 
            // saveAsToolStripMenuItem
            // 
            this.saveAsToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("saveAsToolStripMenuItem.Image")));
            this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            this.saveAsToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.S)));
            this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.saveAsToolStripMenuItem.Text = "Save as";
            this.saveAsToolStripMenuItem.Click += new System.EventHandler(this.saveAsToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(181, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.X)));
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // tsmiFormat
            // 
            this.tsmiFormat.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fontToolStripMenuItem});
            this.tsmiFormat.Name = "tsmiFormat";
            this.tsmiFormat.Size = new System.Drawing.Size(57, 20);
            this.tsmiFormat.Text = "Format";
            // 
            // fontToolStripMenuItem
            // 
            this.fontToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("fontToolStripMenuItem.Image")));
            this.fontToolStripMenuItem.Name = "fontToolStripMenuItem";
            this.fontToolStripMenuItem.Size = new System.Drawing.Size(98, 22);
            this.fontToolStripMenuItem.Text = "Font";
            this.fontToolStripMenuItem.Click += new System.EventHandler(this.fontToolStripMenuItem_Click);
            // 
            // tsmiEdit
            // 
            this.tsmiEdit.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.findToolStripMenuItem,
            this.toolStripMenuItem2,
            this.selectAllToolStripMenuItem});
            this.tsmiEdit.Name = "tsmiEdit";
            this.tsmiEdit.Size = new System.Drawing.Size(39, 20);
            this.tsmiEdit.Text = "Edit";
            // 
            // findToolStripMenuItem
            // 
            this.findToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("findToolStripMenuItem.Image")));
            this.findToolStripMenuItem.Name = "findToolStripMenuItem";
            this.findToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F)));
            this.findToolStripMenuItem.Size = new System.Drawing.Size(204, 22);
            this.findToolStripMenuItem.Text = "Find and Relpace";
            this.findToolStripMenuItem.Click += new System.EventHandler(this.findToolStripMenuItem_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(201, 6);
            // 
            // selectAllToolStripMenuItem
            // 
            this.selectAllToolStripMenuItem.Name = "selectAllToolStripMenuItem";
            this.selectAllToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.A)));
            this.selectAllToolStripMenuItem.Size = new System.Drawing.Size(204, 22);
            this.selectAllToolStripMenuItem.Text = "Select all";
            this.selectAllToolStripMenuItem.Click += new System.EventHandler(this.selectAllToolStripMenuItem_Click);
            // 
            // tsmiAbout
            // 
            this.tsmiAbout.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.infoToolStripMenuItem});
            this.tsmiAbout.Name = "tsmiAbout";
            this.tsmiAbout.Size = new System.Drawing.Size(52, 20);
            this.tsmiAbout.Text = "About";
            // 
            // infoToolStripMenuItem
            // 
            this.infoToolStripMenuItem.Name = "infoToolStripMenuItem";
            this.infoToolStripMenuItem.Size = new System.Drawing.Size(95, 22);
            this.infoToolStripMenuItem.Text = "Info";
            this.infoToolStripMenuItem.Click += new System.EventHandler(this.infoToolStripMenuItem_Click);
            // 
            // tsmiFontStyle
            // 
            this.tsmiFontStyle.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.boldToolStripMenuItem,
            this.italicToolStripMenuItem,
            this.underlineToolStripMenuItem,
            this.leftToolStripMenuItem,
            this.centerToolStripMenuItem,
            this.rightToolStripMenuItem,
            this.justifyToolStripMenuItem});
            this.tsmiFontStyle.Name = "tsmiFontStyle";
            this.tsmiFontStyle.Size = new System.Drawing.Size(66, 20);
            this.tsmiFontStyle.Text = "fontStyle";
            // 
            // boldToolStripMenuItem
            // 
            this.boldToolStripMenuItem.Name = "boldToolStripMenuItem";
            this.boldToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.B)));
            this.boldToolStripMenuItem.Size = new System.Drawing.Size(167, 22);
            this.boldToolStripMenuItem.Text = "Bold";
            this.boldToolStripMenuItem.Click += new System.EventHandler(this.boldToolStripMenuItem_Click);
            // 
            // italicToolStripMenuItem
            // 
            this.italicToolStripMenuItem.Name = "italicToolStripMenuItem";
            this.italicToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.I)));
            this.italicToolStripMenuItem.Size = new System.Drawing.Size(167, 22);
            this.italicToolStripMenuItem.Text = "Italic";
            this.italicToolStripMenuItem.Click += new System.EventHandler(this.italicToolStripMenuItem_Click);
            // 
            // underlineToolStripMenuItem
            // 
            this.underlineToolStripMenuItem.Name = "underlineToolStripMenuItem";
            this.underlineToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.U)));
            this.underlineToolStripMenuItem.Size = new System.Drawing.Size(167, 22);
            this.underlineToolStripMenuItem.Text = "Underline";
            this.underlineToolStripMenuItem.Click += new System.EventHandler(this.underlineToolStripMenuItem_Click);
            // 
            // leftToolStripMenuItem
            // 
            this.leftToolStripMenuItem.Name = "leftToolStripMenuItem";
            this.leftToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.L)));
            this.leftToolStripMenuItem.Size = new System.Drawing.Size(167, 22);
            this.leftToolStripMenuItem.Text = "Left";
            this.leftToolStripMenuItem.Click += new System.EventHandler(this.leftToolStripMenuItem_Click);
            // 
            // centerToolStripMenuItem
            // 
            this.centerToolStripMenuItem.Name = "centerToolStripMenuItem";
            this.centerToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E)));
            this.centerToolStripMenuItem.Size = new System.Drawing.Size(167, 22);
            this.centerToolStripMenuItem.Text = "Center";
            this.centerToolStripMenuItem.Click += new System.EventHandler(this.centerToolStripMenuItem_Click);
            // 
            // rightToolStripMenuItem
            // 
            this.rightToolStripMenuItem.Name = "rightToolStripMenuItem";
            this.rightToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.R)));
            this.rightToolStripMenuItem.Size = new System.Drawing.Size(167, 22);
            this.rightToolStripMenuItem.Text = "Right";
            this.rightToolStripMenuItem.Click += new System.EventHandler(this.rightToolStripMenuItem_Click);
            // 
            // justifyToolStripMenuItem
            // 
            this.justifyToolStripMenuItem.Name = "justifyToolStripMenuItem";
            this.justifyToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.J)));
            this.justifyToolStripMenuItem.Size = new System.Drawing.Size(167, 22);
            this.justifyToolStripMenuItem.Text = "Justify";
            // 
            // rtxtWordPad
            // 
            this.rtxtWordPad.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rtxtWordPad.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rtxtWordPad.Location = new System.Drawing.Point(12, 61);
            this.rtxtWordPad.Name = "rtxtWordPad";
            this.rtxtWordPad.Size = new System.Drawing.Size(773, 287);
            this.rtxtWordPad.TabIndex = 2;
            this.rtxtWordPad.TabStop = false;
            this.rtxtWordPad.Text = "";
            this.rtxtWordPad.TextChanged += new System.EventHandler(this.rtxtWordPad_TextChanged);
            this.rtxtWordPad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.rtxtWordPad_KeyDown);
            // 
            // tsWordPadBottom
            // 
            this.tsWordPadBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.tsWordPadBottom.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsDropDownBtn,
            this.toolStripSeparator1,
            this.toolStripSeparator2,
            this.tsBtnSave,
            this.toolStripSeparator3,
            this.tsBtnUndo,
            this.tsBtnRedo});
            this.tsWordPadBottom.Location = new System.Drawing.Point(0, 351);
            this.tsWordPadBottom.Name = "tsWordPadBottom";
            this.tsWordPadBottom.Size = new System.Drawing.Size(797, 25);
            this.tsWordPadBottom.TabIndex = 3;
            this.tsWordPadBottom.Text = "toolStripWordPadBottom";
            // 
            // tsDropDownBtn
            // 
            this.tsDropDownBtn.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsDropDownBtn.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newToolStripMenuItem2,
            this.openToolStripMenuItem1,
            this.saveAsToolStripMenuItem2,
            this.saveAsToolStripMenuItem1,
            this.newToolStripMenuItem1});
            this.tsDropDownBtn.Image = ((System.Drawing.Image)(resources.GetObject("tsDropDownBtn.Image")));
            this.tsDropDownBtn.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsDropDownBtn.Name = "tsDropDownBtn";
            this.tsDropDownBtn.Size = new System.Drawing.Size(29, 22);
            this.tsDropDownBtn.Text = "toolStripDropDownButton1";
            this.tsDropDownBtn.ToolTipText = "Drop menu";
            // 
            // newToolStripMenuItem2
            // 
            this.newToolStripMenuItem2.Name = "newToolStripMenuItem2";
            this.newToolStripMenuItem2.Size = new System.Drawing.Size(112, 22);
            this.newToolStripMenuItem2.Text = "New";
            this.newToolStripMenuItem2.Click += new System.EventHandler(this.newToolStripMenuItem_Click);
            // 
            // openToolStripMenuItem1
            // 
            this.openToolStripMenuItem1.Name = "openToolStripMenuItem1";
            this.openToolStripMenuItem1.Size = new System.Drawing.Size(112, 22);
            this.openToolStripMenuItem1.Text = "Open";
            this.openToolStripMenuItem1.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // saveAsToolStripMenuItem2
            // 
            this.saveAsToolStripMenuItem2.Name = "saveAsToolStripMenuItem2";
            this.saveAsToolStripMenuItem2.Size = new System.Drawing.Size(112, 22);
            this.saveAsToolStripMenuItem2.Text = "Save as";
            this.saveAsToolStripMenuItem2.Click += new System.EventHandler(this.saveAsToolStripMenuItem_Click);
            // 
            // saveAsToolStripMenuItem1
            // 
            this.saveAsToolStripMenuItem1.Name = "saveAsToolStripMenuItem1";
            this.saveAsToolStripMenuItem1.Size = new System.Drawing.Size(109, 6);
            // 
            // newToolStripMenuItem1
            // 
            this.newToolStripMenuItem1.Name = "newToolStripMenuItem1";
            this.newToolStripMenuItem1.Size = new System.Drawing.Size(112, 22);
            this.newToolStripMenuItem1.Text = "Exit";
            this.newToolStripMenuItem1.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // tsBtnSave
            // 
            this.tsBtnSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsBtnSave.Image = ((System.Drawing.Image)(resources.GetObject("tsBtnSave.Image")));
            this.tsBtnSave.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsBtnSave.Name = "tsBtnSave";
            this.tsBtnSave.Size = new System.Drawing.Size(23, 22);
            this.tsBtnSave.Text = "toolStripButton1";
            this.tsBtnSave.ToolTipText = "Save\r\nCtrl + S";
            this.tsBtnSave.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 25);
            // 
            // tsBtnUndo
            // 
            this.tsBtnUndo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsBtnUndo.Image = ((System.Drawing.Image)(resources.GetObject("tsBtnUndo.Image")));
            this.tsBtnUndo.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsBtnUndo.Name = "tsBtnUndo";
            this.tsBtnUndo.Size = new System.Drawing.Size(23, 22);
            this.tsBtnUndo.Text = "toolStripButton2";
            this.tsBtnUndo.ToolTipText = "Undo\r\nCtrl + Z";
            this.tsBtnUndo.Click += new System.EventHandler(this.tsBtnUndo_Click);
            // 
            // tsBtnRedo
            // 
            this.tsBtnRedo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsBtnRedo.Image = ((System.Drawing.Image)(resources.GetObject("tsBtnRedo.Image")));
            this.tsBtnRedo.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsBtnRedo.Name = "tsBtnRedo";
            this.tsBtnRedo.Size = new System.Drawing.Size(23, 22);
            this.tsBtnRedo.Text = "toolStripButton3";
            this.tsBtnRedo.ToolTipText = "Redo\r\nCtrl + Y";
            this.tsBtnRedo.Click += new System.EventHandler(this.tsBtnRedo_Click);
            // 
            // tsWordPadTop
            // 
            this.tsWordPadTop.Dock = System.Windows.Forms.DockStyle.None;
            this.tsWordPadTop.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsCbFontStyle,
            this.tsCbFontSize,
            this.tsBtnFontColor,
            this.tsBtnHightLight,
            this.toolStripSeparator5,
            this.tsBtnBold,
            this.tsBtnItalic,
            this.tsBtnUnderline,
            this.toolStripSeparator4,
            this.tsBtnLeft,
            this.tsBtnCenter,
            this.tsBtnRight});
            this.tsWordPadTop.Location = new System.Drawing.Point(0, 24);
            this.tsWordPadTop.Name = "tsWordPadTop";
            this.tsWordPadTop.Size = new System.Drawing.Size(518, 25);
            this.tsWordPadTop.TabIndex = 1;
            this.tsWordPadTop.Text = "toolStripWordPadTop";
            // 
            // tsCbFontStyle
            // 
            this.tsCbFontStyle.Name = "tsCbFontStyle";
            this.tsCbFontStyle.Size = new System.Drawing.Size(200, 25);
            this.tsCbFontStyle.Sorted = true;
            this.tsCbFontStyle.ToolTipText = "Font";
            this.tsCbFontStyle.SelectedIndexChanged += new System.EventHandler(this.tsCbFontStyle_SelectedIndexChanged);
            this.tsCbFontStyle.Leave += new System.EventHandler(this.tsCbFontSize_Leave);
            this.tsCbFontStyle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tsCbFontSize_KeyDown);
            // 
            // tsCbFontSize
            // 
            this.tsCbFontSize.Items.AddRange(new object[] {
            "8",
            "9",
            "10",
            "11",
            "12",
            "14",
            "16",
            "18",
            "20",
            "22",
            "24",
            "26",
            "28",
            "36",
            "48",
            "72"});
            this.tsCbFontSize.Name = "tsCbFontSize";
            this.tsCbFontSize.Size = new System.Drawing.Size(75, 25);
            this.tsCbFontSize.ToolTipText = "Size";
            this.tsCbFontSize.SelectedIndexChanged += new System.EventHandler(this.tsCbFontSize_SelectedIndexChanged);
            this.tsCbFontSize.Leave += new System.EventHandler(this.tsCbFontSize_Leave);
            this.tsCbFontSize.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tsCbFontSize_KeyDown);
            // 
            // tsBtnFontColor
            // 
            this.tsBtnFontColor.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsBtnFontColor.Image = ((System.Drawing.Image)(resources.GetObject("tsBtnFontColor.Image")));
            this.tsBtnFontColor.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsBtnFontColor.Name = "tsBtnFontColor";
            this.tsBtnFontColor.Size = new System.Drawing.Size(23, 22);
            this.tsBtnFontColor.Text = "A";
            this.tsBtnFontColor.ToolTipText = "Color";
            this.tsBtnFontColor.Click += new System.EventHandler(this.tsBtnFontColor_Click);
            // 
            // tsBtnHightLight
            // 
            this.tsBtnHightLight.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsBtnHightLight.Image = ((System.Drawing.Image)(resources.GetObject("tsBtnHightLight.Image")));
            this.tsBtnHightLight.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsBtnHightLight.Name = "tsBtnHightLight";
            this.tsBtnHightLight.Size = new System.Drawing.Size(23, 22);
            this.tsBtnHightLight.Text = "H";
            this.tsBtnHightLight.ToolTipText = "Hightlight";
            this.tsBtnHightLight.Click += new System.EventHandler(this.tsBtnHightlight_Click);
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new System.Drawing.Size(6, 25);
            // 
            // tsBtnBold
            // 
            this.tsBtnBold.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsBtnBold.Image = ((System.Drawing.Image)(resources.GetObject("tsBtnBold.Image")));
            this.tsBtnBold.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsBtnBold.Name = "tsBtnBold";
            this.tsBtnBold.Size = new System.Drawing.Size(23, 22);
            this.tsBtnBold.Text = "toolStripButtonBold";
            this.tsBtnBold.ToolTipText = "Bold\r\nCTRL + B";
            this.tsBtnBold.Click += new System.EventHandler(this.boldToolStripMenuItem_Click);
            // 
            // tsBtnItalic
            // 
            this.tsBtnItalic.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsBtnItalic.Image = ((System.Drawing.Image)(resources.GetObject("tsBtnItalic.Image")));
            this.tsBtnItalic.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsBtnItalic.Name = "tsBtnItalic";
            this.tsBtnItalic.Size = new System.Drawing.Size(23, 22);
            this.tsBtnItalic.Text = "toolStripButtonItalic";
            this.tsBtnItalic.ToolTipText = "Italic\r\n CTRL + I";
            this.tsBtnItalic.Click += new System.EventHandler(this.italicToolStripMenuItem_Click);
            // 
            // tsBtnUnderline
            // 
            this.tsBtnUnderline.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsBtnUnderline.Image = ((System.Drawing.Image)(resources.GetObject("tsBtnUnderline.Image")));
            this.tsBtnUnderline.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsBtnUnderline.Name = "tsBtnUnderline";
            this.tsBtnUnderline.Size = new System.Drawing.Size(23, 22);
            this.tsBtnUnderline.Text = "toolStripButtonUnderline";
            this.tsBtnUnderline.ToolTipText = "Underline\r\nCRTL + U";
            this.tsBtnUnderline.Click += new System.EventHandler(this.underlineToolStripMenuItem_Click);
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(6, 25);
            // 
            // tsBtnLeft
            // 
            this.tsBtnLeft.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsBtnLeft.Image = ((System.Drawing.Image)(resources.GetObject("tsBtnLeft.Image")));
            this.tsBtnLeft.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsBtnLeft.Name = "tsBtnLeft";
            this.tsBtnLeft.Size = new System.Drawing.Size(23, 22);
            this.tsBtnLeft.Text = "toolStripButtonLeft";
            this.tsBtnLeft.ToolTipText = "Left\r\nCtrl + L";
            this.tsBtnLeft.Click += new System.EventHandler(this.leftToolStripMenuItem_Click);
            // 
            // tsBtnCenter
            // 
            this.tsBtnCenter.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsBtnCenter.Image = ((System.Drawing.Image)(resources.GetObject("tsBtnCenter.Image")));
            this.tsBtnCenter.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsBtnCenter.Name = "tsBtnCenter";
            this.tsBtnCenter.Size = new System.Drawing.Size(23, 22);
            this.tsBtnCenter.Text = "toolStripButtonCenter";
            this.tsBtnCenter.ToolTipText = "Center\r\nCTRL + E";
            this.tsBtnCenter.Click += new System.EventHandler(this.centerToolStripMenuItem_Click);
            // 
            // tsBtnRight
            // 
            this.tsBtnRight.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsBtnRight.Image = ((System.Drawing.Image)(resources.GetObject("tsBtnRight.Image")));
            this.tsBtnRight.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsBtnRight.Name = "tsBtnRight";
            this.tsBtnRight.Size = new System.Drawing.Size(23, 22);
            this.tsBtnRight.Text = "toolStripButtonRight";
            this.tsBtnRight.ToolTipText = "Right\r\nCTRL + R";
            this.tsBtnRight.Click += new System.EventHandler(this.rightToolStripMenuItem_Click);
            // 
            // openFDWordPad
            // 
            this.openFDWordPad.FileName = "openFileDialog1";
            // 
            // frmWordPad
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(797, 376);
            this.Controls.Add(this.tsWordPadTop);
            this.Controls.Add(this.tsWordPadBottom);
            this.Controls.Add(this.rtxtWordPad);
            this.Controls.Add(this.menuStripWordPad);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStripWordPad;
            this.Name = "frmWordPad";
            this.Text = "Mini WordPad";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmWordPad_FormClosing);
            this.menuStripWordPad.ResumeLayout(false);
            this.menuStripWordPad.PerformLayout();
            this.tsWordPadBottom.ResumeLayout(false);
            this.tsWordPadBottom.PerformLayout();
            this.tsWordPadTop.ResumeLayout(false);
            this.tsWordPadTop.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStripWordPad;
        private System.Windows.Forms.ToolStripMenuItem tsmiFile;
        private System.Windows.Forms.ToolStripMenuItem newToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveAsToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tsmiFormat;
        private System.Windows.Forms.ToolStripMenuItem fontToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tsmiEdit;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem selectAllToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tsmiAbout;
        private System.Windows.Forms.ToolStripMenuItem infoToolStripMenuItem;
        private System.Windows.Forms.RichTextBox rtxtWordPad;
        private System.Windows.Forms.ToolStrip tsWordPadBottom;
        private System.Windows.Forms.ToolStripDropDownButton tsDropDownBtn;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton tsBtnSave;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripButton tsBtnUndo;
        private System.Windows.Forms.ToolStripButton tsBtnRedo;
        private System.Windows.Forms.ToolTip ttWordPad;
        private System.Windows.Forms.ToolStripMenuItem newToolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem saveAsToolStripMenuItem2;
        private System.Windows.Forms.ToolStripSeparator saveAsToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem newToolStripMenuItem1;
        private System.Windows.Forms.ToolStrip tsWordPadTop;
        private System.Windows.Forms.ToolStripComboBox tsCbFontStyle;
        private System.Windows.Forms.ToolStripButton tsBtnBold;
        private System.Windows.Forms.ToolStripButton tsBtnItalic;
        private System.Windows.Forms.ToolStripButton tsBtnUnderline;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripButton tsBtnLeft;
        private System.Windows.Forms.ToolStripButton tsBtnCenter;
        private System.Windows.Forms.ToolStripButton tsBtnRight;
        private System.Windows.Forms.ToolStripMenuItem tsmiFontStyle;
        private System.Windows.Forms.ToolStripMenuItem boldToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem italicToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem underlineToolStripMenuItem;
        private System.Windows.Forms.SaveFileDialog saveFDWordPad;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem leftToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem centerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem rightToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem justifyToolStripMenuItem;
        private System.Windows.Forms.OpenFileDialog openFDWordPad;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
        private System.Windows.Forms.ToolStripButton tsBtnFontColor;
        private System.Windows.Forms.ToolStripComboBox tsCbFontSize;
        private System.Windows.Forms.ToolStripMenuItem findToolStripMenuItem;
        private System.Windows.Forms.ToolStripButton tsBtnHightLight;
    }
}

