using System;
using System.Drawing;
using System.Windows.Forms;

namespace Autoformat_v1
{
    public class DesignSelectionTaskPane : UserControl
    {
        private Button btnApplyDesign;
        private Button btnCancel;
        private Panel pnlDesign1;
        private Panel pnlDesign2;
        private Panel pnlDesign3;
        private Panel pnlDesign4;
        private Label lblSelectedBadge;
        private int _selectedDesign = 1;
        private CheckBox chkAddIcons;

        public DesignSelectionTaskPane()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.BackColor = Color.FromArgb(30, 30, 30);
            this.ForeColor = Color.White;
            this.AutoScroll = true;
            this.Font = new Font("Segoe UI", 9F);

            int currentY = 20;
            int marginX = 20;
            int controlWidth = 310;

            // Title
            Label lblTitle = new Label
            {
                Text = "Choose a design",
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                Location = new Point(marginX, currentY),
                AutoSize = true
            };
            this.Controls.Add(lblTitle);
            currentY += 45;

            // Search box
            TextBox txtSearch = new TextBox
            {
                Text = "Search templates...",
                ForeColor = Color.DarkGray,
                BackColor = Color.FromArgb(45, 45, 45),
                BorderStyle = BorderStyle.FixedSingle,
                Location = new Point(marginX, currentY),
                Width = controlWidth,
                Font = new Font("Segoe UI", 10F)
            };
            
            txtSearch.GotFocus += (s, e) => { 
                if (txtSearch.Text == "Search templates...") { txtSearch.Text = ""; txtSearch.ForeColor = Color.White; } 
            };
            txtSearch.LostFocus += (s, e) => { 
                if (string.IsNullOrWhiteSpace(txtSearch.Text)) { txtSearch.Text = "Search templates..."; txtSearch.ForeColor = Color.DarkGray; } 
            };
            
            this.Controls.Add(txtSearch);
            currentY += 50;

            // Category Label
            Label lblCategory = new Label
            {
                Text = "Card layouts",
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(200, 200, 200),
                Location = new Point(marginX, currentY),
                AutoSize = true
            };
            this.Controls.Add(lblCategory);
            currentY += 30;

            // The reusable Selected badge
            lblSelectedBadge = new Label
            {
                Text = "Selected",
                Font = new Font("Segoe UI", 8F, FontStyle.Bold),
                BackColor = Color.FromArgb(13, 71, 161),
                ForeColor = Color.White,
                AutoSize = false,
                Size = new Size(60, 22),
                TextAlign = ContentAlignment.MiddleCenter
            };

            // Design 1
            pnlDesign1 = CreateDesignCard(1, "Header + body card", "2-4 rows", marginX, controlWidth, currentY);
            currentY += 85;

            // Design 2
            pnlDesign2 = CreateDesignCard(2, "Flush header card", "Right fade dividers", marginX, controlWidth, currentY);
            currentY += 85;

            // Design 3
            pnlDesign3 = CreateDesignCard(3, "Subheader list", "Subheader + Body", marginX, controlWidth, currentY);
            currentY += 85;

            // Design 4 (KPI)
            pnlDesign4 = CreateDesignCard(4, "KPI card", "Large number on top", marginX, controlWidth, currentY);
            currentY += 85;

            SelectDesign(1); // Default selection

            // Option B Auto-Icons Toggle
            chkAddIcons = new CheckBox
            {
                Text = "Auto-generate matching icons",
                Font = new Font("Segoe UI", 9.5F),
                Location = new Point(marginX, currentY + 10),
                AutoSize = true,
                Checked = true,
                Cursor = Cursors.Hand
            };
            this.Controls.Add(chkAddIcons);
            currentY += 40;

            // Bottom Buttons Panel
            Panel pnlButtons = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 70,
                BackColor = Color.FromArgb(30, 30, 30) 
            };

            pnlButtons.Paint += (s, e) =>
            {
                e.Graphics.DrawLine(new Pen(Color.FromArgb(60, 60, 60)), 0, 0, pnlButtons.Width, 0);
            };

            btnCancel = new Button
            {
                Text = "Cancel",
                FlatStyle = FlatStyle.Flat,
                Location = new Point(marginX, 15),
                Size = new Size(100, 35),
                BackColor = Color.FromArgb(45, 45, 45),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);

            btnApplyDesign = new Button
            {
                Text = "Apply design",
                FlatStyle = FlatStyle.Flat,
                Location = new Point(marginX + 110, 15),
                Size = new Size(controlWidth - 110, 35),
                BackColor = Color.White,
                ForeColor = Color.Black,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };

            btnApplyDesign.Click += BtnApplyDesign_Click;
            btnCancel.Click += BtnCancel_Click;

            pnlButtons.Controls.Add(btnCancel);
            pnlButtons.Controls.Add(btnApplyDesign);
            this.Controls.Add(pnlButtons);
        }

        private Panel CreateDesignCard(int designIndex, string title, string desc, int x, int width, int y)
        {
            Panel pnl = new Panel
            {
                Location = new Point(x, y),
                Width = width,
                Height = 75,
                Cursor = Cursors.Hand
            };

            pnl.Paint += (s, e) =>
            {
                Color borderColor = (_selectedDesign == designIndex) ? Color.FromArgb(60, 120, 216) : Color.FromArgb(60, 60, 60);
                ControlPaint.DrawBorder(e.Graphics, pnl.ClientRectangle, borderColor, ButtonBorderStyle.Solid);
            };

            Label lblTitle = new Label { Text = title, Font = new Font("Segoe UI", 11F, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true };
            Label lblDesc = new Label { Text = desc, Font = new Font("Segoe UI", 9F), ForeColor = Color.LightGray, Location = new Point(15, 40), AutoSize = true };

            // Click passthrough
            lblTitle.Click += (s, e) => SelectDesign(designIndex);
            lblDesc.Click += (s, e) => SelectDesign(designIndex);
            pnl.Click += (s, e) => SelectDesign(designIndex);

            pnl.Controls.Add(lblTitle);
            pnl.Controls.Add(lblDesc);
            this.Controls.Add(pnl);

            return pnl;
        }

        private void SelectDesign(int designIndex)
        {
            _selectedDesign = designIndex;

            pnlDesign1.BackColor = Color.FromArgb(34, 34, 34);
            pnlDesign2.BackColor = Color.FromArgb(34, 34, 34);
            if (pnlDesign3 != null) pnlDesign3.BackColor = Color.FromArgb(34, 34, 34);
            if (pnlDesign4 != null) pnlDesign4.BackColor = Color.FromArgb(34, 34, 34);

            Panel selectedPanel = pnlDesign1;
            if (designIndex == 2) selectedPanel = pnlDesign2;
            else if (designIndex == 3) selectedPanel = pnlDesign3;
            else if (designIndex == 4) selectedPanel = pnlDesign4;

            selectedPanel.BackColor = Color.FromArgb(40, 50, 70);
            
            selectedPanel.Controls.Add(lblSelectedBadge);
            lblSelectedBadge.Location = new Point(selectedPanel.Width - 80, 25);

            pnlDesign1.Invalidate();
            pnlDesign2.Invalidate();
            if (pnlDesign3 != null) pnlDesign3.Invalidate();
            if (pnlDesign4 != null) pnlDesign4.Invalidate();
        }

        private void BtnApplyDesign_Click(object sender, EventArgs e)
        {
            string designName = "Design 1";
            if (_selectedDesign == 2) designName = "Design 2";
            else if (_selectedDesign == 3) designName = "Design 1 CRD-SUB-LST";
            else if (_selectedDesign == 4) designName = "KPI-CRD-V1";
            
            FormattingAgent.ApplyDesign(designName, chkAddIcons.Checked);
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.HideTaskPane();
        }
    }
}
