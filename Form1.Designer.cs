using System;
using System.Drawing;
using System.Windows.Forms;

namespace HMS
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private TabControl tabControlMain;
        private TabPage tabDashboard;
        private TabPage tabAddPatient;
        private TabPage tabViewPatients;
        private TabPage tabReports;
        private StatusStrip statusStripMain;
        private ToolStripStatusLabel lblStatus;
        private ToolStripStatusLabel lblDateTime;

        // Dashboard Controls
        private Panel panelHeader;
        private Label lblHospitalName;
        private Label lblHeaderSubtitle;
        private Panel panelStatsTotal;
        private Panel panelStatsToday;
        private Panel panelStatsDoctors;
        private Panel panelStatsWards;
        private DataGridView dgvRecentPatients;
        private Label lblRecentPatients;
        private TextBox txtQuickSearch;
        private Button btnQuickAdd;
        private Label lblQuickSearch;

        // Add Patient Controls
        private GroupBox grpPersonalInfo;
        private GroupBox grpMedicalInfo;
        private GroupBox grpPhotoSection;
        private Label lblFullName;
        private TextBox txtFullName;
        private Label lblAge;
        private TextBox txtAge;
        private Label lblGender;
        private ComboBox cmbGender;
        private Label lblDOB;
        private DateTimePicker dtpDOB;
        private Label lblContact;
        private TextBox txtContact;
        private Label lblEmail;
        private TextBox txtEmail;
        private Label lblAddress;
        private TextBox txtAddress;
        private Label lblEmergencyContact;
        private TextBox txtEmergencyContact;
        private Label lblBloodGroup;
        private ComboBox cmbBloodGroup;
        private Label lblDoctorAssigned;
        private TextBox txtDoctorAssigned;
        private Label lblWardNumber;
        private TextBox txtWardNumber;
        private Label lblMedicalHistory;
        private TextBox txtMedicalHistory;
        private Label lblAllergies;
        private TextBox txtAllergies;
        private Label lblInsuranceProvider;
        private TextBox txtInsuranceProvider;
        private Label lblAdmissionDate;
        private DateTimePicker dtpAdmissionDate;
        private PictureBox pbPatientPhoto;
        private Button btnUploadPhoto;
        private Button btnSavePatient;
        private Button btnClearForm;
        private Label lblPatientID;
        private TextBox txtPatientID;

        // View Patients Controls
        private DataGridView dgvPatients;
        private Panel panelViewTop;
        private TextBox txtViewSearch;
        private ComboBox cmbFilter;
        private Label lblViewSearch;
        private Button btnExportJSON;
        private Button btnExportExcel;
        private Button btnExportText;
        private ComboBox cmbActionFilter;

        // Reports Controls
        private Panel panelReportsControl;
        private Button btnDailyReport;
        private Button btnWardReport;
        private Button btnDoctorReport;
        private RichTextBox rtbReport;
        private Button btnPrintReport;
        private Button btnExportReportExcel;
        private Label lblReportsTitle;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            // ============ MAIN FORM SETTINGS ============
            this.AutoScaleMode = AutoScaleMode.Font;   // Font-based DPI scaling (was None — caused overlapping)
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.Text = "Hospital Patient Records Management System";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(240, 244, 248);
            this.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            this.MinimumSize = new Size(1024, 600);

            // ============ HEADER PANEL ============
            // Uses a single-row TableLayoutPanel:
            //   col0 = icon+title (AutoSize), col1 = subtitle (fills rest)
            // Panel height is fixed; both labels are vertically centred inside.
            this.panelHeader = new Panel();
            this.panelHeader.Dock = DockStyle.Top;
            this.panelHeader.Height = 70;
            this.panelHeader.BackColor = Color.FromArgb(33, 72, 102);
            this.panelHeader.Padding = new Padding(0);

            TableLayoutPanel hdrLayout = new TableLayoutPanel();
            hdrLayout.Dock = DockStyle.Fill;
            hdrLayout.ColumnCount = 2;
            hdrLayout.RowCount = 1;
            // Title column: auto-size so it never clips; subtitle takes the rest
            hdrLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            hdrLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            hdrLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            hdrLayout.Padding = new Padding(20, 0, 20, 0);

            this.lblHospitalName = new Label();
            this.lblHospitalName.Text = "🏥  MEDICARE HOSPITAL";
            this.lblHospitalName.Font = new Font("Segoe UI", 20F, FontStyle.Bold);
            this.lblHospitalName.ForeColor = Color.White;
            this.lblHospitalName.BackColor = Color.Transparent;
            this.lblHospitalName.AutoSize = true;          // lets the cell measure correctly
            this.lblHospitalName.Dock = DockStyle.Fill;
            this.lblHospitalName.TextAlign = ContentAlignment.MiddleLeft;
            this.lblHospitalName.Margin = new Padding(0, 0, 20, 0);

            this.lblHeaderSubtitle = new Label();
            this.lblHeaderSubtitle.Text = "Patient Records Management System";
            this.lblHeaderSubtitle.Font = new Font("Segoe UI", 16F, FontStyle.Bold);
            this.lblHeaderSubtitle.ForeColor = Color.FromArgb(175, 208, 232);
            this.lblHeaderSubtitle.BackColor = Color.Transparent;
            this.lblHeaderSubtitle.AutoSize = false;
            this.lblHeaderSubtitle.Dock = DockStyle.Fill;
            this.lblHeaderSubtitle.TextAlign = ContentAlignment.MiddleLeft;
            this.lblHeaderSubtitle.Margin = new Padding(0);

            hdrLayout.Controls.Add(this.lblHospitalName, 0, 0);
            hdrLayout.Controls.Add(this.lblHeaderSubtitle, 1, 0);
            this.panelHeader.Controls.Add(hdrLayout);

            // ============ STATUS STRIP ============
            this.statusStripMain = new StatusStrip();
            this.statusStripMain.BackColor = Color.FromArgb(33, 72, 102);
            this.statusStripMain.ForeColor = Color.White;
            this.statusStripMain.Font = new Font("Segoe UI", 9F);
            this.statusStripMain.SizingGrip = false;

            this.lblStatus = new ToolStripStatusLabel();
            this.lblStatus.Text = "Ready";
            this.lblStatus.ForeColor = Color.FromArgb(180, 255, 180);
            this.lblStatus.Padding = new Padding(8, 0, 0, 0);

            this.lblDateTime = new ToolStripStatusLabel();
            this.lblDateTime.Text = DateTime.Now.ToString("dddd, MMMM dd, yyyy | HH:mm:ss");
            this.lblDateTime.ForeColor = Color.FromArgb(200, 224, 240);
            this.lblDateTime.Spring = true;
            this.lblDateTime.TextAlign = ContentAlignment.MiddleRight;
            this.lblDateTime.Padding = new Padding(0, 0, 10, 0);

            this.statusStripMain.Items.AddRange(new ToolStripItem[] {
                this.lblStatus,
                this.lblDateTime
            });

            // ============ TAB CONTROL ============
            this.tabControlMain = new TabControl();
            this.tabControlMain.Dock = DockStyle.Fill;
            this.tabControlMain.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            this.tabControlMain.DrawMode = TabDrawMode.OwnerDrawFixed;
            this.tabControlMain.ItemSize = new Size(160, 38);
            this.tabControlMain.SizeMode = TabSizeMode.Fixed;
            this.tabControlMain.Appearance = TabAppearance.Normal;
            this.tabControlMain.Padding = new Point(0, 0);
            this.tabControlMain.Margin = new Padding(0);
            this.tabControlMain.DrawItem += new DrawItemEventHandler(this.TabControlMain_DrawItem);

            // ============ TABS ============
            this.tabDashboard = new TabPage("  📊  Dashboard  ");
            this.tabDashboard.BackColor = Color.FromArgb(240, 244, 248);
            this.tabDashboard.UseVisualStyleBackColor = false;
            this.BuildDashboardTab();

            this.tabAddPatient = new TabPage("  ➕  Add Patient  ");
            this.tabAddPatient.BackColor = Color.FromArgb(240, 244, 248);
            this.tabAddPatient.AutoScroll = true;
            this.tabAddPatient.UseVisualStyleBackColor = false;
            this.BuildAddPatientTab();

            this.tabViewPatients = new TabPage("  👥  View Patients  ");
            this.tabViewPatients.BackColor = Color.FromArgb(240, 244, 248);
            this.tabViewPatients.UseVisualStyleBackColor = false;
            this.BuildViewPatientsTab();

            this.tabReports = new TabPage("  📈  Reports  ");
            this.tabReports.BackColor = Color.FromArgb(240, 244, 248);
            this.tabReports.UseVisualStyleBackColor = false;
            this.BuildReportsTab();

            this.tabControlMain.TabPages.AddRange(new TabPage[] {
                this.tabDashboard,
                this.tabAddPatient,
                this.tabViewPatients,
                this.tabReports
            });

            // ============ ADD TO FORM ============
            this.Controls.Add(this.tabControlMain);
            this.Controls.Add(this.panelHeader);
            this.Controls.Add(this.statusStripMain);

            this.Load += new EventHandler(this.Form1_Load);
        }

        // =====================================================================
        //  DASHBOARD TAB
        // =====================================================================
        private void BuildDashboardTab()
        {
            // Outer fill panel with padding
            Panel outer = new Panel();
            outer.Dock = DockStyle.Fill;
            outer.Padding = new Padding(16);

            // ---- Stats row ----
            TableLayoutPanel statsTable = new TableLayoutPanel();
            statsTable.Dock = DockStyle.Top;
            statsTable.Height = 145;
            statsTable.ColumnCount = 4;
            statsTable.RowCount = 1;
            statsTable.Padding = new Padding(0, 0, 0, 12);
            for (int i = 0; i < 4; i++)
                statsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));

            this.panelStatsTotal = CreateStatCard(Color.FromArgb(59, 130, 246), "Total Patients", "0", "👥");
            this.panelStatsToday = CreateStatCard(Color.FromArgb(16, 185, 129), "Today's Admissions", "0", "🏥");
            this.panelStatsDoctors = CreateStatCard(Color.FromArgb(245, 158, 11), "Active Doctors", "0", "👨‍⚕️");
            this.panelStatsWards = CreateStatCard(Color.FromArgb(139, 92, 246), "Occupied Wards", "0", "🛏️");

            statsTable.Controls.Add(this.panelStatsTotal, 0, 0);
            statsTable.Controls.Add(this.panelStatsToday, 1, 0);
            statsTable.Controls.Add(this.panelStatsDoctors, 2, 0);
            statsTable.Controls.Add(this.panelStatsWards, 3, 0);

            // ---- Recent Patients section ----
            Panel recentSection = new Panel();
            recentSection.Dock = DockStyle.Fill;

            // Toolbar inside recentSection
            TableLayoutPanel toolbar = new TableLayoutPanel();
            toolbar.Dock = DockStyle.Top;
            toolbar.Height = 44;
            toolbar.ColumnCount = 3;
            toolbar.RowCount = 1;
            toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); // title — takes remaining
            toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220F));
            toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 165F));
            toolbar.Padding = new Padding(0, 0, 0, 6);

            this.lblRecentPatients = new Label();
            this.lblRecentPatients.Text = "Recent Admissions";
            this.lblRecentPatients.Font = new Font("Segoe UI", 13F, FontStyle.Bold);
            this.lblRecentPatients.ForeColor = Color.FromArgb(30, 41, 59);
            this.lblRecentPatients.Dock = DockStyle.Fill;
            this.lblRecentPatients.TextAlign = ContentAlignment.MiddleLeft;

            this.txtQuickSearch = new TextBox();
            this.txtQuickSearch.Dock = DockStyle.Fill;
            this.txtQuickSearch.Font = new Font("Segoe UI", 10F);
            this.txtQuickSearch.Text = "Search patients...";
            this.txtQuickSearch.ForeColor = Color.Gray;
            this.txtQuickSearch.Margin = new Padding(0, 4, 8, 4);
            this.txtQuickSearch.TextChanged += new EventHandler(this.TxtQuickSearch_TextChanged);
            this.txtQuickSearch.Enter += new EventHandler(this.TxtQuickSearch_Enter);
            this.txtQuickSearch.Leave += new EventHandler(this.TxtQuickSearch_Leave);

            this.btnQuickAdd = new Button();
            this.btnQuickAdd.Text = "➕  Add Patient";
            this.btnQuickAdd.Dock = DockStyle.Fill;
            this.btnQuickAdd.FlatStyle = FlatStyle.Flat;
            this.btnQuickAdd.BackColor = Color.FromArgb(76, 175, 80);
            this.btnQuickAdd.ForeColor = Color.White;
            this.btnQuickAdd.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.btnQuickAdd.FlatAppearance.BorderSize = 0;
            this.btnQuickAdd.Cursor = Cursors.Hand;
            this.btnQuickAdd.Margin = new Padding(0, 4, 0, 4);
            this.btnQuickAdd.Click += new EventHandler(this.BtnQuickAdd_Click);

            toolbar.Controls.Add(this.lblRecentPatients, 0, 0);
            toolbar.Controls.Add(this.txtQuickSearch, 1, 0);
            toolbar.Controls.Add(this.btnQuickAdd, 2, 0);

            // DataGridView
            this.dgvRecentPatients = BuildStyledGrid();
            this.dgvRecentPatients.Columns.Add("PatientID", "Patient ID");
            this.dgvRecentPatients.Columns.Add("FullName", "Full Name");
            this.dgvRecentPatients.Columns.Add("Age", "Age");
            this.dgvRecentPatients.Columns.Add("Gender", "Gender");
            this.dgvRecentPatients.Columns.Add("DoctorAssigned", "Doctor");
            this.dgvRecentPatients.Columns.Add("WardNumber", "Ward");
            this.dgvRecentPatients.Columns.Add("AdmissionDate", "Admission Date");
            // Age & Gender narrower
            this.dgvRecentPatients.Columns["Age"].FillWeight = 40;
            this.dgvRecentPatients.Columns["Gender"].FillWeight = 60;

            recentSection.Controls.Add(this.dgvRecentPatients);
            recentSection.Controls.Add(toolbar);

            outer.Controls.Add(recentSection);
            outer.Controls.Add(statsTable);
            this.tabDashboard.Controls.Add(outer);
        }

        private Panel CreateStatCard(Color color, string label, string count, string icon)
        {
            // Card outer container
            Panel card = new Panel();
            card.Dock = DockStyle.Fill;
            card.BackColor = Color.White;
            card.Margin = new Padding(5);

            // Left accent stripe (fixed width, full height via Dock)
            Panel accent = new Panel();
            accent.Width = 5;
            accent.Dock = DockStyle.Left;
            accent.BackColor = color;

            // Inner layout: icon on right (fixed 60px), text on left
            TableLayoutPanel inner = new TableLayoutPanel();
            inner.Dock = DockStyle.Fill;
            inner.ColumnCount = 2;
            inner.RowCount = 2;
            inner.Padding = new Padding(10, 8, 8, 8);
            // Left column takes all remaining space; right column fixed for icon
            inner.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            inner.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 54F));
            // Top row for count number; bottom row for label text
            inner.RowStyles.Add(new RowStyle(SizeType.Percent, 58F));
            inner.RowStyles.Add(new RowStyle(SizeType.Percent, 42F));

            // Big number
            Label lblCount = new Label();
            lblCount.Name = "lblCount";
            lblCount.Text = count;
            lblCount.Font = new Font("Segoe UI", 22F, FontStyle.Bold);
            lblCount.ForeColor = Color.FromArgb(30, 41, 59);
            lblCount.Dock = DockStyle.Fill;
            lblCount.TextAlign = ContentAlignment.BottomLeft;
            lblCount.Margin = new Padding(0);

            // Emoji icon — spans both rows, smaller font to avoid clipping
            Label lblIcon = new Label();
            lblIcon.Text = icon;
            lblIcon.Font = new Font("Segoe UI Emoji", 18F);
            lblIcon.ForeColor = color;
            lblIcon.Dock = DockStyle.Fill;
            lblIcon.TextAlign = ContentAlignment.MiddleCenter;
            lblIcon.Margin = new Padding(0);
            inner.SetRowSpan(lblIcon, 2);

            // Subtitle label
            Label lblLabel = new Label();
            lblLabel.Text = label;
            lblLabel.Font = new Font("Segoe UI", 9F);
            lblLabel.ForeColor = Color.FromArgb(100, 116, 139);
            lblLabel.Dock = DockStyle.Fill;
            lblLabel.TextAlign = ContentAlignment.TopLeft;
            lblLabel.Margin = new Padding(0);

            inner.Controls.Add(lblCount, 0, 0);
            inner.Controls.Add(lblIcon, 1, 0);
            inner.Controls.Add(lblLabel, 0, 1);

            card.Controls.Add(inner);
            card.Controls.Add(accent);
            return card;
        }

        // =====================================================================
        //  ADD PATIENT TAB
        // =====================================================================
        private void BuildAddPatientTab()
        {
            // Scrollable wrapper
            Panel scrollWrapper = new Panel();
            scrollWrapper.Dock = DockStyle.Fill;
            scrollWrapper.AutoScroll = true;
            scrollWrapper.Padding = new Padding(16);

            // ---- Patient ID header bar ----
            TableLayoutPanel idBar = new TableLayoutPanel();
            idBar.Dock = DockStyle.Top;
            idBar.Height = 44;
            idBar.ColumnCount = 3;
            idBar.RowCount = 1;
            idBar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            idBar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180F));
            idBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            idBar.Margin = new Padding(0, 0, 0, 8);

            this.lblPatientID = new Label();
            this.lblPatientID.Text = "Patient ID:";
            this.lblPatientID.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            this.lblPatientID.ForeColor = Color.FromArgb(30, 41, 59);
            this.lblPatientID.Dock = DockStyle.Fill;
            this.lblPatientID.TextAlign = ContentAlignment.MiddleLeft;
            this.lblPatientID.Padding = new Padding(0, 0, 8, 0);

            this.txtPatientID = new TextBox();
            this.txtPatientID.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            this.txtPatientID.ForeColor = Color.FromArgb(42, 92, 130);
            this.txtPatientID.BackColor = Color.FromArgb(220, 234, 245);
            this.txtPatientID.ReadOnly = true;
            this.txtPatientID.BorderStyle = BorderStyle.FixedSingle;
            this.txtPatientID.Dock = DockStyle.Fill;
            this.txtPatientID.Margin = new Padding(0, 4, 12, 4);

            idBar.Controls.Add(this.lblPatientID, 0, 0);
            idBar.Controls.Add(this.txtPatientID, 1, 0);

            // ---- Two-column main content ----
            TableLayoutPanel twoCol = new TableLayoutPanel();
            twoCol.Dock = DockStyle.Top;
            twoCol.AutoSize = true;
            twoCol.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            twoCol.ColumnCount = 2;
            twoCol.RowCount = 1;
            twoCol.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 63F));
            twoCol.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 37F));

            // === LEFT: Personal + Medical groups ===
            Panel leftCol = new Panel();
            leftCol.Dock = DockStyle.Fill;
            leftCol.Padding = new Padding(0, 0, 10, 0);
            leftCol.AutoSize = true;
            leftCol.AutoSizeMode = AutoSizeMode.GrowAndShrink;

            this.grpPersonalInfo = BuildGroupBox("Personal Information", 370);
            TableLayoutPanel personalGrid = Build4ColGrid(8);

            // Row 0 — Full Name (spans 3 cols)
            AddFieldRow(personalGrid, 0, 0, "Full Name *", out this.lblFullName, out this.txtFullName, colSpan: 3);

            // Row 1 — Age | DOB
            AddFieldRow(personalGrid, 0, 1, "Age", out this.lblAge, out this.txtAge);
            this.txtAge.ReadOnly = true;
            this.txtAge.BackColor = Color.FromArgb(240, 244, 248);

            AddLabel(personalGrid, "Date of Birth *", 2, 1);
            this.dtpDOB = MakeDTP(DateTime.Now.AddYears(-25), true);
            this.dtpDOB.ValueChanged += new EventHandler(this.DtpDOB_ValueChanged);
            personalGrid.Controls.Add(this.dtpDOB, 3, 1);

            // Row 2 — Gender | Blood Group
            AddLabel(personalGrid, "Gender *", 0, 2);
            this.cmbGender = MakeCombo(new[] { "Male", "Female", "Other" });
            personalGrid.Controls.Add(this.cmbGender, 1, 2);

            AddLabel(personalGrid, "Blood Group", 2, 2);
            this.cmbBloodGroup = MakeCombo(new[] { "A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-" });
            personalGrid.Controls.Add(this.cmbBloodGroup, 3, 2);

            // Row 3 — Contact | Email
            AddFieldRow(personalGrid, 0, 3, "Contact No. *", out this.lblContact, out this.txtContact);
            AddFieldRow(personalGrid, 2, 3, "Email", out this.lblEmail, out this.txtEmail);

            // Row 4 — Address (spans 3)
            AddFieldRow(personalGrid, 0, 4, "Address", out this.lblAddress, out this.txtAddress, colSpan: 3);

            // Row 5 — Emergency Contact
            AddFieldRow(personalGrid, 0, 5, "Emergency Contact", out this.lblEmergencyContact, out this.txtEmergencyContact);

            this.grpPersonalInfo.Controls.Add(personalGrid);

            this.grpMedicalInfo = BuildGroupBox("Medical Information", 350);
            TableLayoutPanel medGrid = Build4ColGrid(6);
            medGrid.RowStyles[2].Height = 64F;
            medGrid.RowStyles[3].Height = 64F;

            // Row 0 — Doctor | Admission Date
            AddFieldRow(medGrid, 0, 0, "Doctor Assigned *", out this.lblDoctorAssigned, out this.txtDoctorAssigned);
            AddLabel(medGrid, "Admission Date *", 2, 0);
            this.dtpAdmissionDate = MakeDTP(DateTime.Now, false);
            medGrid.Controls.Add(this.dtpAdmissionDate, 3, 0);

            // Row 1 — Ward | (empty)
            AddFieldRow(medGrid, 0, 1, "Ward Number", out this.lblWardNumber, out this.txtWardNumber);

            // Row 2 — Medical History (spans 3)
            AddFieldRow(medGrid, 0, 2, "Medical History", out this.lblMedicalHistory, out this.txtMedicalHistory, colSpan: 3);
            this.txtMedicalHistory.Multiline = true;

            // Row 3 — Allergies (spans 3)
            AddFieldRow(medGrid, 0, 3, "Allergies", out this.lblAllergies, out this.txtAllergies, colSpan: 3);
            this.txtAllergies.Multiline = true;

            // Row 4 — Insurance (spans 3)
            AddFieldRow(medGrid, 0, 4, "Insurance Provider", out this.lblInsuranceProvider, out this.txtInsuranceProvider, colSpan: 3);

            this.grpMedicalInfo.Controls.Add(medGrid);

            // Stack groups in leftCol (Medical on top because DockStyle.Top stacks in reverse)
            leftCol.Controls.Add(this.grpMedicalInfo);
            leftCol.Controls.Add(this.grpPersonalInfo);

            // === RIGHT: Photo + Buttons ===
            Panel rightCol = new Panel();
            rightCol.Dock = DockStyle.Fill;
            rightCol.Padding = new Padding(4, 0, 0, 0);
            rightCol.AutoSize = true;
            rightCol.AutoSizeMode = AutoSizeMode.GrowAndShrink;

            this.grpPhotoSection = BuildGroupBox("Patient Photo", 0); // height auto
            this.grpPhotoSection.Dock = DockStyle.Top;
            this.grpPhotoSection.Height = 330;

            this.pbPatientPhoto = new PictureBox();
            this.pbPatientPhoto.Dock = DockStyle.Top;
            this.pbPatientPhoto.Height = 240;
            this.pbPatientPhoto.Margin = new Padding(10, 4, 10, 8);
            this.pbPatientPhoto.BackColor = Color.FromArgb(235, 242, 250);
            this.pbPatientPhoto.BorderStyle = BorderStyle.FixedSingle;
            this.pbPatientPhoto.SizeMode = PictureBoxSizeMode.Zoom;

            this.btnUploadPhoto = MakeActionButton("📷  Upload Photo", Color.FromArgb(42, 92, 130));
            this.btnUploadPhoto.Dock = DockStyle.Top;
            this.btnUploadPhoto.Height = 42;
            this.btnUploadPhoto.Margin = new Padding(10, 0, 10, 8);
            this.btnUploadPhoto.Click += new EventHandler(this.BtnUploadPhoto_Click);

            this.grpPhotoSection.Controls.Add(this.btnUploadPhoto);
            this.grpPhotoSection.Controls.Add(this.pbPatientPhoto);

            // Action buttons panel
            TableLayoutPanel btnRow = new TableLayoutPanel();
            btnRow.Dock = DockStyle.Top;
            btnRow.Height = 54;
            btnRow.ColumnCount = 2;
            btnRow.RowCount = 1;
            btnRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            btnRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            btnRow.Margin = new Padding(0, 10, 0, 0);

            this.btnSavePatient = MakeActionButton("💾  Save", Color.FromArgb(76, 175, 80));
            this.btnSavePatient.Dock = DockStyle.Fill;
            this.btnSavePatient.Margin = new Padding(0, 0, 4, 0);
            this.btnSavePatient.Click += new EventHandler(this.BtnSavePatient_Click);

            this.btnClearForm = MakeActionButton("🗑️  Clear", Color.FromArgb(239, 68, 68));
            this.btnClearForm.Dock = DockStyle.Fill;
            this.btnClearForm.Margin = new Padding(4, 0, 0, 0);
            this.btnClearForm.Click += new EventHandler(this.BtnClearForm_Click);

            btnRow.Controls.Add(this.btnSavePatient, 0, 0);
            btnRow.Controls.Add(this.btnClearForm, 1, 0);

            rightCol.Controls.Add(btnRow);
            rightCol.Controls.Add(this.grpPhotoSection);

            twoCol.Controls.Add(leftCol, 0, 0);
            twoCol.Controls.Add(rightCol, 1, 0);

            scrollWrapper.Controls.Add(twoCol);
            scrollWrapper.Controls.Add(idBar);
            this.tabAddPatient.Controls.Add(scrollWrapper);
        }

        // =====================================================================
        //  VIEW PATIENTS TAB
        // =====================================================================
        private void BuildViewPatientsTab()
        {
            // Top toolbar — FlowLayoutPanel so buttons never overlap
            this.panelViewTop = new Panel();
            this.panelViewTop.Dock = DockStyle.Top;
            this.panelViewTop.Height = 56;
            this.panelViewTop.BackColor = Color.White;
            this.panelViewTop.Padding = new Padding(10, 0, 10, 0);

            FlowLayoutPanel flow = new FlowLayoutPanel();
            flow.Dock = DockStyle.Fill;
            flow.FlowDirection = FlowDirection.LeftToRight;
            flow.WrapContents = false;
            flow.AutoSize = false;
            flow.Padding = new Padding(0, 10, 0, 10);

            // Search icon label
            this.lblViewSearch = new Label();
            this.lblViewSearch.Text = "🔍";
            this.lblViewSearch.Font = new Font("Segoe UI", 12F);
            this.lblViewSearch.AutoSize = true;
            this.lblViewSearch.Margin = new Padding(0, 4, 4, 0);
            this.lblViewSearch.TextAlign = ContentAlignment.MiddleCenter;

            this.txtViewSearch = new TextBox();
            this.txtViewSearch.Width = 200;
            this.txtViewSearch.Height = 30;
            this.txtViewSearch.Font = new Font("Segoe UI", 10F);
            this.txtViewSearch.BorderStyle = BorderStyle.FixedSingle;
            this.txtViewSearch.Margin = new Padding(0, 2, 8, 0);
            this.txtViewSearch.TextChanged += new EventHandler(this.TxtViewSearch_TextChanged);

            this.cmbFilter = new ComboBox();
            this.cmbFilter.Width = 130;
            this.cmbFilter.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbFilter.Items.AddRange(new string[] { "All", "By Doctor", "By Ward", "By Blood", "Today" });
            this.cmbFilter.SelectedIndex = 0;
            this.cmbFilter.Font = new Font("Segoe UI", 10F);
            this.cmbFilter.Margin = new Padding(0, 2, 8, 0);
            this.cmbFilter.SelectedIndexChanged += new EventHandler(this.CmbFilter_SelectedIndexChanged);

            this.cmbActionFilter = new ComboBox();
            this.cmbActionFilter.Width = 150;
            this.cmbActionFilter.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbActionFilter.Font = new Font("Segoe UI", 10F);
            this.cmbActionFilter.Margin = new Padding(0, 2, 12, 0);
            this.cmbActionFilter.Visible = false;
            this.cmbActionFilter.SelectedIndexChanged += new EventHandler(this.CmbActionFilter_SelectedIndexChanged);

            // Separator
            Label sep = new Label();
            sep.Width = 1;
            sep.Height = 30;
            sep.BackColor = Color.FromArgb(200, 210, 220);
            sep.Margin = new Padding(0, 2, 12, 0);

            // Export buttons
            this.btnExportJSON = MakeFlowButton("📄 JSON", Color.FromArgb(42, 92, 130));
            this.btnExportExcel = MakeFlowButton("📊 Excel", Color.FromArgb(59, 130, 246));
            this.btnExportText = MakeFlowButton("📝 Text", Color.FromArgb(139, 92, 246));
            this.btnExportJSON.Click += new EventHandler(this.BtnExportJSON_Click);
            this.btnExportExcel.Click += new EventHandler(this.BtnExportExcel_Click);
            this.btnExportText.Click += new EventHandler(this.BtnExportText_Click);

            flow.Controls.Add(this.lblViewSearch);
            flow.Controls.Add(this.txtViewSearch);
            flow.Controls.Add(this.cmbFilter);
            flow.Controls.Add(this.cmbActionFilter);
            flow.Controls.Add(sep);
            flow.Controls.Add(this.btnExportJSON);
            flow.Controls.Add(this.btnExportExcel);
            flow.Controls.Add(this.btnExportText);

            this.panelViewTop.Controls.Add(flow);

            // Grid
            this.dgvPatients = BuildStyledGrid();
            this.dgvPatients.Columns.Add("PatientID", "Patient ID");
            this.dgvPatients.Columns.Add("FullName", "Full Name");
            this.dgvPatients.Columns.Add("Age", "Age");
            this.dgvPatients.Columns.Add("Gender", "Gender");
            this.dgvPatients.Columns.Add("ContactNumber", "Contact");
            this.dgvPatients.Columns.Add("DoctorAssigned", "Doctor");
            this.dgvPatients.Columns.Add("WardNumber", "Ward");
            this.dgvPatients.Columns.Add("AdmissionDate", "Admission Date");
            this.dgvPatients.Columns["Age"].FillWeight = 40;
            this.dgvPatients.Columns["Gender"].FillWeight = 55;
            this.dgvPatients.Columns["WardNumber"].FillWeight = 55;

            AddButtonCol(this.dgvPatients, "btnView", "👁 View", Color.FromArgb(59, 130, 246), Color.FromArgb(37, 99, 235), 90);
            AddButtonCol(this.dgvPatients, "btnEdit", "✏ Edit", Color.FromArgb(245, 158, 11), Color.FromArgb(217, 119, 6), 90);
            AddButtonCol(this.dgvPatients, "btnDelete", "🗑 Delete", Color.FromArgb(239, 68, 68), Color.FromArgb(220, 38, 38), 90);
            // Give action columns a fixed weight so they stay wide relative to data cols
            this.dgvPatients.Columns["btnView"].FillWeight = 55;
            this.dgvPatients.Columns["btnEdit"].FillWeight = 55;
            this.dgvPatients.Columns["btnDelete"].FillWeight = 55;
            this.dgvPatients.Columns["btnView"].MinimumWidth = 85;
            this.dgvPatients.Columns["btnEdit"].MinimumWidth = 85;
            this.dgvPatients.Columns["btnDelete"].MinimumWidth = 85;

            this.dgvPatients.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                if (e.ColumnIndex == this.dgvPatients.Columns["btnView"].Index)
                { e.CellStyle.BackColor = Color.FromArgb(59, 130, 246); e.CellStyle.ForeColor = Color.White; e.CellStyle.SelectionBackColor = Color.FromArgb(37, 99, 235); e.CellStyle.SelectionForeColor = Color.White; }
                else if (e.ColumnIndex == this.dgvPatients.Columns["btnEdit"].Index)
                { e.CellStyle.BackColor = Color.FromArgb(245, 158, 11); e.CellStyle.ForeColor = Color.White; e.CellStyle.SelectionBackColor = Color.FromArgb(217, 119, 6); e.CellStyle.SelectionForeColor = Color.White; }
                else if (e.ColumnIndex == this.dgvPatients.Columns["btnDelete"].Index)
                { e.CellStyle.BackColor = Color.FromArgb(239, 68, 68); e.CellStyle.ForeColor = Color.White; e.CellStyle.SelectionBackColor = Color.FromArgb(220, 38, 38); e.CellStyle.SelectionForeColor = Color.White; }
            };
            this.dgvPatients.CellClick += new DataGridViewCellEventHandler(this.DgvPatients_CellClick);

            this.tabViewPatients.Controls.Add(this.dgvPatients);
            this.tabViewPatients.Controls.Add(this.panelViewTop);
        }

        // =====================================================================
        //  REPORTS TAB
        // =====================================================================
        private void BuildReportsTab()
        {
            // Toolbar — TableLayoutPanel: 6 equal columns so buttons
            // spread across the full width on any screen size.
            this.panelReportsControl = new Panel();
            this.panelReportsControl.Dock = DockStyle.Top;
            this.panelReportsControl.Height = 62;
            this.panelReportsControl.BackColor = Color.White;
            this.panelReportsControl.Padding = new Padding(10, 0, 10, 0);

            TableLayoutPanel rptBar = new TableLayoutPanel();
            rptBar.Dock = DockStyle.Fill;
            rptBar.ColumnCount = 6;
            rptBar.RowCount = 1;
            rptBar.Padding = new Padding(0, 10, 0, 10);
            // Title label: 1.5x weight so it has a bit more room
            rptBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            rptBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16F));
            rptBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16F));
            rptBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16F));
            rptBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16F));
            rptBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16F));
            rptBar.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            this.lblReportsTitle = new Label();
            this.lblReportsTitle.Text = "📊  Reports";
            this.lblReportsTitle.Font = new Font("Segoe UI", 13F, FontStyle.Bold);
            this.lblReportsTitle.Dock = DockStyle.Fill;
            this.lblReportsTitle.ForeColor = Color.FromArgb(30, 41, 59);
            this.lblReportsTitle.TextAlign = ContentAlignment.MiddleLeft;
            this.lblReportsTitle.Margin = new Padding(4, 0, 8, 0);

            this.btnDailyReport = MakeBarButton("📅  Daily", Color.FromArgb(42, 92, 130));
            this.btnWardReport = MakeBarButton("🏥  Ward", Color.FromArgb(59, 130, 246));
            this.btnDoctorReport = MakeBarButton("👨‍⚕️  Doctor", Color.FromArgb(139, 92, 246));
            this.btnPrintReport = MakeBarButton("🖨️  Print", Color.FromArgb(245, 158, 11));
            this.btnExportReportExcel = MakeBarButton("📊  Export Excel", Color.FromArgb(16, 185, 129));

            this.btnDailyReport.Click += new EventHandler(this.BtnDailyReport_Click);
            this.btnWardReport.Click += new EventHandler(this.BtnWardReport_Click);
            this.btnDoctorReport.Click += new EventHandler(this.BtnDoctorReport_Click);
            this.btnPrintReport.Click += new EventHandler(this.BtnPrintReport_Click);
            this.btnExportReportExcel.Click += new EventHandler(this.BtnExportReportExcel_Click);

            rptBar.Controls.Add(this.lblReportsTitle, 0, 0);
            rptBar.Controls.Add(this.btnDailyReport, 1, 0);
            rptBar.Controls.Add(this.btnWardReport, 2, 0);
            rptBar.Controls.Add(this.btnDoctorReport, 3, 0);
            rptBar.Controls.Add(this.btnPrintReport, 4, 0);
            rptBar.Controls.Add(this.btnExportReportExcel, 5, 0);
            this.panelReportsControl.Controls.Add(rptBar);

            this.rtbReport = new RichTextBox();
            this.rtbReport.Dock = DockStyle.Fill;
            this.rtbReport.Font = new Font("Consolas", 10F);
            this.rtbReport.BackColor = Color.White;
            this.rtbReport.ReadOnly = true;
            this.rtbReport.BorderStyle = BorderStyle.None;
            this.rtbReport.Padding = new Padding(12);

            this.tabReports.Controls.Add(this.rtbReport);
            this.tabReports.Controls.Add(this.panelReportsControl);
        }

        // =====================================================================
        //  HELPERS
        // =====================================================================

        /// Shared styled DataGridView
        private DataGridView BuildStyledGrid()
        {
            var dgv = new DataGridView();
            dgv.Dock = DockStyle.Fill;
            dgv.BackgroundColor = Color.White;
            dgv.BorderStyle = BorderStyle.None;
            dgv.RowHeadersVisible = false;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.ReadOnly = true;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.DefaultCellStyle.Font = new Font("Segoe UI", 10F);
            dgv.DefaultCellStyle.Padding = new Padding(4);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(42, 92, 130);
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252);
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(33, 72, 102);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Padding = new Padding(4);
            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersHeight = 38;
            dgv.RowTemplate.Height = 38;
            dgv.GridColor = Color.FromArgb(225, 232, 240);
            return dgv;
        }

        /// 4-column TableLayoutPanel for form fields (label | field | label | field)
        private TableLayoutPanel Build4ColGrid(int rows)
        {
            var t = new TableLayoutPanel();
            t.Dock = DockStyle.Fill;
            t.ColumnCount = 4;
            t.RowCount = rows;
            t.Padding = new Padding(6, 4, 6, 4);
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 135F));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 135F));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            for (int i = 0; i < rows; i++)
                t.RowStyles.Add(new RowStyle(SizeType.Absolute, 44F));
            return t;
        }

        private GroupBox BuildGroupBox(string title, int height)
        {
            var g = new GroupBox();
            g.Text = title;
            g.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            g.ForeColor = Color.FromArgb(42, 92, 130);
            g.BackColor = Color.White;
            g.Dock = DockStyle.Top;
            if (height > 0) g.Height = height;
            g.Padding = new Padding(10, 22, 10, 10);
            g.Margin = new Padding(0, 0, 0, 12);
            return g;
        }

        private void AddFieldRow(TableLayoutPanel t, int col, int row, string labelText,
            out Label label, out TextBox textBox, int colSpan = 0)
        {
            label = new Label();
            label.Text = labelText;
            label.Font = new Font("Segoe UI", 10F);
            label.ForeColor = Color.FromArgb(71, 85, 105);
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.Padding = new Padding(0, 0, 4, 0);
            t.Controls.Add(label, col, row);

            textBox = new TextBox();
            textBox.Font = new Font("Segoe UI", 10F);
            textBox.BorderStyle = BorderStyle.FixedSingle;
            textBox.Dock = DockStyle.Fill;
            textBox.Margin = new Padding(0, 4, 4, 4);
            t.Controls.Add(textBox, col + 1, row);
            if (colSpan > 0) t.SetColumnSpan(textBox, colSpan);
        }

        private void AddLabel(TableLayoutPanel t, string text, int col, int row)
        {
            var l = new Label();
            l.Text = text;
            l.Font = new Font("Segoe UI", 10F);
            l.ForeColor = Color.FromArgb(71, 85, 105);
            l.Dock = DockStyle.Fill;
            l.TextAlign = ContentAlignment.MiddleLeft;
            l.Padding = new Padding(4, 0, 4, 0);
            t.Controls.Add(l, col, row);
        }

        private DateTimePicker MakeDTP(DateTime value, bool hasMaxDate)
        {
            var d = new DateTimePicker();
            d.Font = new Font("Segoe UI", 10F);
            d.Value = value;
            if (hasMaxDate) d.MaxDate = DateTime.Now;
            d.Format = DateTimePickerFormat.Short;
            d.Dock = DockStyle.Fill;
            d.Margin = new Padding(0, 4, 4, 4);
            return d;
        }

        private ComboBox MakeCombo(string[] items)
        {
            var c = new ComboBox();
            c.Font = new Font("Segoe UI", 10F);
            c.DropDownStyle = ComboBoxStyle.DropDownList;
            c.Items.AddRange(items);
            c.Dock = DockStyle.Fill;
            c.Margin = new Padding(0, 4, 4, 4);
            return c;
        }

        private Button MakeActionButton(string text, Color color)
        {
            var b = new Button();
            b.Text = text;
            b.FlatStyle = FlatStyle.Flat;
            b.BackColor = color;
            b.ForeColor = Color.White;
            b.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            b.FlatAppearance.BorderSize = 0;
            b.Cursor = Cursors.Hand;
            return b;
        }

        /// Small button for toolbars / flow panels
        private Button MakeFlowButton(string text, Color color)
        {
            var b = new Button();
            b.Text = text;
            b.Width = 95;
            b.Height = 34;
            b.FlatStyle = FlatStyle.Flat;
            b.BackColor = color;
            b.ForeColor = Color.White;
            b.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            b.FlatAppearance.BorderSize = 0;
            b.Cursor = Cursors.Hand;
            b.Margin = new Padding(0, 0, 6, 0);
            return b;
        }

        /// Full-height button for toolbar TableLayoutPanels (fills its cell)
        private Button MakeBarButton(string text, Color color)
        {
            var b = new Button();
            b.Text = text;
            b.Dock = DockStyle.Fill;
            b.FlatStyle = FlatStyle.Flat;
            b.BackColor = color;
            b.ForeColor = Color.White;
            b.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            b.FlatAppearance.BorderSize = 0;
            b.Cursor = Cursors.Hand;
            b.Margin = new Padding(3, 0, 3, 0);
            return b;
        }

        private void AddButtonCol(DataGridView dgv, string name, string text, Color back, Color selBack, int width)
        {
            var col = new DataGridViewButtonColumn();
            col.Text = text;
            col.UseColumnTextForButtonValue = true;
            col.Name = name;
            col.Width = width;
            col.FillWeight = 1;
            col.FlatStyle = FlatStyle.Flat;
            var s = new DataGridViewCellStyle();
            s.BackColor = back;
            s.ForeColor = Color.White;
            s.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            s.SelectionBackColor = selBack;
            s.SelectionForeColor = Color.White;
            s.Alignment = DataGridViewContentAlignment.MiddleCenter;
            s.Padding = new Padding(2);
            col.DefaultCellStyle = s;
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgv.Columns.Add(col);
        }
    }
}