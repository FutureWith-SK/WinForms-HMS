using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace HMS
{
    public partial class Form1 : Form
    {
        private List<Patient> patients = new List<Patient>();
        private string dataFilePath = Path.Combine(Application.StartupPath, "patients.json");
        private string photoFolderPath = Path.Combine(Application.StartupPath, "PatientPhotos");
        private string currentPhotoPath = string.Empty;
        private int editingPatientIndex = -1;

        public Form1()
        {
            InitializeComponent();
            InitializeEPPlus();
        }

        private void InitializeEPPlus()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                Directory.CreateDirectory(photoFolderPath);
                LoadPatients();
                GenerateNewPatientID();
                RefreshDashboard();
                RefreshPatientsGrid();
                UpdateStatus("System Ready");
                lblDateTime.Text = DateTime.Now.ToString("dddd, MMMM dd, yyyy | HH:mm:ss");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing application: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus("Error loading data");
            }
        }

        private void TabControlMain_DrawItem(object sender, DrawItemEventArgs e)
        {
            TabControl tabControl = sender as TabControl;
            TabPage tabPage = tabControl.TabPages[e.Index];
            Rectangle tabRect = tabControl.GetTabRect(e.Index);

            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            Color backColor = isSelected ? Color.FromArgb(42, 92, 130) : Color.FromArgb(240, 244, 248);
            Color textColor = isSelected ? Color.White : Color.FromArgb(42, 92, 130);

            using (SolidBrush brush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(brush, tabRect);
            }

            TextRenderer.DrawText(e.Graphics, tabPage.Text, e.Font, tabRect, textColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        private void TxtQuickSearch_Enter(object sender, EventArgs e)
        {
            if (txtQuickSearch.Text == "Search patients...")
            {
                txtQuickSearch.Text = "";
                txtQuickSearch.ForeColor = Color.Black;
            }
        }

        private void TxtQuickSearch_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtQuickSearch.Text))
            {
                txtQuickSearch.Text = "Search patients...";
                txtQuickSearch.ForeColor = Color.Gray;
            }
        }

        private void BtnQuickAdd_Click(object sender, EventArgs e)
        {
            ClearForm();
            tabControlMain.SelectedTab = tabAddPatient;
        }

        #region Patient Management

        private void GenerateNewPatientID()
        {
            int maxId = 0;
            foreach (var patient in patients)
            {
                if (patient.PatientID != null && patient.PatientID.StartsWith("HOSP-") &&
                    int.TryParse(patient.PatientID.Substring(5), out int id))
                {
                    if (id > maxId) maxId = id;
                }
            }
            txtPatientID.Text = $"HOSP-{(maxId + 1):D3}";
        }

        private void CalculateAge()
        {
            if (dtpDOB.Value <= DateTime.Now)
            {
                int age = DateTime.Now.Year - dtpDOB.Value.Year;
                if (DateTime.Now < dtpDOB.Value.AddYears(age)) age--;
                txtAge.Text = age.ToString();
            }
        }

        private void DtpDOB_ValueChanged(object sender, EventArgs e)
        {
            CalculateAge();
        }

        private void BtnUploadPhoto_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Title = "Select Patient Photo";
                    ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif";
                    ofd.Multiselect = false;

                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        string fileName = $"{txtPatientID.Text}_{DateTime.Now:yyyyMMddHHmmss}{Path.GetExtension(ofd.FileName)}";
                        string destinationPath = Path.Combine(photoFolderPath, fileName);

                        if (!Directory.Exists(photoFolderPath))
                        {
                            Directory.CreateDirectory(photoFolderPath);
                        }

                        File.Copy(ofd.FileName, destinationPath, true);
                        currentPhotoPath = destinationPath;

                        if (pbPatientPhoto.Image != null)
                        {
                            pbPatientPhoto.Image.Dispose();
                        }

                        pbPatientPhoto.Image = Image.FromFile(destinationPath);
                        UpdateStatus("Photo uploaded successfully");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error uploading photo: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool ValidatePatientForm()
        {
            if (string.IsNullOrWhiteSpace(txtFullName.Text))
            {
                MessageBox.Show("Full Name is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtFullName.Focus();
                return false;
            }

            if (cmbGender.SelectedIndex == -1)
            {
                MessageBox.Show("Gender is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbGender.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtContact.Text))
            {
                MessageBox.Show("Contact Number is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtContact.Focus();
                return false;
            }

            // Updated: Accept 10 or 11 digit phone numbers
            if ((txtContact.Text.Length != 10 && txtContact.Text.Length != 11) || !txtContact.Text.All(char.IsDigit))
            {
                MessageBox.Show("Contact Number must be 10 or 11 digits.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtContact.Focus();
                return false;
            }

            // Emergency contact validation (optional field)
            if (!string.IsNullOrWhiteSpace(txtEmergencyContact.Text))
            {
                if ((txtEmergencyContact.Text.Length != 10 && txtEmergencyContact.Text.Length != 11) ||
                    !txtEmergencyContact.Text.All(char.IsDigit))
                {
                    MessageBox.Show("Emergency Contact must be 10 or 11 digits.", "Validation Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtEmergencyContact.Focus();
                    return false;
                }
            }

            if (!string.IsNullOrWhiteSpace(txtEmail.Text) && !IsValidEmail(txtEmail.Text))
            {
                MessageBox.Show("Please enter a valid email address.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtEmail.Focus();
                return false;
            }

            if (int.TryParse(txtAge.Text, out int age) && (age < 0 || age > 120))
            {
                MessageBox.Show("Age must be between 0 and 120 years.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (dtpDOB.Value > DateTime.Now)
            {
                MessageBox.Show("Date of Birth cannot be in the future.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtDoctorAssigned.Text))
            {
                MessageBox.Show("Doctor Assigned is required.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtDoctorAssigned.Focus();
                return false;
            }

            if (dtpAdmissionDate.Value > DateTime.Now)
            {
                MessageBox.Show("Admission Date cannot be in the future.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        private void BtnSavePatient_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidatePatientForm()) return;

                Patient patient = new Patient
                {
                    PatientID = txtPatientID.Text ?? string.Empty,
                    FullName = txtFullName.Text?.Trim() ?? string.Empty,
                    Age = int.TryParse(txtAge.Text, out int age) ? age : 0,
                    Gender = cmbGender.Text ?? string.Empty,
                    DateOfBirth = dtpDOB.Value,
                    ContactNumber = txtContact.Text?.Trim() ?? string.Empty,
                    Email = txtEmail.Text?.Trim() ?? string.Empty,
                    Address = txtAddress.Text?.Trim() ?? string.Empty,
                    EmergencyContact = txtEmergencyContact.Text?.Trim() ?? string.Empty,
                    BloodGroup = cmbBloodGroup.Text ?? string.Empty,
                    DoctorAssigned = txtDoctorAssigned.Text?.Trim() ?? string.Empty,
                    WardNumber = txtWardNumber.Text?.Trim() ?? string.Empty,
                    MedicalHistory = txtMedicalHistory.Text?.Trim() ?? string.Empty,
                    Allergies = txtAllergies.Text?.Trim() ?? string.Empty,
                    InsuranceProvider = txtInsuranceProvider.Text?.Trim() ?? string.Empty,
                    AdmissionDate = dtpAdmissionDate.Value,
                    PhotoPath = currentPhotoPath ?? string.Empty
                };

                if (editingPatientIndex >= 0)
                {
                    patients[editingPatientIndex] = patient;
                    editingPatientIndex = -1;
                    UpdateStatus("Patient updated successfully");
                }
                else
                {
                    patients.Add(patient);
                    UpdateStatus("Patient added successfully");
                }

                SavePatients();
                ClearForm();
                GenerateNewPatientID();
                RefreshDashboard();
                RefreshPatientsGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving patient: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus("Error saving patient");
            }
        }

        private void BtnClearForm_Click(object sender, EventArgs e)
        {
            ClearForm();
            editingPatientIndex = -1;
            UpdateStatus("Form cleared");
        }

        private void ClearForm()
        {
            txtFullName.Clear();
            txtAge.Clear();
            cmbGender.SelectedIndex = -1;
            dtpDOB.Value = DateTime.Now.AddYears(-25);
            txtContact.Clear();
            txtEmail.Clear();
            txtAddress.Clear();
            txtEmergencyContact.Clear();
            cmbBloodGroup.SelectedIndex = -1;
            txtDoctorAssigned.Clear();
            txtWardNumber.Clear();
            txtMedicalHistory.Clear();
            txtAllergies.Clear();
            txtInsuranceProvider.Clear();
            dtpAdmissionDate.Value = DateTime.Now;
            currentPhotoPath = string.Empty;

            if (pbPatientPhoto.Image != null)
            {
                pbPatientPhoto.Image.Dispose();
                pbPatientPhoto.Image = null;
            }

            GenerateNewPatientID();
        }

        #endregion

        #region Data Persistence

        private void LoadPatients()
        {
            try
            {
                if (File.Exists(dataFilePath))
                {
                    string json = File.ReadAllText(dataFilePath);
                    patients = JsonConvert.DeserializeObject<List<Patient>>(json) ?? new List<Patient>();
                    UpdateStatus($"Loaded {patients.Count} patients");
                }
                else
                {
                    patients = new List<Patient>();
                    UpdateStatus("No existing data found. Starting fresh.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading data: {ex.Message}\n\nStarting with empty dataset.", "Load Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                patients = new List<Patient>();
            }
        }

        private void SavePatients()
        {
            try
            {
                string json = JsonConvert.SerializeObject(patients, Formatting.Indented);
                File.WriteAllText(dataFilePath, json);
                UpdateStatus("Data saved successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving data: {ex.Message}", "Save Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus("Error saving data");
            }
        }

        #endregion

        #region Dashboard

        private void RefreshDashboard()
        {
            try
            {
                UpdateStatLabel(panelStatsTotal, patients.Count.ToString());
                UpdateStatLabel(panelStatsToday,
                    patients.Count(p => p.AdmissionDate.Date == DateTime.Today).ToString());
                UpdateStatLabel(panelStatsDoctors,
                    patients.Select(p => p.DoctorAssigned).Where(d => !string.IsNullOrWhiteSpace(d)).Distinct().Count().ToString());
                UpdateStatLabel(panelStatsWards,
                    patients.Select(p => p.WardNumber).Where(w => !string.IsNullOrWhiteSpace(w)).Distinct().Count().ToString());

                dgvRecentPatients.Rows.Clear();
                var recentPatients = patients.OrderByDescending(p => p.AdmissionDate).Take(10);
                foreach (var patient in recentPatients)
                {
                    dgvRecentPatients.Rows.Add(
                        patient.PatientID ?? "",
                        patient.FullName ?? "",
                        patient.Age,
                        patient.Gender ?? "",
                        patient.DoctorAssigned ?? "",
                        patient.WardNumber ?? "",
                        patient.AdmissionDate.ToString("dd/MM/yyyy")
                    );
                }

                lblDateTime.Text = DateTime.Now.ToString("dddd, MMMM dd, yyyy | HH:mm:ss");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error refreshing dashboard: {ex.Message}");
            }
        }

        private void UpdateStatLabel(Panel card, string value)
        {
            foreach (Control control in card.Controls)
            {
                if (control.Name == "lblCount")
                {
                    control.Text = value;
                    break;
                }
            }
        }

        private void TxtQuickSearch_TextChanged(object sender, EventArgs e)
        {
            if (txtQuickSearch.Text == "Search patients...") return;

            string searchText = txtQuickSearch.Text.ToLower();
            dgvRecentPatients.Rows.Clear();

            var filtered = patients.Where(p =>
                (p.PatientID != null && p.PatientID.ToLower().Contains(searchText)) ||
                (p.FullName != null && p.FullName.ToLower().Contains(searchText)) ||
                (p.ContactNumber != null && p.ContactNumber.Contains(searchText)) ||
                (p.DoctorAssigned != null && p.DoctorAssigned.ToLower().Contains(searchText))
            ).OrderByDescending(p => p.AdmissionDate).Take(10);

            foreach (var patient in filtered)
            {
                dgvRecentPatients.Rows.Add(
                    patient.PatientID ?? "",
                    patient.FullName ?? "",
                    patient.Age,
                    patient.Gender ?? "",
                    patient.DoctorAssigned ?? "",
                    patient.WardNumber ?? "",
                    patient.AdmissionDate.ToString("dd/MM/yyyy")
                );
            }
        }

        #endregion

        #region View Patients

        private void RefreshPatientsGrid()
        {
            try
            {
                dgvPatients.Rows.Clear();
                foreach (var patient in patients.OrderByDescending(p => p.AdmissionDate))
                {
                    int rowIndex = dgvPatients.Rows.Add(
                        patient.PatientID ?? "",
                        patient.FullName ?? "",
                        patient.Age,
                        patient.Gender ?? "",
                        patient.ContactNumber ?? "",
                        patient.DoctorAssigned ?? "",
                        patient.WardNumber ?? "",
                        patient.AdmissionDate.ToString("dd/MM/yyyy"),
                        "👁 View",
                        "✏ Edit",
                        "🗑 Delete"
                    );

                    // Apply button styling to each row
                    dgvPatients.Rows[rowIndex].Cells["btnView"].Style.BackColor = Color.FromArgb(59, 130, 246);
                    dgvPatients.Rows[rowIndex].Cells["btnView"].Style.ForeColor = Color.White;
                    dgvPatients.Rows[rowIndex].Cells["btnView"].Style.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

                    dgvPatients.Rows[rowIndex].Cells["btnEdit"].Style.BackColor = Color.FromArgb(245, 158, 11);
                    dgvPatients.Rows[rowIndex].Cells["btnEdit"].Style.ForeColor = Color.White;
                    dgvPatients.Rows[rowIndex].Cells["btnEdit"].Style.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

                    dgvPatients.Rows[rowIndex].Cells["btnDelete"].Style.BackColor = Color.FromArgb(239, 68, 68);
                    dgvPatients.Rows[rowIndex].Cells["btnDelete"].Style.ForeColor = Color.White;
                    dgvPatients.Rows[rowIndex].Cells["btnDelete"].Style.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
                }
                UpdateStatus($"Showing {patients.Count} patients");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error refreshing grid: {ex.Message}");
            }
        }

        private void TxtViewSearch_TextChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void CmbFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbActionFilter.Items.Clear();
            cmbActionFilter.Visible = false;

            if (cmbFilter.SelectedIndex == 1)
            {
                var doctors = patients.Select(p => p.DoctorAssigned)
                    .Where(d => !string.IsNullOrWhiteSpace(d))
                    .Distinct()
                    .ToArray();
                if (doctors.Length > 0)
                {
                    cmbActionFilter.Items.AddRange(doctors);
                    cmbActionFilter.Visible = true;
                }
            }
            else if (cmbFilter.SelectedIndex == 2)
            {
                var wards = patients.Select(p => p.WardNumber)
                    .Where(w => !string.IsNullOrWhiteSpace(w))
                    .Distinct()
                    .ToArray();
                if (wards.Length > 0)
                {
                    cmbActionFilter.Items.AddRange(wards);
                    cmbActionFilter.Visible = true;
                }
            }
            else if (cmbFilter.SelectedIndex == 3)
            {
                var bloodGroups = patients.Select(p => p.BloodGroup)
                    .Where(b => !string.IsNullOrWhiteSpace(b))
                    .Distinct()
                    .ToArray();
                if (bloodGroups.Length > 0)
                {
                    cmbActionFilter.Items.AddRange(bloodGroups);
                    cmbActionFilter.Visible = true;
                }
            }

            ApplyFilters();
        }

        private void CmbActionFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void ApplyFilters()
        {
            try
            {
                IEnumerable<Patient> filtered = patients;

                if (!string.IsNullOrWhiteSpace(txtViewSearch.Text))
                {
                    string search = txtViewSearch.Text.ToLower();
                    filtered = filtered.Where(p =>
                        (p.PatientID != null && p.PatientID.ToLower().Contains(search)) ||
                        (p.FullName != null && p.FullName.ToLower().Contains(search)) ||
                        (p.ContactNumber != null && p.ContactNumber.Contains(search)) ||
                        (p.DoctorAssigned != null && p.DoctorAssigned.ToLower().Contains(search)));
                }

                if (cmbFilter.SelectedIndex == 4)
                {
                    filtered = filtered.Where(p => p.AdmissionDate.Date == DateTime.Today);
                }
                else if (cmbActionFilter.Visible && cmbActionFilter.SelectedItem != null)
                {
                    string filterValue = cmbActionFilter.SelectedItem.ToString();
                    if (cmbFilter.SelectedIndex == 1)
                    {
                        filtered = filtered.Where(p => p.DoctorAssigned == filterValue);
                    }
                    else if (cmbFilter.SelectedIndex == 2)
                    {
                        filtered = filtered.Where(p => p.WardNumber == filterValue);
                    }
                    else if (cmbFilter.SelectedIndex == 3)
                    {
                        filtered = filtered.Where(p => p.BloodGroup == filterValue);
                    }
                }

                dgvPatients.Rows.Clear();
                foreach (var patient in filtered.OrderByDescending(p => p.AdmissionDate))
                {
                    dgvPatients.Rows.Add(
                        patient.PatientID ?? "",
                        patient.FullName ?? "",
                        patient.Age,
                        patient.Gender ?? "",
                        patient.ContactNumber ?? "",
                        patient.DoctorAssigned ?? "",
                        patient.WardNumber ?? "",
                        patient.AdmissionDate.ToString("dd/MM/yyyy"),
                        "View",
                        "Edit",
                        "Delete"
                    );
                }
                UpdateStatus($"Showing {filtered.Count()} patients");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error applying filters: {ex.Message}");
            }
        }

        private void DgvPatients_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0) return;

                string patientId = dgvPatients.Rows[e.RowIndex].Cells["PatientID"].Value?.ToString();
                if (string.IsNullOrEmpty(patientId)) return;

                Patient patient = patients.FirstOrDefault(p => p.PatientID == patientId);
                if (patient == null) return;

                if (e.ColumnIndex == dgvPatients.Columns["btnView"].Index)
                {
                    ViewPatientDetails(patient);
                }
                else if (e.ColumnIndex == dgvPatients.Columns["btnEdit"].Index)
                {
                    EditPatient(patient);
                }
                else if (e.ColumnIndex == dgvPatients.Columns["btnDelete"].Index)
                {
                    DeletePatient(patient);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error processing action: {ex.Message}");
            }
        }

        private void ViewPatientDetails(Patient patient)
        {
            string details = $"Patient ID: {patient.PatientID}\n" +
                             $"Full Name: {patient.FullName}\n" +
                             $"Age: {patient.Age}\n" +
                             $"Gender: {patient.Gender}\n" +
                             $"Date of Birth: {patient.DateOfBirth:dd/MM/yyyy}\n" +
                             $"Contact: {patient.ContactNumber}\n" +
                             $"Email: {patient.Email}\n" +
                             $"Address: {patient.Address}\n" +
                             $"Emergency Contact: {patient.EmergencyContact}\n" +
                             $"Blood Group: {patient.BloodGroup}\n" +
                             $"Doctor: {patient.DoctorAssigned}\n" +
                             $"Ward: {patient.WardNumber}\n" +
                             $"Medical History: {patient.MedicalHistory}\n" +
                             $"Allergies: {patient.Allergies}\n" +
                             $"Insurance: {patient.InsuranceProvider}\n" +
                             $"Admission Date: {patient.AdmissionDate:dd/MM/yyyy}";

            MessageBox.Show(details, "Patient Details", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void EditPatient(Patient patient)
        {
            editingPatientIndex = patients.IndexOf(patient);
            txtPatientID.Text = patient.PatientID ?? string.Empty;
            txtFullName.Text = patient.FullName ?? string.Empty;
            cmbGender.Text = patient.Gender ?? string.Empty;
            dtpDOB.Value = patient.DateOfBirth;
            txtContact.Text = patient.ContactNumber ?? string.Empty;
            txtEmail.Text = patient.Email ?? string.Empty;
            txtAddress.Text = patient.Address ?? string.Empty;
            txtEmergencyContact.Text = patient.EmergencyContact ?? string.Empty;
            cmbBloodGroup.Text = patient.BloodGroup ?? string.Empty;
            txtDoctorAssigned.Text = patient.DoctorAssigned ?? string.Empty;
            txtWardNumber.Text = patient.WardNumber ?? string.Empty;
            txtMedicalHistory.Text = patient.MedicalHistory ?? string.Empty;
            txtAllergies.Text = patient.Allergies ?? string.Empty;
            txtInsuranceProvider.Text = patient.InsuranceProvider ?? string.Empty;
            dtpAdmissionDate.Value = patient.AdmissionDate;
            currentPhotoPath = patient.PhotoPath ?? string.Empty;

            if (pbPatientPhoto.Image != null)
            {
                pbPatientPhoto.Image.Dispose();
                pbPatientPhoto.Image = null;
            }

            if (!string.IsNullOrEmpty(patient.PhotoPath) && File.Exists(patient.PhotoPath))
            {
                pbPatientPhoto.Image = Image.FromFile(patient.PhotoPath);
            }

            tabControlMain.SelectedTab = tabAddPatient;
            UpdateStatus($"Editing patient: {patient.FullName}");
        }

        private void DeletePatient(Patient patient)
        {
            DialogResult result = MessageBox.Show(
                $"Are you sure you want to delete {patient.FullName} (ID: {patient.PatientID})?",
                "Confirm Delete",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                patients.Remove(patient);
                SavePatients();
                RefreshDashboard();
                RefreshPatientsGrid();
                UpdateStatus($"Patient {patient.FullName} deleted successfully");
            }
        }

        #endregion

        #region Export Functions

        private void BtnExportJSON_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "JSON files|*.json";
                    sfd.FileName = $"PatientBackup_{DateTime.Now:yyyyMMdd_HHmmss}.json";
                    sfd.Title = "Export Patients to JSON";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        string json = JsonConvert.SerializeObject(patients, Formatting.Indented);
                        File.WriteAllText(sfd.FileName, json);
                        UpdateStatus($"Exported {patients.Count} patients to JSON");
                        MessageBox.Show($"Successfully exported {patients.Count} patients to:\n{sfd.FileName}",
                            "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting JSON: {ex.Message}", "Export Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportExcel_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Excel files|*.xlsx";
                    sfd.FileName = $"PatientRecords_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                    sfd.Title = "Export Patients to Excel";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        ExportToExcel(sfd.FileName, patients);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting Excel: {ex.Message}", "Export Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportToExcel(string filePath, List<Patient> patientList, string reportTitle = null)
        {
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Patients");

                    int startRow = 1;

                    if (!string.IsNullOrEmpty(reportTitle))
                    {
                        worksheet.Cells["A1"].Value = reportTitle;
                        worksheet.Cells["A1:P1"].Merge = true;
                        worksheet.Cells["A1"].Style.Font.Size = 16;
                        worksheet.Cells["A1"].Style.Font.Bold = true;
                        worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells["A2"].Value = $"Generated: {DateTime.Now:dd/MM/yyyy HH:mm}";
                        worksheet.Cells["A2:P2"].Merge = true;
                        worksheet.Cells["A2"].Style.Font.Size = 10;
                        startRow = 4;
                    }

                    string[] headers = { "Patient ID", "Full Name", "Age", "Gender", "Date of Birth",
                        "Contact", "Email", "Address", "Emergency Contact", "Blood Group",
                        "Doctor", "Ward", "Medical History", "Allergies", "Insurance",
                        "Admission Date" };

                    for (int i = 0; i < headers.Length; i++)
                    {
                        var cell = worksheet.Cells[startRow, i + 1];
                        cell.Value = headers[i];
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(42, 92, 130));
                        cell.Style.Font.Color.SetColor(Color.White);
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    for (int i = 0; i < patientList.Count; i++)
                    {
                        var patient = patientList[i];
                        int row = startRow + 1 + i;

                        worksheet.Cells[row, 1].Value = patient.PatientID ?? "";
                        worksheet.Cells[row, 2].Value = patient.FullName ?? "";
                        worksheet.Cells[row, 3].Value = patient.Age;
                        worksheet.Cells[row, 4].Value = patient.Gender ?? "";
                        worksheet.Cells[row, 5].Value = patient.DateOfBirth.ToString("dd/MM/yyyy");
                        worksheet.Cells[row, 6].Value = patient.ContactNumber ?? "";
                        worksheet.Cells[row, 7].Value = patient.Email ?? "";
                        worksheet.Cells[row, 8].Value = patient.Address ?? "";
                        worksheet.Cells[row, 9].Value = patient.EmergencyContact ?? "";
                        worksheet.Cells[row, 10].Value = patient.BloodGroup ?? "";
                        worksheet.Cells[row, 11].Value = patient.DoctorAssigned ?? "";
                        worksheet.Cells[row, 12].Value = patient.WardNumber ?? "";
                        worksheet.Cells[row, 13].Value = patient.MedicalHistory ?? "";
                        worksheet.Cells[row, 14].Value = patient.Allergies ?? "";
                        worksheet.Cells[row, 15].Value = patient.InsuranceProvider ?? "";
                        worksheet.Cells[row, 16].Value = patient.AdmissionDate.ToString("dd/MM/yyyy");

                        if (i % 2 == 0)
                        {
                            for (int j = 1; j <= 16; j++)
                            {
                                worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(
                                    Color.FromArgb(240, 244, 248));
                            }
                        }

                        for (int j = 1; j <= 16; j++)
                        {
                            worksheet.Cells[row, j].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }
                    }

                    worksheet.Cells.AutoFitColumns();
                    worksheet.Column(8).Width = 30;
                    worksheet.Column(13).Width = 30;
                    worksheet.Column(14).Width = 25;

                    package.SaveAs(new FileInfo(filePath));
                }

                UpdateStatus("Exported to Excel successfully");
                MessageBox.Show($"Successfully exported to:\n{filePath}", "Export Complete",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                throw new Exception($"Excel export failed: {ex.Message}", ex);
            }
        }

        private void BtnExportText_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Text files|*.txt";
                    sfd.FileName = $"PatientRecords_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                    sfd.Title = "Export Patients to Text";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        using (StreamWriter writer = new StreamWriter(sfd.FileName))
                        {
                            writer.WriteLine("═══════════════════════════════════════════════════════");
                            writer.WriteLine("           MEDICARE HOSPITAL - PATIENT RECORDS         ");
                            writer.WriteLine("═══════════════════════════════════════════════════════");
                            writer.WriteLine($"Generated: {DateTime.Now:dd/MM/yyyy HH:mm}");
                            writer.WriteLine($"Total Patients: {patients.Count}");
                            writer.WriteLine("═══════════════════════════════════════════════════════");
                            writer.WriteLine();

                            foreach (var patient in patients)
                            {
                                writer.WriteLine($"Patient ID: {patient.PatientID}");
                                writer.WriteLine($"Name: {patient.FullName}");
                                writer.WriteLine($"Age: {patient.Age} | Gender: {patient.Gender}");
                                writer.WriteLine($"Contact: {patient.ContactNumber}");
                                writer.WriteLine($"Doctor: {patient.DoctorAssigned} | Ward: {patient.WardNumber}");
                                writer.WriteLine($"Admission: {patient.AdmissionDate:dd/MM/yyyy}");
                                writer.WriteLine("───────────────────────────────────────────────────────");
                            }
                        }

                        UpdateStatus("Exported to Text successfully");
                        MessageBox.Show($"Successfully exported to:\n{sfd.FileName}", "Export Complete",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting Text: {ex.Message}", "Export Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Reports

        private void BtnDailyReport_Click(object sender, EventArgs e)
        {
            try
            {
                var todayPatients = patients.Where(p => p.AdmissionDate.Date == DateTime.Today).ToList();
                GenerateReportText("Daily Admissions Report", todayPatients);
                UpdateStatus($"Daily report generated - {todayPatients.Count} patients today");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error generating daily report: {ex.Message}");
            }
        }

        private void BtnWardReport_Click(object sender, EventArgs e)
        {
            try
            {
                var wardGroups = patients.GroupBy(p => p.WardNumber)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .OrderBy(g => g.Key);

                rtbReport.Clear();
                rtbReport.AppendText("═══════════════════════════════════════════════════════\n");
                rtbReport.AppendText("              WARD OCCUPANCY REPORT                    \n");
                rtbReport.AppendText("═══════════════════════════════════════════════════════\n");
                rtbReport.AppendText($"Generated: {DateTime.Now:dd/MM/yyyy HH:mm}\n\n");

                foreach (var group in wardGroups)
                {
                    rtbReport.AppendText($"\nWARD: {group.Key}\n");
                    rtbReport.AppendText($"Occupancy: {group.Count()} patients\n");
                    rtbReport.AppendText("───────────────────────────────────────────────────────\n");
                    foreach (var patient in group)
                    {
                        rtbReport.AppendText($"  • {patient.FullName} (ID: {patient.PatientID}) - {patient.DoctorAssigned}\n");
                    }
                }

                rtbReport.AppendText($"\n\nTotal Wards: {wardGroups.Count()}");
                rtbReport.AppendText($"\nTotal Patients: {patients.Count}");

                UpdateStatus("Ward occupancy report generated");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error generating ward report: {ex.Message}");
            }
        }

        private void BtnDoctorReport_Click(object sender, EventArgs e)
        {
            try
            {
                var doctorGroups = patients.GroupBy(p => p.DoctorAssigned)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .OrderBy(g => g.Key);

                rtbReport.Clear();
                rtbReport.AppendText("═══════════════════════════════════════════════════════\n");
                rtbReport.AppendText("           DOCTOR-WISE PATIENT REPORT                  \n");
                rtbReport.AppendText("═══════════════════════════════════════════════════════\n");
                rtbReport.AppendText($"Generated: {DateTime.Now:dd/MM/yyyy HH:mm}\n\n");

                foreach (var group in doctorGroups)
                {
                    rtbReport.AppendText($"\nDOCTOR: {group.Key}\n");
                    rtbReport.AppendText($"Total Patients: {group.Count()}\n");
                    rtbReport.AppendText("───────────────────────────────────────────────────────\n");
                    foreach (var patient in group.OrderBy(p => p.AdmissionDate))
                    {
                        rtbReport.AppendText($"  • {patient.FullName} (ID: {patient.PatientID}) - " +
                                           $"Ward: {patient.WardNumber} | {patient.AdmissionDate:dd/MM/yyyy}\n");
                    }
                }

                rtbReport.AppendText($"\n\nTotal Doctors: {doctorGroups.Count()}");
                rtbReport.AppendText($"\nTotal Patients: {patients.Count}");

                UpdateStatus("Doctor-wise report generated");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error generating doctor report: {ex.Message}");
            }
        }

        private void GenerateReportText(string title, List<Patient> patientList)
        {
            rtbReport.Clear();
            rtbReport.AppendText("═══════════════════════════════════════════════════════\n");
            rtbReport.AppendText($"           {title.ToUpper()}            \n");
            rtbReport.AppendText("═══════════════════════════════════════════════════════\n");
            rtbReport.AppendText($"Generated: {DateTime.Now:dd/MM/yyyy HH:mm}\n");
            rtbReport.AppendText($"Total Patients: {patientList.Count}\n\n");

            if (patientList.Count > 0)
            {
                foreach (var patient in patientList.OrderBy(p => p.AdmissionDate))
                {
                    rtbReport.AppendText($"Patient: {patient.FullName}\n");
                    rtbReport.AppendText($"ID: {patient.PatientID} | Age: {patient.Age} | Gender: {patient.Gender}\n");
                    rtbReport.AppendText($"Doctor: {patient.DoctorAssigned} | Ward: {patient.WardNumber}\n");
                    rtbReport.AppendText($"Admission: {patient.AdmissionDate:dd/MM/yyyy}\n");
                    rtbReport.AppendText("───────────────────────────────────────────────────────\n");
                }
            }
            else
            {
                rtbReport.AppendText("No patients found for this report.\n");
            }
        }

        private void BtnPrintReport_Click(object sender, EventArgs e)
        {
            try
            {
                if (rtbReport.Text.Length == 0)
                {
                    MessageBox.Show("Please generate a report first.", "No Report",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (PrintDialog printDialog = new PrintDialog())
                {
                    System.Drawing.Printing.PrintDocument printDoc = new System.Drawing.Printing.PrintDocument();
                    printDoc.PrintPage += (s, ev) =>
                    {
                        ev.Graphics.DrawString(rtbReport.Text, new Font("Consolas", 10),
                            Brushes.Black, new PointF(50, 50));
                    };

                    printDialog.Document = printDoc;
                    if (printDialog.ShowDialog() == DialogResult.OK)
                    {
                        printDoc.Print();
                        UpdateStatus("Report sent to printer");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error printing: {ex.Message}", "Print Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportReportExcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (patients.Count == 0)
                {
                    MessageBox.Show("No data to export.", "Empty Data",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Excel files|*.xlsx";
                    sfd.FileName = $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                    sfd.Title = "Export Report to Excel";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        ExportToExcel(sfd.FileName, patients, "Hospital Patient Report");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting report: {ex.Message}", "Export Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Utility Methods

        private void UpdateStatus(string message)
        {
            if (lblStatus != null)
            {
                lblStatus.Text = message;
            }
        }

        #endregion
    }

    public class Patient
    {
        public string PatientID { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Gender { get; set; } = string.Empty;
        public DateTime DateOfBirth { get; set; }
        public string ContactNumber { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Address { get; set; } = string.Empty;
        public string EmergencyContact { get; set; } = string.Empty;
        public string BloodGroup { get; set; } = string.Empty;
        public string DoctorAssigned { get; set; } = string.Empty;
        public string WardNumber { get; set; } = string.Empty;
        public string MedicalHistory { get; set; } = string.Empty;
        public string Allergies { get; set; } = string.Empty;
        public string InsuranceProvider { get; set; } = string.Empty;
        public DateTime AdmissionDate { get; set; }
        public string PhotoPath { get; set; } = string.Empty;
    }
}