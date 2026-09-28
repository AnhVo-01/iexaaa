using DBConnect;
using DevExpress.Spreadsheet;
using Dto;
using IEXAAA.Dto;
using Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IEXAAA
{
    public partial class Main : Form
    {
        private readonly IEDBContext _DbConnect;
        private readonly HSNSContext _HSNSContext;

        private volatile bool isCancelled = false;
        private GithubAsset latestInstaller;
        private CancellationTokenSource downloadCts;

        public Main()
        {
            InitializeComponent();
            _DbConnect = new IEDBContext();
            _HSNSContext = new HSNSContext();

            appVersion.Text = "Current version: " + Assembly.GetExecutingAssembly().GetName().Version.ToString();

            btn_import.Enabled = false;
            btn_cancel.Visible = false;
            btn_install.Visible = false;
            checkLoad.Visible = false;
            panel1.Visible = false;
            panel2.Visible = false;
            lblProgress.Visible = false;
            progressInstall.Visible = false;
        }

        private void Btn_Select(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "Select file Excel"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                file_url.Text = filePath;
            }
        }

        private async void Btn_Import(object sender, EventArgs e)
        {
            listBoxLog.Items.Clear();
            progressBar.Value = 0;
            string filePath = file_url.Text;

            if (btn_import.Text == "Finish")
            {
                //Application.Exit();
                file_url.Text = null;
                btn_import.Text = "Import";
                btn_select.Enabled = true;
                return;
            }

            if (!File.Exists(filePath))
            {
                MessageBox.Show("Error reading file!");
                return;
            }

            if (!TryGetSelectedCongTyId(out int congTyId))
            {
                return;
            }

            try
            {
                btn_cancel.Visible = true;
                btn_import.Enabled = false;
                btn_select.Enabled = false;
                await Task.Run(() => ProcessFile(filePath, congTyId));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Main_Load(object sender, EventArgs e)
        {
            tabControl.SelectedIndexChanged += tabControl_SelectedIndexChanged;
            LoadCompanies();
            LoadColumnMappingFromJson();
        }

        private void LoadCompanies()
        {
            cbCompany.DataSource = _HSNSContext.CongTy.Where(x => x.FlagDel == 0).ToList();
            cbCompany.DisplayMember = "MaCongTy";
            cbCompany.ValueMember = "Id";
            cbCompany.SelectedIndex = -1;
        }

        private int? GetSelectedCongTyId()
        {
            return cbCompany.SelectedValue as int?;
        }

        private bool TryGetSelectedCongTyId(out int congTyId)
        {
            var selected = GetSelectedCongTyId();
            if (selected == null)
            {
                MessageBox.Show("Please select a Company first.", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                congTyId = 0;
                return false;
            }

            congTyId = selected.Value;
            return true;
        }

        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Kiểm tra nếu tab hiện tại là tab ánh xạ
            var selectedTab = tabControl.SelectedTab;
            if (selectedTab != null && selectedTab.Name == "tabPage2")
            {
                LoadColumnMappingFromJson();
            }
        }

        private void InputFilePath(object sender, EventArgs e)
        {
            file_url.Text = file_url.Text.Trim('"');
            btn_import.Enabled = !string.IsNullOrWhiteSpace(file_url.Text);
        }

        private async void ProcessFile(string filePath, int congTyId)
        {
            var currentTime = DateTime.Now;
            var workbook = new Workbook();
            workbook.LoadDocument(filePath, DocumentFormat.Xlsx);

            var worksheet = workbook.Worksheets[0];
            var cells = worksheet.Cells;

            var usedRange = worksheet.GetUsedRange();
            int startRow = usedRange.TopRowIndex;
            int endRow = usedRange.BottomRowIndex;

            // Xác định tổng số dòng hợp lệ
            int totalRows = 0;
            for (int i = startRow + 1; i <= endRow; i++)
            {
                var columnB = cells[i, 1].Value?.ToString();
                var columnC = cells[i, 2].Value?.ToString();

                if (!string.IsNullOrWhiteSpace(columnB) && !string.IsNullOrWhiteSpace(columnC))
                {
                    totalRows++;
                }
            }

            int processedRows = 0;
            int successCount = 0;
            int errorCount = 0;

            var thuocTinhList = await _DbConnect.DMThuocTinh
                .Where(x => x.FlagDel == 0 && x.CongTyId == congTyId)
                .ToListAsync();

            for (int i = startRow + 1; i <= endRow; i++)
            {
                if (isCancelled)
                {
                    Invoke(() =>
                    {
                        listBoxLog.Items.Add("Process cancelled by user.");
                        listBoxLog.TopIndex = listBoxLog.Items.Count - 1;
                    });
                    break;
                }

                string columnB = cells[i, 1].Value?.ToString();
                string columnC = cells[i, 2].Value?.ToString();
                if (string.IsNullOrWhiteSpace(columnB) || string.IsNullOrWhiteSpace(columnC))
                {
                    continue;
                }

                int currentIndex = processedRows;
                Invoke(() =>
                {
                    listBoxLog.Items.Add($"Importing product {columnB}...");
                });

                try
                {
                    // Các cột cố định
                    string maSPNB = columnB.Trim();
                    string maSPKH = columnC.Trim();
                    string tenSanPham = "";

                    // Load từ JSON
                    var columnMappings = LoadMappingSettings().Mappings;
                    var thuocTinhValues = new Dictionary<string, string>();

                    foreach (var kvp in columnMappings)
                    {
                        string propertyName = kvp.PropertyCode;
                        string columnIndex = kvp.ExcelCol;

                        var cellValue = cells[columnIndex + (i + 1)].Value?.ToString();
                        thuocTinhValues[propertyName] = cellValue;
                    }

                    // Gọi hàm xử lý
                    var newDict = thuocTinhValues.Where(kvp => kvp.Key != "tenSanPham").ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                    await UpsertSanPhamAsync(maSPNB, maSPKH, thuocTinhValues["tenSanPham"], newDict, thuocTinhList, currentTime, congTyId);

                    successCount++;
                    processedRows++;
                    int percent = (int)((double)processedRows / totalRows * 100);

                    Invoke(() =>
                    {
                        listBoxLog.Items[currentIndex] = $"Importing product {columnB}...Done";
                        listBoxLog.TopIndex = listBoxLog.Items.Count - 1;
                        progressBar.Value = percent;
                    });
                }
                catch (Exception ex)
                {
                    errorCount++;
                    processedRows++;
                    int percent = (int)((double)processedRows / totalRows * 100);

                    Invoke(() =>
                    {
                        listBoxLog.Items[currentIndex] = $"Importing product {columnB}...Error ({ex.Message})";
                        listBoxLog.TopIndex = listBoxLog.Items.Count - 1;
                        progressBar.Value = percent;
                    });
                }
            }

            if (!isCancelled)
            {
                Invoke(() =>
                {
                    listBoxLog.Items.Add($"Finished!");
                    listBoxLog.Items.Add($"✔️ Success: {successCount} products");
                    listBoxLog.Items.Add($"❌ Error: {errorCount} products");
                    listBoxLog.TopIndex = listBoxLog.Items.Count - 1;

                    btn_import.Text = "Finish";
                    btn_import.Enabled = true;
                    btn_cancel.Enabled = false;

                    //MessageBox.Show($"Import complete!\n\n✔️ Success: {successCount}\n❌ Error: {errorCount}", "Import Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                });
            }

        }

        // Hàm Invoke an toàn khi đang chạy background
        private void Invoke(Action action)
        {
            if (this.InvokeRequired)
                this.Invoke(new MethodInvoker(action));
            else
                action();
        }

        public async Task<int> UpsertSanPhamAsync(string maSPNB, string maSPKH, string tenSanPham, Dictionary<string, string> thuocTinhValues, List<DMThuocTinh> thuocTinhs, DateTime currentTime, int congTyId)
        {
            using(var transaction = _DbConnect.Database.BeginTransaction())
            {
                try
                {
                    var existing = await _DbConnect.DMSanPham
                        .FirstOrDefaultAsync(x => x.FlagDel == 0 && x.MaSPNB.Trim() == maSPNB);

                    int sanPhamId;

                    if (existing == null)
                    {
                        var newSP = new DMSanPham
                        {
                            CongTyId = congTyId,
                            MaSPNB = maSPNB,
                            MaSPKH = maSPKH,
                            TenSanPham = tenSanPham,
                            ThuocTinhSP = SerializeThuocTinh(new Dictionary<string, string>(), thuocTinhValues, thuocTinhs),
                            HoatDong = 1,
                            FlagDel = 0,
                            CreatedDate = currentTime,
                            UpdatedDate = currentTime
                        };

                        _DbConnect.DMSanPham.Add(newSP);
                        await _DbConnect.SaveChangesAsync(); // Lưu để lấy Id
                        sanPhamId = newSP.Id;
                    }
                    else
                    {
                        var currentThuocTinh = DeserializeThuocTinh(existing.ThuocTinhSP);

                        existing.TenSanPham = tenSanPham;
                        existing.ThuocTinhSP = SerializeThuocTinh(currentThuocTinh, thuocTinhValues, thuocTinhs);
                        existing.UpdatedDate = currentTime;
                        sanPhamId = existing.Id;
                    }

                    await _DbConnect.SaveChangesAsync();

                    transaction.Commit(); // Chỉ commit khi mọi thứ xong xuôi

                    return sanPhamId;
                }
                catch (Exception ex)
                {
                    transaction.Rollback(); // Nếu lỗi, rollback toàn bộ
                    throw new Exception(ex.Message, ex);
                }
            }
        }

        private Dictionary<string, string> DeserializeThuocTinh(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return new Dictionary<string, string>();

            try
            {
                return JsonSerializer.Deserialize<Dictionary<string, string>>(json) ?? new Dictionary<string, string>();
            }
            catch (JsonException)
            {
                return new Dictionary<string, string>();
            }
        }

        private string SerializeThuocTinh(Dictionary<string, string> currentValues, Dictionary<string, string> newValues, List<DMThuocTinh> thuocTinhs)
        {
            foreach (var kvp in newValues)
            {
                if (string.IsNullOrWhiteSpace(kvp.Value))
                    continue;

                var maThuocTinh = thuocTinhs.FirstOrDefault(x => x.MaThuocTinh.Equals(kvp.Key, StringComparison.OrdinalIgnoreCase))?.MaThuocTinh;

                if (maThuocTinh == null)
                    continue;

                currentValues[maThuocTinh] = kvp.Value;
            }

            return JsonSerializer.Serialize(currentValues);
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            isCancelled = true;
            btn_cancel.Enabled = false;
            btn_import.Enabled = true;
            btn_import.Text = "Finish";
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string filePath = Path.Combine(Application.StartupPath, "ImportTemplate.xlsx");
            if (File.Exists(filePath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Cannot find template.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private MappingSettings LoadMappingSettings()
        {
            string jsonPath = Path.Combine(Application.StartupPath, "column_mapping.json");

            if (!File.Exists(jsonPath))
                return new MappingSettings();

            string json = File.ReadAllText(jsonPath);

            try
            {
                var settings = JsonSerializer.Deserialize<MappingSettings>(json);
                if (settings != null)
                    return settings;
            }
            catch (JsonException)
            {
                // File cũ chỉ chứa mảng ColumnMapping, rơi xuống nhánh bên dưới để đọc theo định dạng cũ
            }

            try
            {
                var mappings = JsonSerializer.Deserialize<List<ColumnMapping>>(json);
                return new MappingSettings { Mappings = mappings ?? new List<ColumnMapping>() };
            }
            catch (JsonException)
            {
                return new MappingSettings();
            }
        }

        private void LoadColumnMappingFromJson()
        {
            var settings = LoadMappingSettings();

            dgvMapping.Rows.Clear(); // clear old data

            foreach (var mapping in settings.Mappings)
            {
                dgvMapping.Rows.Add(mapping.Property, mapping.PropertyCode, mapping.ExcelCol);
            }

            if (settings.CongTyId.HasValue)
            {
                cbCompany.SelectedValue = settings.CongTyId.Value;
            }
        }

        private void btn_load_Click(object sender, EventArgs e)
        {
            if (!TryGetSelectedCongTyId(out int congTyId))
            {
                return;
            }

            var thuocTinhList = _DbConnect.DMThuocTinh
                .Where(x => x.CongTyId == congTyId && x.FlagDel == 0)
                .ToList();

            foreach (var prop in thuocTinhList)
            {
                string property = prop.TenThuocTinh;
                string propertyCode = prop.MaThuocTinh;

                // Tìm dòng có propertyCode tương ứng
                bool updated = false;
                foreach (DataGridViewRow row in dgvMapping.Rows)
                {
                    if (row.IsNewRow) continue;

                    string existingCode = row.Cells["property"].Value?.ToString();
                    if (existingCode == propertyCode)
                    {
                        // Nếu đã tồn tại, cập nhật lại property (giữ excelCol cũ)
                        row.Cells["property"].Value = propertyCode;
                        updated = true;
                        break;
                    }
                }

                // Nếu chưa có, thêm dòng mới
                if (!updated)
                {
                    dgvMapping.Rows.Add(property, propertyCode, null); // excelCol để user chọn
                }
            }
        }

        private void saveSetting(object sender, EventArgs e)
        {
            if (!TryGetSelectedCongTyId(out int congTyId))
            {
                return;
            }

            List<ColumnMapping> mappings = new List<ColumnMapping>();

            foreach (DataGridViewRow row in dgvMapping.Rows)
            {
                if (row.IsNewRow) continue; // Bỏ dòng trống cuối cùng

                var mapping = new ColumnMapping
                {
                    Property = row.Cells["propertyName"].Value?.ToString(),
                    PropertyCode = row.Cells["property"].Value?.ToString(),
                    ExcelCol = row.Cells["excelCol"].Value?.ToString()
                };

                mappings.Add(mapping);
            }

            var settings = new MappingSettings { CongTyId = congTyId, Mappings = mappings };

            // Chuyển đổi thành JSON
            string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });

            // Đường dẫn mặc định (AppDomain.CurrentDomain.BaseDirectory) sẽ lưu file ngay cạnh file .exe
            string outputPath = Path.Combine(Application.StartupPath, "column_mapping.json");

            try
            {
                File.WriteAllText(outputPath, json);
                MessageBox.Show("Save mapping successful:\n" + outputPath, "Notification", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error when saving the file:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void CheckUpdate(object sender, EventArgs e)
        {
            checkLoad.Visible = true;

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("iexaaa", "1.0"));

                var response = await client.GetAsync("https://api.github.com/repos/AnhVo-01/iexaaa/releases/latest");

                if (!response.IsSuccessStatusCode)
                {
                    MessageBox.Show("Không thể kiểm tra cập nhật.");
                    return;
                }

                var json = await response.Content.ReadAsStringAsync();
                var release = JsonSerializer.Deserialize<GithubRelease>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                if (release != null)
                {
                    Version latest = ParseVersion(release.TagName);
                    Version current = Assembly.GetExecutingAssembly().GetName().Version;

                    if (latest > current)
                    {
                        latestInstaller = release.Assets.FirstOrDefault(a =>
                            a.Name.EndsWith(".exe") || a.Name.EndsWith(".msi"));

                        if (latestInstaller != null)
                        {
                            btn_install.Visible = true;
                            checkLoad.Visible = false;
                            panel2.Visible = true;
                            label4.Text = "New version: " + latest;
                        }
                    }
                    else
                    {
                        checkLoad.Visible = false;
                        panel1.Visible = true;
                    }
                }
            }
        }

        private Version ParseVersion(string tag)
        {
            // Giả sử tag dạng "v1.2.3" hoặc "1.2.3"
            var clean = tag.StartsWith("v") ? tag.Substring(1) : tag;
            return new Version(clean);
        }

        private async void btn_install_Click(object sender, EventArgs e)
        {
            progressInstall.Value = 0;
            progressInstall.Visible = true;
            btn_install.Enabled = false;

            if (btn_install.Text == "Cancel") {
                downloadCts?.Cancel();
                btn_install.Text = "Install";
                btn_install.Enabled = true;
                return;
            }

            if (latestInstaller == null)
            {
                MessageBox.Show("Installation file not found.");
                return;
            }

            string tempPath = Path.Combine(Path.GetTempPath(), latestInstaller.Name);
            downloadCts = new CancellationTokenSource();
            var token = downloadCts.Token;

            try
            {
                using (var client = new HttpClient())
                using (var response = await client.GetAsync(latestInstaller.BrowserDownloadUrl, HttpCompletionOption.ResponseHeadersRead, token))
                {
                    lblProgress.Visible = true;
                    btn_install.Text = "Cancel";
                    btn_install.Enabled = true;
                    response.EnsureSuccessStatusCode();

                    var totalBytes = response.Content.Headers.ContentLength ?? -1L;
                    var canReportProgress = totalBytes != -1;

                    using (var contentStream = await response.Content.ReadAsStreamAsync())
                    using (var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        var buffer = new byte[8192];
                        long totalRead = 0;
                        int read;

                        while ((read = await contentStream.ReadAsync(buffer, 0, buffer.Length, token)) > 0)
                        {
                            await fileStream.WriteAsync(buffer, 0, read, token);
                            totalRead += read;

                            if (canReportProgress)
                            {
                                int percent = (int)(totalRead * 100 / totalBytes);
                                progressInstall.Value = Math.Min(percent, 100);

                                lblProgress.Text = $"Downloading...{FormatBytes(totalRead)} / {FormatBytes(totalBytes)} ({percent}%)";
                            }
                            else
                            {
                                lblProgress.Text = $"Downloaded: {FormatBytes(totalRead)}";
                            }
                        }
                    }
                }

                // Nếu tải thành công, chạy file
                Process.Start(new ProcessStartInfo
                {
                    FileName = tempPath,
                    UseShellExecute = true
                });

                Application.Exit();
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Download was cancelled.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Download error: {ex.Message}");
            }
            finally
            {
                btn_install.Text = "Install";
                lblProgress.Visible = false;
                progressInstall.Visible = false;
            }
        }

        private string FormatBytes(long bytes)
        {
            if (bytes >= 1073741824)
                return $"{bytes / 1073741824.0:F2} GB";
            if (bytes >= 1048576)
                return $"{bytes / 1048576.0:F2} MB";
            if (bytes >= 1024)
                return $"{bytes / 1024.0:F2} KB";
            return $"{bytes} B";
        }
    }
}
