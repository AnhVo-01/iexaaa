using DBConnect;
using DevExpress.Spreadsheet;
using Model;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IEXAAA
{
    public partial class Main : Form
    {
        private readonly IEDBContext _DbConnect;
        private volatile bool isCancelled = false;

        public Main()
        {
            InitializeComponent();
            _DbConnect = new IEDBContext();

            btn_import.Enabled = false;
            btn_cancel.Visible = false;
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
                Application.Exit();
            }

            if (!File.Exists(filePath))
            {
                MessageBox.Show("Error reading file!");
                return;
            }

            try
            {
                btn_cancel.Visible = true;
                btn_import.Enabled = false;
                btn_select.Enabled = false;
                await Task.Run(() => ProcessFile(filePath));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Main_Load(object sender, EventArgs e)
        {
        }

        private void InputFilePath(object sender, EventArgs e)
        {
            file_url.Text = file_url.Text.Trim('"');
            btn_import.Enabled = !string.IsNullOrWhiteSpace(file_url.Text);
        }

        private async void ProcessFile(string filePath)
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
            var thuocTinhList = await _DbConnect.DMThuocTinh
                .Where(x => x.FlagDel == 0 && x.CongTyId == 1)
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
                    string maSPNB = columnB.Trim();
                    string maSPKH = columnC.Trim();
                    string bagType = cells[i, 3].Value?.ToString();
                    string tenSanPham = cells[i, 4].Value?.ToString();
                    string dimension = cells[i, 5].Value?.ToString();
                    string thickness = cells[i, 6].Value?.ToString();
                    string bagWeight = cells[i, 7].Value?.ToString();
                    string pcs = cells[i, 8].Value?.ToString();
                    string cartonRolls = cells[i, 9].Value?.ToString();
                    string cartonWeight = cells[i, 10].Value?.ToString();

                    var thuocTinhValues = new Dictionary<string, string>
                    {
                        { "bagType", bagType },
                        { "dimension", dimension },
                        { "thickness", thickness },
                        { "bagWeight", bagWeight },
                        { "pcs", pcs },
                        { "cartonRolls", cartonRolls },
                        { "cartonWeight", cartonWeight }
                    };

                    await UpsertSanPhamAsync(maSPNB, maSPKH, tenSanPham, thuocTinhValues, thuocTinhList, currentTime);

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
                    listBoxLog.Items.Add("Finished!");
                    listBoxLog.TopIndex = listBoxLog.Items.Count - 1;
                    btn_import.Text = "Finish";
                    btn_import.Enabled = true;
                    btn_cancel.Enabled = false;
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

        public async Task<int> UpsertSanPhamAsync(string maSPNB, string maSPKH, string tenSanPham, Dictionary<string, string> thuocTinhValues, List<DMThuocTinh> thuocTinhs, DateTime currentTime)
        {
            using(var transaction = _DbConnect.Database.BeginTransaction())
            {
                try
                {
                    var existing = await _DbConnect.DMSanPham
                        .FirstOrDefaultAsync(x => x.FlagDel == 0
                            && x.MaSPNB.Trim() == maSPNB
                            && x.MaSPKH.Trim() == maSPKH);

                    int sanPhamId;

                    if (existing == null)
                    {
                        var newSP = new DMSanPham
                        {
                            CongTyId = 1,
                            MaSPNB = maSPNB,
                            MaSPKH = maSPKH,
                            TenSanPham = tenSanPham,
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
                        existing.TenSanPham = tenSanPham;
                        existing.UpdatedDate = currentTime;
                        sanPhamId = existing.Id;
                    }

                    foreach (var kvp in thuocTinhValues)
                    {
                        UpsertThuocTinh(sanPhamId, kvp.Key, kvp.Value, thuocTinhs, currentTime);
                    }

                    await _DbConnect.SaveChangesAsync();
                    transaction.Commit(); // Chỉ commit khi mọi thứ xong xuôi

                    return sanPhamId;
                }
                catch (Exception ex)
                {
                    transaction.Rollback(); // Nếu lỗi, rollback toàn bộ
                    throw new Exception("Có lỗi khi lưu sản phẩm: " + ex.Message, ex);
                }
            }
        }

        private void UpsertThuocTinh(int spId, string maThuocTinh, string value, List<DMThuocTinh> thuocTinhs, DateTime time)
        {
            if (string.IsNullOrWhiteSpace(value))
                return;

            var thuocTinhId = thuocTinhs
                .FirstOrDefault(x => x.MaThuocTinh.Equals(maThuocTinh, StringComparison.OrdinalIgnoreCase))?.Id;

            if (thuocTinhId == null)
                return;

            var existingThuocTinh = _DbConnect.SanPhamThuocTinh
                .FirstOrDefault(x => x.SanPhamId == spId
                    && x.ThuocTinhSPId == thuocTinhId
                    && x.FlagDel == 0);

            if (existingThuocTinh == null)
            {
                _DbConnect.SanPhamThuocTinh.Add(new SanPhamThuocTinh
                {
                    SanPhamId = spId,
                    ThuocTinhSPId = thuocTinhId.Value,
                    NoiDung = value,
                    FlagDel = 0,
                    CreatedDate = time,
                    UpdatedDate = time
                });
            }
            else
            {
                existingThuocTinh.NoiDung = value;
                existingThuocTinh.UpdatedDate = time;
            }
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
    }
}
