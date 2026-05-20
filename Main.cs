using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LTCSDL
{
    public partial class Main : Form
    {
        // 1. KHAI BÁO SỬ DỤNG CHUNG HỆ THỐNG DATAPROVIDER
        DataProvider data = new DataProvider();

        // Chuỗi kết nối dự phòng dùng riêng cho các hàm chạy Giao dịch (Transaction) phức tạp liên hoàn
        string cnStr = @"Data Source=HONGPHUC\SQLEXPRESS;Initial Catalog=LTCSDL;Integrated Security=True;TrustServerCertificate=True;";

        string quyenNguoiDung = "";
        string tenNguoiDung = "";

        public Main(string quyen, string user)
        {
            InitializeComponent();
            this.quyenNguoiDung = quyen;
            this.tenNguoiDung = user;
        }

        string TaoMaTheoThongSo()
        {
            string loai = cboPhanLoai.Text.Trim();
            string prefixLoai = "SP";
            if (loai == "RAM") prefixLoai = "RAM";
            else if (loai == "Ổ cứng") prefixLoai = "SSD";
            else if (loai == "Thẻ nhớ") prefixLoai = "SD";
            else if (loai == "CPU") prefixLoai = "CPU";
            else if (loai == "Màn hình") prefixLoai = "MH";

            string hang = txtThuongHieu.Text.Trim();
            string prefixHang = "XXX";
            if (hang.Length >= 3)
                prefixHang = hang.Substring(0, 3).ToUpper();
            else if (hang.Length > 0)
                prefixHang = hang.ToUpper();

            string code = DateTime.Now.ToString("ssmm");
            return $"{prefixLoai}-{prefixHang}-{code}";
        }

        // ==========================================================
        // 🖥️ TAB QUẢN LÝ SẢN PHẨM
        // ==========================================================
        void LoadData()
        {
            string sql = "SELECT * FROM SanPham";
            dgvSanPham.DataSource = data.ExecuteQuery(sql);
            dgvSanPham.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void Main_Load(object sender, EventArgs e)
        {
            txtPathAnh.MaxLength = int.MaxValue;
            PhanQuyen();
            LoadData();
            LoadKhachHang();
            LoadDonWeb();
            LoadNhaCungCap();
            LoadComboBoxNhapHang();
            LoadBaoHanh();

            // KÍCH HOẠT TIMER TỰ ĐỘNG LÀM MỚI ĐƠN WEB NGẦM (MỖI 5 GIÂY)
            Timer timerAutoLoad = new Timer();
            timerAutoLoad.Interval = 5000;
            timerAutoLoad.Tick += TimerAutoLoad_Tick;
            timerAutoLoad.Start();

            if (quyenNguoiDung != "Nhân viên")
            {
                LoadNhanVien();
                LoadChamCong();
            }

            lblUserHienThi.Text = "Xin chào: " + tenNguoiDung;

            this.StartPosition = FormStartPosition.CenterScreen;
            dgvSanPham.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvKetQuaTim.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvGioHang.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
        }

        void PhanQuyen()
        {
            if (quyenNguoiDung == "Nhân viên")
            {
                if (tabControl1.TabPages.Contains(tabPage1)) tabControl1.TabPages.Remove(tabPage1);
                if (tabControl1.TabPages.Contains(tabPage3)) tabControl1.TabPages.Remove(tabPage3);
                if (tabControl1.TabPages.Contains(tabPage5)) tabControl1.TabPages.Remove(tabPage5);

                btnXoaKH.Enabled = false;
                tabControl1.SelectedTab = tabPage2;
                this.Text = "CỬA HÀNG LINH KIỆN PC - [NHÂN VIÊN: " + DateTime.Now.ToShortDateString() + "]";
            }
            else if (quyenNguoiDung == "Admin" || quyenNguoiDung == "Quản trị viên")
            {
                this.Text = "CỬA HÀNG LINH KIỆN PC - [ADMIN]";
                btnXoaKH.Enabled = true;
            }
        }

        private void btnBanHang_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMaSP.Text)) return;
            string sql = "UPDATE SanPham SET [Tồn kho] = [Tồn kho] - 1 WHERE [Mã sản phẩm]=@ma AND [Tồn kho] > 0";
            SqlParameter[] pr = { new SqlParameter("@ma", txtMaSP.Text) };

            int rows = data.ExecuteNonQuery(sql, pr);
            if (rows > 0)
            {
                LoadData();
                MessageBox.Show("Đã bán 1 sản phẩm!");
            }
            else MessageBox.Show("Hết hàng hoặc mã sản phẩm không đúng!");
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMaSP.Text)) return;

            string sql = @"UPDATE SanPham 
                           SET [Tên sản phẩm]=@ten, [Thương hiệu]=@th, [Tồn kho]=@ton, 
                               [Giá thành (VNĐ)]=@gia, [Thông số kỹ thuật]=@ts, 
                               [Bảo hành]=@bh, [Phân loại]=@loai, HinhAnh=@anh 
                           WHERE [Mã sản phẩm]=@ma";

            SqlParameter[] pr = {
                new SqlParameter("@ten", txtTenSP.Text),
                new SqlParameter("@th", txtThuongHieu.Text),
                new SqlParameter("@ton", nmTonKho.Value),
                new SqlParameter("@gia", long.Parse(txtGia.Text.Replace(",", "").Replace(".", ""))),
                new SqlParameter("@ts", txtThongSo.Text),
                new SqlParameter("@bh", txtBaoHanh.Text),
                new SqlParameter("@loai", cboPhanLoai.Text),
                new SqlParameter("@anh", txtPathAnh.Text),
                new SqlParameter("@ma", txtMaSP.Text)
            };

            data.ExecuteNonQuery(sql, pr);
            LoadData();
            MessageBox.Show("Cập nhật thành công!", "Thông báo");
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMaSP.Text) || string.IsNullOrEmpty(txtTenSP.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ Mã và Tên sản phẩm!");
                return;
            }

            string sql = @"INSERT INTO SanPham 
                           ([Mã sản phẩm], [Tên sản phẩm], [Thương hiệu], [Giá thành (VNĐ)], 
                            [Thông số kỹ thuật], [Tồn kho], [Bảo hành], [Phân loại], HinhAnh) 
                           VALUES (@ma, @ten, @th, @gia, @ts, @tk, @bh, @loai, @anh)";

            SqlParameter[] pr = {
                new SqlParameter("@ma", txtMaSP.Text),
                new SqlParameter("@ten", txtTenSP.Text),
                new SqlParameter("@th", txtThuongHieu.Text),
                new SqlParameter("@gia", txtGia.Text),
                new SqlParameter("@ts", txtThongSo.Text),
                new SqlParameter("@tk", nmTonKho.Value),
                new SqlParameter("@bh", txtBaoHanh.Text),
                new SqlParameter("@loai", cboPhanLoai.Text),
                new SqlParameter("@anh", txtPathAnh.Text)
            };

            data.ExecuteNonQuery(sql, pr);
            MessageBox.Show("Thêm sản phẩm mới thành công!");
            LoadData();
            ClearInputs();
        }

        void ClearInputs()
        {
            txtMaSP.Clear();
            txtTenSP.Clear();
            txtGia.Clear();
            txtThongSo.Clear();
            nmTonKho.Value = 0;
            txtPathAnh.Clear();
            picSanPham_QuanLy.Image = null;
        }

        private void dgvSanPham_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvSanPham.Rows[e.RowIndex];
                txtMaSP.Text = row.Cells["Mã sản phẩm"].Value?.ToString();
                txtTenSP.Text = row.Cells["Tên sản phẩm"].Value?.ToString();
                txtThuongHieu.Text = row.Cells["Thương hiệu"].Value?.ToString();
                txtGia.Text = row.Cells["Giá thành (VNĐ)"].Value?.ToString();
                txtThongSo.Text = row.Cells["Thông số kỹ thuật"].Value?.ToString();
                txtBaoHanh.Text = row.Cells["Bảo hành"].Value?.ToString();
                nmTonKho.Value = Convert.ToDecimal(row.Cells["Tồn kho"].Value ?? 0);
                cboPhanLoai.Text = row.Cells["Phân loại"].Value?.ToString();

                string chuoiAnh = row.Cells["HinhAnh"].Value?.ToString() ?? "";
                txtPathAnh.Text = chuoiAnh;
                LoadImageToPictureBox(chuoiAnh, picSanPham_QuanLy);
            }
        }

        private void dgvSanPham_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dgvSanPham_CellClick(sender, e);
        }

        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            txtMaSP.Clear(); txtTenSP.Clear(); txtThuongHieu.Clear();
            nmTonKho.Value = 0; txtGia.Clear(); txtThongSo.Clear();
            txtBaoHanh.Clear(); cboPhanLoai.SelectedIndex = -1;
            LoadData();
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Xóa linh kiện này?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string sql = "DELETE FROM SanPham WHERE [Mã sản phẩm]=@ma";
                SqlParameter[] pr = { new SqlParameter("@ma", txtMaSP.Text) };

                int kq = data.ExecuteNonQuery(sql, pr);
                if (kq > 0)
                {
                    LoadData();
                    ClearInputs();
                    MessageBox.Show("Đã xóa món hàng và làm mới ô nhập!");
                }
            }
        }

        private void UpdateMaSP(object sender, EventArgs e)
        {
            if (txtMaSP.ReadOnly == false)
            {
                txtMaSP.Text = TaoMaTheoThongSo();
            }
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            string sql = "SELECT * FROM SanPham WHERE [Mã sản phẩm] LIKE @key OR [Tên sản phẩm] LIKE @key";
            SqlParameter[] pr = { new SqlParameter("@key", "%" + txtSearch.Text + "%") };
            dgvSanPham.DataSource = data.ExecuteQuery(sql, pr);
        }

        private void btnLuuAnh_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMaSP.Text))
            {
                MessageBox.Show("Vui lòng chọn một sản phẩm từ bảng trước khi lưu ảnh!");
                return;
            }
            if (string.IsNullOrEmpty(txtPathAnh.Text))
            {
                MessageBox.Show("Vui lòng nhấn 'Chọn ảnh' để chọn hình trước!");
                return;
            }

            string sql = "UPDATE SanPham SET HinhAnh = @anh WHERE [Mã sản phẩm] = @ma";
            SqlParameter[] pr = {
                new SqlParameter("@anh", txtPathAnh.Text),
                new SqlParameter("@ma", txtMaSP.Text)
            };

            int rows = data.ExecuteNonQuery(sql, pr);
            if (rows > 0)
            {
                MessageBox.Show("Đã lưu ảnh cho sản phẩm: " + txtMaSP.Text, "Thành công");
                LoadData();
            }
            else MessageBox.Show("Không tìm thấy sản phẩm này để cập nhật!");
        }

        private void btnChonAnh_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif";

            if (open.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    byte[] imageBytes = File.ReadAllBytes(open.FileName);
                    string base64String = "data:image/jpeg;base64," + Convert.ToBase64String(imageBytes);
                    txtPathAnh.Text = base64String;
                    LoadImageToPictureBox(base64String, picSanPham_QuanLy);
                }
                catch (Exception ex) { MessageBox.Show("Lỗi khi chuyển đổi file ảnh: " + ex.Message); }
            }
        }

        private void LoadImageToPictureBox(string chuoiBase64, PictureBox pic)
        {
            try
            {
                if (pic.Image != null) pic.Image.Dispose();

                if (!string.IsNullOrEmpty(chuoiBase64) && chuoiBase64.Contains(","))
                {
                    string base64Data = chuoiBase64.Split(',')[1];
                    byte[] imageBytes = Convert.FromBase64String(base64Data);
                    using (MemoryStream ms = new MemoryStream(imageBytes))
                    {
                        pic.Image = Image.FromStream(ms);
                    }
                    pic.SizeMode = PictureBoxSizeMode.Zoom;
                }
                else pic.Image = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi nạp ảnh Base64: " + ex.Message);
                pic.Image = null;
            }
        }


        // ==========================================================
        // 🛒 TAB BÁN HÀNG TẠI QUẦY (POS)
        // ==========================================================
        private void txtTimKiemBanHang_TextChanged(object sender, EventArgs e)
        {
            string sql = "SELECT [Mã sản phẩm], [Tên sản phẩm], [Giá thành (VNĐ)], [Tồn kho], HinhAnh " +
                         "FROM SanPham WHERE [Mã sản phẩm] LIKE @key OR [Tên sản phẩm] LIKE @key";
            SqlParameter[] pr = { new SqlParameter("@key", "%" + txtTimKiemBanHang.Text + "%") };
            dgvKetQuaTim.DataSource = data.ExecuteQuery(sql, pr);
        }

        private void dgvKetQuaTim_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvKetQuaTim.Rows[e.RowIndex];
                lblTenSP.Text = "Tên sản phẩm: " + row.Cells["Tên sản phẩm"].Value?.ToString();
                lblGiaBan.Text = "Giá bán: " + row.Cells["Giá thành (VNĐ)"].Value?.ToString();
                lblThongSo.Text = "Tồn kho: " + row.Cells["Tồn kho"].Value?.ToString();

                string tenFile = row.Cells["HinhAnh"].Value?.ToString() ?? "";
                LoadImageToPictureBox(tenFile, picSanPham);
            }
        }

        private void btnThemGioHang_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(lblTenSP.Text) || lblTenSP.Text == "Tên sản phẩm: ")
            {
                MessageBox.Show("Vui lòng chọn một sản phẩm để thêm vào giỏ!", "Nhắc nhở");
                return;
            }

            string tonKhoRaw = lblThongSo.Text;
            string cleanTonKho = new string(tonKhoRaw.Where(char.IsDigit).ToArray());
            int tonKho = 0;
            int.TryParse(cleanTonKho, out tonKho);

            int soLuong = (int)nmSoLuongMua.Value;

            if (soLuong <= 0)
            {
                MessageBox.Show("Số lượng mua phải lớn hơn 0!", "Nhắc nhở");
                return;
            }
            if (soLuong > tonKho)
            {
                MessageBox.Show("Không đủ hàng! Trong kho hiện chỉ còn " + tonKho + " sản phẩm.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string giaRaw = lblGiaBan.Text;
            string cleanGia = new string(giaRaw.Where(char.IsDigit).ToArray());
            long giaBan = 0;

            if (long.TryParse(cleanGia, out giaBan))
            {
                long thanhTien = giaBan * soLuong;
                string maSP = dgvKetQuaTim.CurrentRow.Cells["Mã sản phẩm"].Value.ToString();

                dgvGioHang.Rows.Add(maSP, lblTenSP.Text.Replace("Tên sản phẩm: ", ""), soLuong, giaBan, thanhTien);
                TinhTongTien();
            }
            else MessageBox.Show("Không thể lấy giá bán từ nhãn!");
        }

        void TinhTongTien()
        {
            long tong = 0;
            foreach (DataGridViewRow row in dgvGioHang.Rows)
            {
                if (row.Cells[4].Value != null)
                {
                    tong += Convert.ToInt64(row.Cells[4].Value);
                }
            }
            lblTongTien.Text = tong.ToString("N0") + " VNĐ";
        }

        private void btnXoaMon_Click(object sender, EventArgs e)
        {
            if (dgvGioHang.CurrentRow != null && dgvGioHang.CurrentRow.Index != -1)
            {
                if (MessageBox.Show("Bạn có chắc muốn xóa món này khỏi giỏ hàng?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    dgvGioHang.Rows.RemoveAt(dgvGioHang.CurrentRow.Index);
                    TinhTongTien();
                    MessageBox.Show("Đã xóa món hàng!");
                }
            }
            else MessageBox.Show("Vui lòng chọn món hàng cần xóa trong giỏ!");
        }

        // ĐỂ BẢO ĐẢM TÍNH AN TOÀN CHO HỆ THỐNG GIAO DỊCH, GIỮ NGUYÊN TRANSACTIONS DÙNG CNSTR
        private void btnThanhToan_Click(object sender, EventArgs e)
        {
            if (dgvGioHang.Rows.Count == 0 || (dgvGioHang.Rows.Count == 1 && dgvGioHang.Rows[0].IsNewRow))
            {
                MessageBox.Show("Giỏ hàng đang trống, vui lòng thêm sản phẩm!", "Thông báo");
                return;
            }
            if (string.IsNullOrEmpty(txtTenKH_BanHang.Text) || string.IsNullOrEmpty(txtSDT_BanHang.Text))
            {
                MessageBox.Show("Vui lòng nhập Tên và Số điện thoại khách hàng trước khi thanh toán!", "Nhắc nhở");
                txtTenKH_BanHang.Focus();
                return;
            }

            string tenNV = lblUserHienThi.Text.Replace("Xin chào: ", "").Trim();

            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                conn.Open();
                SqlTransaction tran = conn.BeginTransaction();
                try
                {
                    string sdt = txtSDT_BanHang.Text.Trim();
                    string tenKH = txtTenKH_BanHang.Text.Trim();
                    int maKH = 0;

                    string sqlCheckKH = "SELECT MaKH FROM KhachHang WHERE SDT = @sdt";
                    SqlCommand cmdCheck = new SqlCommand(sqlCheckKH, conn, tran);
                    cmdCheck.Parameters.AddWithValue("@sdt", sdt);
                    object resultKH = cmdCheck.ExecuteScalar();

                    if (resultKH != null) maKH = Convert.ToInt32(resultKH);
                    else
                    {
                        string sqlInsertKH = "INSERT INTO KhachHang (TenKH, SDT, NgayMua) OUTPUT INSERTED.MaKH VALUES (@ten, @sdt, @ngay)";
                        SqlCommand cmdInsertKH = new SqlCommand(sqlInsertKH, conn, tran);
                        cmdInsertKH.Parameters.AddWithValue("@ten", tenKH);
                        cmdInsertKH.Parameters.AddWithValue("@sdt", sdt);
                        cmdInsertKH.Parameters.AddWithValue("@ngay", DateTime.Now);
                        maKH = (int)cmdInsertKH.ExecuteScalar();
                    }

                    string maHD = "POS" + DateTime.Now.ToString("ddMMyyHHmmss");
                    string cleanTongTien = lblTongTien.Text.Replace(".", "").Replace("VNĐ", "").Replace(" ", "").Trim();
                    long tongTien = long.Parse(cleanTongTien);

                    string sqlHD = "INSERT INTO HoaDon (MaHD, NgayBan, TongTien, NhanVien, MaKH, NguonDon, TrangThai) " +
                                   "VALUES (@ma, @ngay, @tong, @nv, @makh, 'Tại cửa hàng', 'Hoàn thành')";
                    SqlCommand cmdHD = new SqlCommand(sqlHD, conn, tran);
                    cmdHD.Parameters.AddWithValue("@ma", maHD);
                    cmdHD.Parameters.AddWithValue("@ngay", DateTime.Now);
                    cmdHD.Parameters.AddWithValue("@tong", tongTien);
                    cmdHD.Parameters.AddWithValue("@nv", tenNV);
                    cmdHD.Parameters.AddWithValue("@makh", maKH);
                    cmdHD.ExecuteNonQuery();

                    foreach (DataGridViewRow row in dgvGioHang.Rows)
                    {
                        if (row.IsNewRow || row.Cells[0].Value == null) continue;

                        string maSP = row.Cells[0].Value.ToString();
                        int slMua = Convert.ToInt32(row.Cells[2].Value);
                        long gia = Convert.ToInt64(row.Cells[3].Value);
                        long tt = Convert.ToInt64(row.Cells[4].Value);

                        string sqlCT = "INSERT INTO ChiTietHoaDon (MaHD, MaSP, SoLuong, DonGia, ThanhTien, NhanVien) " +
                                       "VALUES (@mahd, @masp, @sl, @dg, @tt, @nv)";
                        SqlCommand cmdCT = new SqlCommand(sqlCT, conn, tran);
                        cmdCT.Parameters.AddWithValue("@mahd", maHD);
                        cmdCT.Parameters.AddWithValue("@masp", maSP);
                        cmdCT.Parameters.AddWithValue("@sl", slMua);
                        cmdCT.Parameters.AddWithValue("@dg", gia);
                        cmdCT.Parameters.AddWithValue("@tt", tt);
                        cmdCT.Parameters.AddWithValue("@nv", tenNV);
                        cmdCT.ExecuteNonQuery();

                        string sqlUp = "UPDATE SanPham SET [Tồn kho] = [Tồn kho] - @sl WHERE [Mã sản phẩm] = @masp";
                        SqlCommand cmdUp = new SqlCommand(sqlUp, conn, tran);
                        cmdUp.Parameters.AddWithValue("@sl", slMua);
                        cmdUp.Parameters.AddWithValue("@masp", maSP);
                        cmdUp.ExecuteNonQuery();
                    }

                    tran.Commit();

                    if (MessageBox.Show("Thanh toán thành công! Bạn có muốn xem hóa đơn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        InHoaDonChinhChu();
                    }

                    dgvGioHang.Rows.Clear();
                    lblTongTien.Text = "0 VNĐ";
                    LoadData();
                    HienThiKetQuaTimKiem("");
                    ResetTabBanHang();
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    MessageBox.Show("Lỗi trong quá trình thanh toán: " + ex.Message, "Lỗi hệ thống");
                }
            }
        }

        void InHoaDonChinhChu()
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(VeHoaDon);

            PrintPreviewDialog ppd = new PrintPreviewDialog();
            ppd.Document = pd;
            ppd.WindowState = FormWindowState.Maximized;
            ppd.ShowDialog();
        }

        private void VeHoaDon(object sender, PrintPageEventArgs e)
        {
            Graphics g = e.Graphics;
            Font fCuaHang = new Font("Arial", 16, FontStyle.Bold);
            Font fTieuDe = new Font("Arial", 18, FontStyle.Bold);
            Font fChu = new Font("Arial", 11);
            Font fDam = new Font("Arial", 11, FontStyle.Bold);
            Pen pen = new Pen(Color.Black, 1);

            int y = 40; int x = 80;

            g.DrawString("HỒNG PHÚC COMPUTER", fCuaHang, Brushes.Blue, 280, y); y += 35;
            g.DrawString("HÓA ĐƠN BÁN HÀNG", fTieuDe, Brushes.Black, 300, y); y += 45;

            string maHD = "HD" + DateTime.Now.ToString("ddMMyyHHmmss");
            g.DrawString("Mã HD: " + maHD, fChu, Brushes.Black, x, y);
            g.DrawString("Ngày: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm"), fChu, Brushes.Black, 500, y); y += 35;

            g.DrawRectangle(pen, x, y, 680, 30);
            g.DrawString("Tên Sản Phẩm", fDam, Brushes.Black, x + 5, y + 7);
            g.DrawString("SL", fDam, Brushes.Black, x + 380, y + 7);
            g.DrawString("Đơn\".Giá", fDam, Brushes.Black, x + 450, y + 7);
            g.DrawString("Thành Tiền", fDam, Brushes.Black, x + 560, y + 7); y += 30;

            foreach (DataGridViewRow row in dgvGioHang.Rows)
            {
                if (row.IsNewRow || row.Cells[0].Value == null) continue;

                string ten = row.Cells[1].Value.ToString();
                string sl = row.Cells[2].Value.ToString();
                string gia = row.Cells[3].Value.ToString();
                string tt = row.Cells[4].Value.ToString();

                g.DrawString(ten, fChu, Brushes.Black, x + 5, y + 7);
                g.DrawString(sl, fChu, Brushes.Black, x + 380, y + 7);
                g.DrawString(gia, fChu, Brushes.Black, x + 450, y + 7);
                g.DrawString(tt, fChu, Brushes.Black, x + 560, y + 7); y += 30;
                g.DrawLine(pen, x, y, x + 680, y);
            }

            y += 15;
            g.DrawString("TỔNG TIỀN THANH TOÁN:", fDam, Brushes.Black, x + 350, y);
            g.DrawString(lblTongTien.Text, fDam, Brushes.Red, x + 560, y); y += 60;

            string nhanVien = lblUserHienThi.Text.Replace("Xin chào: ", "").Trim();
            g.DrawString("Người lập hóa đơn", fDam, Brushes.Black, x + 510, y); y += 20;
            g.DrawString("(Ký và ghi rõ họ tên)", new Font("Arial", 9, FontStyle.Italic), Brushes.Black, x + 515, y); y += 60;
            g.DrawString(nhanVien, fDam, Brushes.Black, x + 520, y); y += 50;
            g.DrawString("--- Cảm ơn quý khách và hẹn gặp lại! ---", new Font("Arial", 10, FontStyle.Italic), Brushes.Gray, 280, y);
        }

        void HienThiKetQuaTimKiem(string keyword)
        {
            string sql = "SELECT [Mã sản phẩm], [Tên sản phẩm], [Giá thành (VNĐ)], [Tồn kho], [Thông số kỹ thuật], HinhAnh " +
                         "FROM SanPham WHERE [Tên sản phẩm] LIKE @key";
            SqlParameter[] pr = { new SqlParameter("@key", "%" + keyword + "%") };
            dgvKetQuaTim.DataSource = data.ExecuteQuery(sql, pr);
        }

        void ResetTabBanHang()
        {
            lblTenSP.Text = "Tên sản phẩm: ";
            lblGiaBan.Text = "Giá bán: ";
            lblThongSo.Text = "Tồn kho: ";
            nmSoLuongMua.Value = 0;
            picSanPham.Image = null;
            txtTenKH_BanHang.Clear();
            txtSDT_BanHang.Clear();
        }


        // ==========================================================
        // 👥 TAB QUẢN LÝ NHÂN VIÊN & CHẤM CÔNG
        // ==========================================================
        void LoadChamCong()
        {
            try
            {
                string sql = @"SELECT CC.Username AS [Tài khoản], 
                                      CC.ThoiGianVao AS [Giờ Vào], 
                                      CC.ThoiGianRa AS [Giờ Ra], 
                                      ROUND(CC.TongGio, 2) AS [Tổng Giờ] 
                               FROM ChamCong CC ORDER BY CC.ThoiGianVao DESC";
                dgvChamCong.DataSource = data.ExecuteQuery(sql);
                dgvChamCong.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            catch (Exception ex) { Console.WriteLine("Chưa có bảng Chấm Công: " + ex.Message); }
        }

        void LoadNhanVien()
        {
            string sql = "SELECT Username, Password, TenNV, ChucVu, SDT, DiaChi FROM NhanVien";
            DataTable dt = data.ExecuteQuery(sql);
            dgvNhanVien.DataSource = dt;
            dgvNhanVien.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            if (dt.Columns.Contains("Username")) dgvNhanVien.Columns["Username"].HeaderText = "Tài khoản";
            if (dt.Columns.Contains("Password")) dgvNhanVien.Columns["Password"].HeaderText = "Mật khẩu";
            if (dt.Columns.Contains("TenNV")) dgvNhanVien.Columns["TenNV"].HeaderText = "Họ và Tên";
            if (dt.Columns.Contains("ChucVu")) dgvNhanVien.Columns["ChucVu"].HeaderText = "Chức vụ";
            if (dt.Columns.Contains("SDT")) dgvNhanVien.Columns["SDT"].HeaderText = "Số điện thoại";
            if (dt.Columns.Contains("DiaChi")) dgvNhanVien.Columns["DiaChi"].HeaderText = "Địa chỉ";
        }

        private void dgvNhanVien_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvNhanVien.Rows[e.RowIndex];
                txtUsername.Text = row.Cells["Username"].Value?.ToString();
                txtPassword.Text = row.Cells["Password"].Value?.ToString();
                txtTenNV.Text = row.Cells["TenNV"].Value?.ToString();
                txtSDT.Text = row.Cells["SDT"].Value?.ToString();
                txtDiaChi.Text = row.Cells["DiaChi"].Value?.ToString();
                cboChucVu.Text = row.Cells["ChucVu"].Value?.ToString();

                txtUsername.ReadOnly = true;

                string userSelected = row.Cells["Username"].Value?.ToString();
                string sql = @"SELECT ThoiGianVao AS [Giờ Vào], ThoiGianRa AS [Giờ Ra], ROUND(TongGio, 2) AS [Tổng Giờ] 
                               FROM ChamCong WHERE Username = @user ORDER BY ThoiGianVao DESC";
                SqlParameter[] pr = { new SqlParameter("@user", userSelected) };
                dgvChamCong.DataSource = data.ExecuteQuery(sql, pr);
            }
        }

        private void btnThemNV_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtUsername.Text) || string.IsNullOrEmpty(txtPassword.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ Tài khoản và Mật khẩu!");
                return;
            }
            try
            {
                string sql = "INSERT INTO NhanVien (Username, Password, TenNV, ChucVu, SDT, DiaChi) VALUES (@user, @pass, @ten, @cv, @sdt, @dc)";
                SqlParameter[] pr = {
                    new SqlParameter("@user", txtUsername.Text.Trim()),
                    new SqlParameter("@pass", txtPassword.Text.Trim()),
                    new SqlParameter("@ten", txtTenNV.Text.Trim()),
                    new SqlParameter("@cv", cboChucVu.Text),
                    new SqlParameter("@sdt", txtSDT.Text.Trim()),
                    new SqlParameter("@dc", txtDiaChi.Text.Trim())
                };
                data.ExecuteNonQuery(sql, pr);
                MessageBox.Show("Thêm nhân viên: " + txtUsername.Text + " thành công!", "Thông báo");
                LoadNhanVien();
                btnLammoiNV_Click(sender, e);
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: Tài khoản đã tồn tại hoặc hệ thống lỗi! \nChi tiết: " + ex.Message); }
        }

        private void btnSuaNV_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtUsername.Text)) return;
            string sql = "UPDATE NhanVien SET Password=@pass, TenNV=@ten, ChucVu=@cv, SDT=@sdt, DiaChi=@dc WHERE Username=@user";
            SqlParameter[] pr = {
                new SqlParameter("@user", txtUsername.Text.Trim()),
                new SqlParameter("@pass", txtPassword.Text.Trim()),
                new SqlParameter("@ten", txtTenNV.Text.Trim()),
                new SqlParameter("@cv", cboChucVu.Text),
                new SqlParameter("@sdt", txtSDT.Text.Trim()),
                new SqlParameter("@dc", txtDiaChi.Text.Trim())
            };
            int kq = data.ExecuteNonQuery(sql, pr);
            if (kq > 0)
            {
                MessageBox.Show("Cập nhật thông tin thành công!");
                LoadNhanVien();
            }
        }

        private void btnXoaNV_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtUsername.Text)) return;
            if (txtUsername.Text.ToLower() == "admin")
            {
                MessageBox.Show("Không thể xóa tài khoản Admin hệ thống!");
                return;
            }

            DialogResult dr = MessageBox.Show("Bạn có chắc chắn muốn xóa nhân viên " + txtUsername.Text + "?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dr == DialogResult.Yes)
            {
                try
                {
                    string sql = "DELETE FROM NhanVien WHERE Username=@user";
                    SqlParameter[] pr = { new SqlParameter("@user", txtUsername.Text.Trim()) };
                    data.ExecuteNonQuery(sql, pr);
                    MessageBox.Show("Đã xóa nhân viên!");
                    LoadNhanVien();
                    btnLammoiNV_Click(sender, e);
                }
                catch (Exception ex) { MessageBox.Show("Lỗi: Nhân viên này có liên kết hóa đơn, không thể xóa! \nChi tiết: " + ex.Message); }
            }
        }

        private void btnLammoiNV_Click(object sender, EventArgs e)
        {
            txtUsername.Clear(); txtPassword.Clear(); txtTenNV.Clear(); txtSDT.Clear(); txtDiaChi.Clear();
            cboChucVu.SelectedIndex = -1; txtUsername.ReadOnly = false; txtUsername.Focus();
            LoadChamCong();
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!string.IsNullOrEmpty(tenNguoiDung))
            {
                string sql = @"UPDATE ChamCong SET ThoiGianRa = @ra, TongGio = CAST(DATEDIFF(SECOND, ThoiGianVao, @ra) AS float) / 3600.0 
                               WHERE Username = @user AND ThoiGianRa IS NULL";
                SqlParameter[] pr = { new SqlParameter("@ra", DateTime.Now), new SqlParameter("@user", tenNguoiDung) };
                data.ExecuteNonQuery(sql, pr);
            }
        }


        // ==========================================================
        // 👥 TAB QUẢN LÝ KHÁCH HÀNG & LIÊN THÔNG BẢO HÀNH CHUẨN
        // ==========================================================
        void LoadKhachHang()
        {
            string sql = @"
                SELECT CASE WHEN K.MaSP IS NULL OR K.MaSP = '' THEN 'WEB-' + CAST(K.MaKH AS VARCHAR) ELSE CAST(K.MaKH AS VARCHAR) END AS [Mã KH], 
                       K.TenKH AS [Họ Tên], K.SDT AS [Số Điện Thoại], 
                       CASE WHEN K.MaSP IS NOT NULL AND K.MaSP <> '' THEN K.MaSP
                            WHEN (SELECT COUNT(*) FROM HoaDon H INNER JOIN ChiTietHoaDon C ON H.MaHD = C.MaHD WHERE H.MaKH = K.MaKH) = 1 
                                THEN (SELECT TOP 1 S.[Tên sản phẩm] + ' (x' + CAST(C.SoLuong AS VARCHAR) + ')' FROM HoaDon H INNER JOIN ChiTietHoaDon C ON H.MaHD = C.MaHD INNER JOIN SanPham S ON C.MaSP = S.[Mã sản phẩm] WHERE H.MaKH = K.MaKH)
                            ELSE N'Đơn Web (Nhiều món)' END AS [Sản Phẩm], 
                       K.NgayMua AS [Ngày Mua], ISNULL(K.ThoiHanBH, 0) AS [BH (Tháng)], K.NgayMua AS [Hết Hạn] 
                FROM KhachHang K
                UNION ALL 
                SELECT 'WEB-' + CAST(Id AS VARCHAR) AS [Mã KH], HoTen AS [Họ Tên], SoDienThoai AS [Số Điện Thoại], DanhSachSanPham AS [Sản Phẩm], NgayDat AS [Ngày Mua], 0 AS [BH (Tháng)], NgayDat AS [Hết Hạn] FROM ChiTietKhachHang 
                ORDER BY [Ngày Mua] DESC";

            DataTable dt = data.ExecuteQuery(sql);

            foreach (DataRow row in dt.Rows)
            {
                string maKH = row["Mã KH"].ToString();
                if (maKH.StartsWith("WEB-"))
                {
                    DateTime ngayMua = Convert.ToDateTime(row["Ngày Mua"]);
                    string spStr = row["Sản Phẩm"].ToString();

                    if (spStr == "Đơn Web (Nhiều món)")
                    {
                        row["BH (Tháng)"] = 0;
                        row["Hết Hạn"] = DBNull.Value;
                    }
                    else
                    {
                        int bhThang = 0; string tenSPLookup = "";
                        int pos = spStr.IndexOf(" (x"); if (pos == -1) pos = spStr.IndexOf(" x"); if (pos == -1) pos = spStr.IndexOf("=");
                        if (pos > 0) tenSPLookup = spStr.Substring(0, pos).Trim(); else tenSPLookup = spStr.Trim();

                        string sqlLookup = "SELECT [Bảo hành] FROM SanPham WHERE [Tên sản phẩm] LIKE @ten";
                        SqlParameter[] pr = { new SqlParameter("@ten", "%" + tenSPLookup + "%") };
                        object kq = data.ExecuteScalar(sqlLookup, pr);

                        if (kq != null && kq != DBNull.Value)
                        {
                            string numStr = new string(kq.ToString().Where(char.IsDigit).ToArray());
                            if (!string.IsNullOrEmpty(numStr)) bhThang = int.Parse(numStr);
                        }
                        row["BH (Tháng)"] = bhThang; row["Hết Hạn"] = ngayMua.AddMonths(bhThang);
                    }
                }
                else
                {
                    int bh = Convert.ToInt32(row["BH (Tháng)"]);
                    row["Hết Hạn"] = Convert.ToDateTime(row["Ngày Mua"]).AddMonths(bh);
                }
            }
            dgvKhachHang.DataSource = dt;
            dgvKhachHang.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void dgvKhachHang_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvKhachHang.Rows[e.RowIndex];
                txtMaKH.Text = row.Cells["Mã KH"].Value?.ToString();
                txtTenKH.Text = row.Cells["Họ Tên"].Value?.ToString();
                txtSDT_KH.Text = row.Cells["Số Điện Thoại"].Value?.ToString();
                txtMaSP_Mua.Text = row.Cells["Sản Phẩm"].Value?.ToString();

                if (row.Cells["Ngày Mua"].Value != DBNull.Value && row.Cells["Ngày Mua"].Value != null)
                    dtpNgayMua.Value = Convert.ToDateTime(row.Cells["Ngày Mua"].Value);

                if (row.Cells["BH (Tháng)"].Value != DBNull.Value && row.Cells["BH (Tháng)"].Value != null)
                    nmThoiHanBH.Value = Convert.ToDecimal(row.Cells["BH (Tháng)"].Value);
                else nmThoiHanBH.Value = 0;

                TinhNgayHetHan();

                string maKHRaw = txtMaKH.Text;
                if (maKHRaw.StartsWith("WEB-") && txtMaSP_Mua.Text == "Đơn Web (Nhiều món)")
                {
                    string maKHId = maKHRaw.Replace("WEB-", "").Trim();
                    string sqlLookup = @"SELECT S.[Tên sản phẩm] AS [TenSP], C.SoLuong AS [SL], ISNULL(S.[Bảo hành], N'Không bảo hành') AS [BH]
                                         FROM HoaDon H INNER JOIN ChiTietHoaDon C ON H.MaHD = C.MaHD INNER JOIN SanPham S ON C.MaSP = S.[Mã sản phẩm]
                                         WHERE H.MaKH = @maKH";

                    SqlParameter[] prPopup = { new SqlParameter("@maKH", maKHId) };
                    DataTable dtPopup = data.ExecuteQuery(sqlLookup, prPopup);

                    if (dtPopup.Rows.Count > 0)
                    {
                        string thongTinBH = $"📦 CHI TIẾT BẢO HÀNH ĐƠN HÀNG [{maKHRaw}]\n👤 Khách: {txtTenKH.Text}\n-----------------------------------\n\n";
                        int stt = 1;
                        foreach (DataRow r in dtPopup.Rows)
                        {
                            thongTinBH += $"{stt}. {r["TenSP"]} (SL: {r["SL"]})\n   👉 Bảo hành: {r["BH"]}\n\n";
                            stt++;
                        }
                        MessageBox.Show(thongTinBH, "Tra cứu bảo hành Web", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        void TinhNgayHetHan()
        {
            if (txtMaSP_Mua.Text == "Đơn Web (Nhiều món)")
            {
                txtNgayHetHan.Text = "Theo từng linh kiện";
                return;
            }
            DateTime ngayMua = dtpNgayMua.Value;
            int soThang = (int)nmThoiHanBH.Value;
            DateTime ngayHetHan = ngayMua.AddMonths(soThang);
            txtNgayHetHan.Text = ngayHetHan.ToString("dd/MM/yyyy");
        }

        private void txtTimKiemKH_TextChanged(object sender, EventArgs e)
        {
            string sql = @"
                SELECT CASE WHEN K.MaSP IS NULL OR K.MaSP = '' THEN 'WEB-' + CAST(K.MaKH AS VARCHAR) ELSE CAST(K.MaKH AS VARCHAR) END AS [Mã KH], K.TenKH AS [Họ Tên], K.SDT AS [Số Điện Thoại], CASE WHEN K.MaSP IS NOT NULL AND K.MaSP <> '' THEN K.MaSP WHEN (SELECT COUNT(*) FROM HoaDon H INNER JOIN ChiTietHoaDon C ON H.MaHD = C.MaHD WHERE H.MaKH = K.MaKH) = 1 THEN (SELECT TOP 1 S.[Tên sản phẩm] + ' (x' + CAST(C.SoLuong AS VARCHAR) + ')' FROM HoaDon H INNER JOIN ChiTietHoaDon C ON H.MaHD = C.MaHD INNER JOIN SanPham S ON C.MaSP = S.[Mã sản phẩm] WHERE H.MaKH = K.MaKH) ELSE N'Đơn Web (Nhiều món)' END AS [Sản Phẩm], K.NgayMua AS [Ngày Mua], ISNULL(K.ThoiHanBH, 0) AS [BH (Tháng)], K.NgayMua AS [Hết Hạn] 
                FROM KhachHang K WHERE K.TenKH LIKE @key OR K.SDT LIKE @key
                UNION ALL 
                SELECT 'WEB-' + CAST(Id AS VARCHAR) AS [Mã KH], HoTen AS [Họ Tên], SoDienThoai AS [Số Điện Thoại], DanhSachSanPham AS [Sản Phẩm], NgayDat AS [Ngày Mua], 0 AS [BH (Tháng)], NULL AS [Hết Hạn] FROM ChiTietKhachHang WHERE HoTen LIKE @key OR SoDienThoai LIKE @key
                ORDER BY [Ngày Mua] DESC";

            SqlParameter[] prSearch = { new SqlParameter("@key", "%" + txtTimKiemKH.Text.Trim() + "%") };
            DataTable dt = data.ExecuteQuery(sql, prSearch);

            foreach (DataRow row in dt.Rows)
            {
                string maKH = row["Mã KH"].ToString();
                if (maKH.StartsWith("WEB-"))
                {
                    DateTime ngayMua = Convert.ToDateTime(row["Ngày Mua"]);
                    string spStr = row["Sản Phẩm"].ToString();

                    if (spStr == "Đơn Web (Nhiều món)")
                    {
                        row["BH (Tháng)"] = 0; row["Hết Hạn"] = DBNull.Value;
                    }
                    else
                    {
                        int bhThang = 0; string tenSPLookup = "";
                        int pos = spStr.IndexOf(" (x"); if (pos == -1) pos = spStr.IndexOf(" x"); if (pos == -1) pos = spStr.IndexOf("=");
                        if (pos > 0) tenSPLookup = spStr.Substring(0, pos).Trim(); else tenSPLookup = spStr.Trim();

                        string sqlLookup = "SELECT [Bảo hành] FROM SanPham WHERE [Tên sản phẩm] LIKE @ten";
                        SqlParameter[] prLook = { new SqlParameter("@ten", "%" + tenSPLookup + "%") };
                        object kq = data.ExecuteScalar(sqlLookup, prLook);
                        if (kq != null && kq != DBNull.Value)
                        {
                            string numStr = new string(kq.ToString().Where(char.IsDigit).ToArray());
                            if (!string.IsNullOrEmpty(numStr)) bhThang = int.Parse(numStr);
                        }
                        row["BH (Tháng)"] = bhThang; row["Hết Hạn"] = ngayMua.AddMonths(bhThang);
                    }
                }
                else
                {
                    int bh = Convert.ToInt32(row["BH (Tháng)"]);
                    row["Hết Hạn"] = Convert.ToDateTime(row["Ngày Mua"]).AddMonths(bh);
                }
            }
            dgvKhachHang.DataSource = dt;
        }

        private void btnSuaKH_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMaKH.Text)) return;
            if (txtMaKH.Text.StartsWith("WEB"))
            {
                MessageBox.Show("Đây là lịch sử khách hàng Online cũ. Không thể sửa thông tin này!", "Cảnh báo");
                return;
            }

            string sql = "UPDATE KhachHang SET TenKH=@ten, SDT=@sdt, MaSP=@masp, NgayMua=@ngay, ThoiHanBH=@bh WHERE MaKH=@ma";
            SqlParameter[] pr = {
                new SqlParameter("@ten", txtTenKH.Text),
                new SqlParameter("@sdt", txtSDT_KH.Text),
                new SqlParameter("@masp", txtMaSP_Mua.Text),
                new SqlParameter("@ngay", dtpNgayMua.Value),
                new SqlParameter("@bh", nmThoiHanBH.Value),
                new SqlParameter("@ma", txtMaKH.Text)
            };
            data.ExecuteNonQuery(sql, pr);
            MessageBox.Show("Đã cập nhật thông tin khách hàng!");
            LoadKhachHang();
        }

        private void btnXoaKH_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMaKH.Text)) return;
            if (txtMaKH.Text.StartsWith("WEB"))
            {
                MessageBox.Show("Đây là lịch sử khách hàng Online cũ. Không thể xóa!", "Cảnh báo");
                return;
            }

            if (MessageBox.Show("Xóa lịch sử bảo hành của khách này?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string sql = "DELETE FROM KhachHang WHERE MaKH=@ma";
                SqlParameter[] pr = { new SqlParameter("@ma", txtMaKH.Text) };
                data.ExecuteNonQuery(sql, pr);
                LoadKhachHang();
                btnLamMoiKH_Click(sender, e);
            }
        }

        private void btnLamMoiKH_Click(object sender, EventArgs e)
        {
            txtMaKH.Clear(); txtTenKH.Clear(); txtSDT_KH.Clear(); txtMaSP_Mua.Clear();
            dtpNgayMua.Value = DateTime.Now; nmThoiHanBH.Value = 0; txtNgayHetHan.Clear();
            LoadKhachHang();
        }

        private void dtpNgayMua_ValueChanged(object sender, EventArgs e) { TinhNgayHetHan(); }
        private void nmThoiHanBH_ValueChanged(object sender, EventArgs e) { TinhNgayHetHan(); }


        // ==========================================================
        // 🔄 TAB DUYỆT ĐƠN HÀNG TỪ WEBSITE ONLINE
        // ==========================================================
        void LoadDonWeb()
        {
            string sql = @"SELECT H.MaHD AS [Mã Đơn], K.TenKH AS [Khách Hàng], K.SDT AS [SĐT], H.NgayBan AS [Ngày Đặt], H.TongTien AS [Tổng Tiền], H.TrangThai AS [Trạng Thái]
                           FROM HoaDon H INNER JOIN KhachHang K ON H.MaKH = K.MaKH
                           WHERE H.NguonDon = 'Web Online'
                           ORDER BY CASE WHEN H.TrangThai = N'Chờ xác nhận' THEN 1 ELSE 2 END, H.NgayBan DESC";
            dgvDonWeb.DataSource = data.ExecuteQuery(sql);
            dgvDonWeb.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvDonWeb.ClearSelection();
        }

        private void TimerAutoLoad_Tick(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab != null && tabControl1.SelectedTab.Text == "Duyệt đơn Web")
            {
                string maHD_DangXem = "";
                if (dgvDonWeb.CurrentRow != null && dgvDonWeb.CurrentRow.Index >= 0)
                {
                    maHD_DangXem = dgvDonWeb.CurrentRow.Cells["Mã Đơn"].Value?.ToString() ?? "";
                }

                LoadDonWeb();
                this.Text = "CỬA HÀNG LINH KIỆN PC - [Đã cập nhật lúc: " + DateTime.Now.ToString("HH:mm:ss") + "]";

                bool timThay = false;
                if (!string.IsNullOrEmpty(maHD_DangXem))
                {
                    foreach (DataGridViewRow row in dgvDonWeb.Rows)
                    {
                        if (row.Cells["Mã Đơn"].Value?.ToString() == maHD_DangXem)
                        {
                            row.Selected = true;
                            dgvDonWeb.CurrentCell = row.Cells[0];
                            timThay = true;
                            break;
                        }
                    }
                }
                if (!timThay) dgvDonWeb.ClearSelection();
            }
        }

        private void dgvDonWeb_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                string maHD = dgvDonWeb.Rows[e.RowIndex].Cells["Mã Đơn"].Value.ToString();
                string sql = @"SELECT S.[Tên sản phẩm] AS [Sản Phẩm], C.SoLuong AS [Số Lượng], C.DonGia AS [Đơn Giá], C.ThanhTien AS [Thành Tiền]
                               FROM ChiTietHoaDon C INNER JOIN SanPham S ON C.MaSP = S.[Mã sản phẩm] WHERE C.MaHD = @ma";
                SqlParameter[] pr = { new SqlParameter("@ma", maHD) };
                dgvChiTietDon.DataSource = data.ExecuteQuery(sql, pr);
                dgvChiTietDon.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
        }

        private void btnDuyetDon_Click(object sender, EventArgs e)
        {
            if (dgvDonWeb.CurrentRow != null && dgvDonWeb.CurrentRow.Index >= 0)
            {
                string maHD = dgvDonWeb.CurrentRow.Cells["Mã Đơn"].Value.ToString();
                string trangThai = dgvDonWeb.CurrentRow.Cells["Trạng Thái"].Value.ToString();

                if (trangThai != "Chờ xác nhận")
                {
                    MessageBox.Show("Đơn này đã được xử lý rồi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (MessageBox.Show($"Xác nhận duyệt đơn hàng {maHD} và xuất kho?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    using (SqlConnection conn = new SqlConnection(cnStr))
                    {
                        conn.Open();
                        SqlTransaction tran = conn.BeginTransaction();
                        try
                        {
                            string sqlUpd = "UPDATE HoaDon SET TrangThai = N'Hoàn thành' WHERE MaHD = @ma";
                            SqlCommand cmdUpd = new SqlCommand(sqlUpd, conn, tran);
                            cmdUpd.Parameters.AddWithValue("@ma", maHD);
                            cmdUpd.ExecuteNonQuery();

                            string sqlTruKho = @"UPDATE SanPham SET [Tồn kho] = [Tồn kho] - C.SoLuong
                                                 FROM SanPham S INNER JOIN ChiTietHoaDon C ON S.[Mã sản phẩm] = C.MaSP WHERE C.MaHD = @ma";
                            SqlCommand cmdTruKho = new SqlCommand(sqlTruKho, conn, tran);
                            cmdTruKho.Parameters.AddWithValue("@ma", maHD);
                            cmdTruKho.ExecuteNonQuery();

                            tran.Commit();
                            MessageBox.Show("Đã duyệt đơn và xuất kho thành công!", "Hoàn tất");

                            LoadDonWeb();
                            dgvChiTietDon.DataSource = null;
                            LoadData();
                        }
                        catch (Exception ex)
                        {
                            tran.Rollback();
                            MessageBox.Show("Lỗi duyệt đơn: " + ex.Message);
                        }
                    }
                }
            }
            else MessageBox.Show("Vui lòng chọn 1 đơn hàng bên danh sách để duyệt!");
        }


        // ==========================================================
        // 📊 TAB THỐNG KÊ DOANH THU
        // ==========================================================
        private void btnThongKe_Click(object sender, EventArgs e)
        {
            DateTime tuNgay = dtpTuNgay.Value.Date;
            DateTime denNgay = dtpDenNgay.Value.Date.AddDays(1);

            string sql = @"SELECT MaHD AS [Mã Hóa Đơn], NgayBan AS [Ngày Bán], TongTien AS [Tổng Tiền (VNĐ)], NhanVien AS [Nhân Viên Bán] 
                           FROM HoaDon WHERE NgayBan >= @tuNgay AND NgayBan < @denNgay ORDER BY NgayBan DESC";

            SqlParameter[] pr = { new SqlParameter("@tuNgay", tuNgay), new SqlParameter("@denNgay", denNgay) };
            DataTable dt = data.ExecuteQuery(sql, pr);

            dgvDoanhThu.DataSource = dt;
            dgvDoanhThu.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            long tongTien = 0;
            int soDon = dt.Rows.Count;

            foreach (DataRow row in dt.Rows)
            {
                tongTien += Convert.ToInt64(row["Tổng Tiền (VNĐ)"]);
            }

            lblTongDoanhThu.Text = tongTien.ToString("N0") + " VNĐ";
            lblTongSoDon.Text = soDon.ToString() + " đơn";
        }

        void LoadNhaCungCap()
        {
            string sql = @"
        SELECT 
            NCC.MaNCC AS [Mã NCC], 
            NCC.TenNCC AS [Tên Nhà Cung Cấp], 
            NCC.SoDienThoai AS [SĐT], 
            NCC.DiaChi AS [Địa Chỉ],
            CASE WHEN (ISNULL(SUM(PN.TongTienNhap), 0) - ISNULL(SUM(PN.DaThanhToan), 0)) < 0 
                 THEN 0 
                 ELSE (ISNULL(SUM(PN.TongTienNhap), 0) - ISNULL(SUM(PN.DaThanhToan), 0)) 
            END AS [Tổng Nợ (VNĐ)]
        FROM NhaCungCap NCC
        LEFT JOIN PhieuNhap PN ON NCC.MaNCC = PN.MaNCC
        GROUP BY NCC.MaNCC, NCC.TenNCC, NCC.SoDienThoai, NCC.DiaChi";

            dgvNhaCungCap.DataSource = data.ExecuteQuery(sql);
            dgvNhaCungCap.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        // 2. Hàm load dữ liệu vào các ComboBox chọn lúc nhập hàng
        void LoadComboBoxNhapHang()
        {
            // Cbo Nhà Cung Cấp
            cboNhaCungCap.DataSource = data.ExecuteQuery("SELECT MaNCC, TenNCC FROM NhaCungCap");
            cboNhaCungCap.DisplayMember = "TenNCC";
            cboNhaCungCap.ValueMember = "MaNCC";

            // Cbo Sản Phẩm - ĐẢO TÊN SP LÊN TRƯỚC ĐỂ TÌM KIẾM CỰC MƯỢT
            string sqlSP = @"SELECT [Mã sản phẩm], 
                            [Tên sản phẩm] + ' - (Mã: ' + [Mã sản phẩm] + ') - Tồn: ' + CAST([Tồn kho] AS VARCHAR) AS ThongTin 
                     FROM SanPham";
            cboSanPhamNhap.DataSource = data.ExecuteQuery(sqlSP);

            cboSanPhamNhap.DisplayMember = "ThongTin";
            cboSanPhamNhap.ValueMember = "Mã sản phẩm";
        }

        // 3. Sự kiện tính tiền tự động khi gõ Số lượng hoặc Đơn giá
        private void TinhTienNhapHang(object sender, EventArgs e)
        {
            try
            {
                int soLuong = 0; decimal donGia = 0; decimal daTra = 0;
                int.TryParse(txtSoLuongNhap.Text, out soLuong);
                decimal.TryParse(txtDonGiaNhap.Text, out donGia);

                // Bắt buộc đọc từ cái ô vừa đổi tên ở trên
                decimal.TryParse(txtTraTruoc.Text, out daTra);

                decimal tongTien = soLuong * donGia;
                txtTongTienNhap.Text = tongTien.ToString("0");
                txtConNo.Text = (tongTien - daTra).ToString("0");
            }
            catch { }
        }

        // Trong hàm btnTaoPhieuNhap_Click bạn cũng phải sửa dòng lấy biến daTra tương tự:
        // decimal.TryParse(txtTraTruoc.Text, out daTra);
        // txtTraTruoc.Clear(); // Nhớ sửa chỗ dọn dẹp sau khi nhập xong
        // Các hàm sự kiện trống giữ lại để tránh lỗi Designer
        private void label1_Click(object sender, EventArgs e) { }
        private void label6_Click(object sender, EventArgs e) { }
        private void label8_Click(object sender, EventArgs e) { }
        private void button5_Click(object sender, EventArgs e) { }
        private void tabPage2_Click(object sender, EventArgs e) { }
        private void groupBox1_Enter(object sender, EventArgs e) { }
        private void label23_Click(object sender, EventArgs e) { }
        private void textBox1_TextChanged(object sender, EventArgs e) { }
        private void lblTongSoDon_Click(object sender, EventArgs e) { }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void btnTaoPhieuNhap_Click(object sender, EventArgs e)
        {
            if (cboNhaCungCap.SelectedValue == null || cboSanPhamNhap.SelectedValue == null) return;

            decimal soLuong = decimal.Parse(txtSoLuongNhap.Text);
            decimal donGia = decimal.Parse(txtDonGiaNhap.Text);
            decimal daTra = decimal.Parse(txtTraTruoc.Text.Replace(".", "").Replace(",", ""));

            decimal tongTien = soLuong * donGia;
            string maPN = "PN" + DateTime.Now.ToString("ddMMyyHHmmss");

            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                conn.Open();
                SqlTransaction tran = conn.BeginTransaction();
                try
                {
                    string sqlPN = @"INSERT INTO PhieuNhap (MaPhieuNhap, MaNCC, [Mã sản phẩm], NgayNhap, SoLuong, DonGiaNhap, TongTienNhap, DaThanhToan, ConNo) 
                             VALUES (@maPN, @maNCC, @maSP, GETDATE(), @sl, @gia, @tong, @datra, @no)";

                    SqlCommand cmd = new SqlCommand(sqlPN, conn, tran);
                    cmd.Parameters.AddWithValue("@maPN", maPN);
                    cmd.Parameters.AddWithValue("@maNCC", cboNhaCungCap.SelectedValue);
                    cmd.Parameters.AddWithValue("@maSP", cboSanPhamNhap.SelectedValue);
                    cmd.Parameters.AddWithValue("@sl", soLuong);
                    cmd.Parameters.AddWithValue("@gia", donGia);
                    cmd.Parameters.AddWithValue("@tong", tongTien);
                    cmd.Parameters.AddWithValue("@datra", daTra);
                    cmd.Parameters.AddWithValue("@no", tongTien - daTra);
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE SanPham SET [Tồn kho] = [Tồn kho] + @sl WHERE [Mã sản phẩm] = @maSP";
                    cmd.ExecuteNonQuery();

                    tran.Commit();
                    MessageBox.Show("Nhập hàng thành công!");
                    LoadNhaCungCap();
                    LoadData();
                }
                catch { tran.Rollback(); }
            }
        }
        void LoadBaoHanh(string filterTrangThai = "")
        {
            try
            {
                string sql = @"SELECT P.MaPhieuBH AS [Mã Phiếu], K.TenKH AS [Khách Hàng], K.SDT AS [SĐT], 
                              S.[Tên sản phẩm] AS [Sản Phẩm Lỗi], P.TinhTrangLoi AS [Tình Trạng Lỗi], 
                              P.NgayTiepNhan AS [Ngày Nhận], P.TrangThai AS [Trạng Thái]
                       FROM PhieuBaoHanh P 
                       INNER JOIN KhachHang K ON P.MaKH = K.MaKH
                       INNER JOIN SanPham S ON P.[Mã sản phẩm] = S.[Mã sản phẩm]";

                // Nếu có chọn lọc trạng thái (và không phải là "Tất cả")
                if (!string.IsNullOrEmpty(filterTrangThai) && filterTrangThai != "Tất cả")
                {
                    sql += " WHERE P.TrangThai = N'" + filterTrangThai + "'";
                }

                sql += " ORDER BY P.NgayTiepNhan DESC";

                dgvBaoHanh.DataSource = data.ExecuteQuery(sql);
                dgvBaoHanh.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            catch (Exception ex) { Console.WriteLine("Lỗi Load bảng Bảo Hành: " + ex.Message); }
        }
        private void btnLamMoiNCC_Click(object sender, EventArgs e)
        {
            txtSoLuongNhap.Clear();
            txtDonGiaNhap.Clear();
            txtDaThanhToan.Clear();
            txtTongTienNhap.Clear();
            txtConNo.Clear();

            // 2. Load lại dữ liệu mới nhất từ Database lên Bảng và ComboBox
            LoadNhaCungCap();
            LoadComboBoxNhapHang();

            // 3. Xóa trạng thái chọn của ComboBox (để trống)
            cboNhaCungCap.SelectedIndex = -1;
            cboSanPhamNhap.SelectedIndex = -1;

            // (Tùy chọn) Focus con trỏ chuột về ô chọn Nhà cung cấp để gõ tiếp cho nhanh
            cboNhaCungCap.Focus();
        }

        private void btnTimKH_BH_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSDT_TimBH.Text))
            {
                MessageBox.Show("Vui lòng nhập số điện thoại để tìm!", "Nhắc nhở");
                return;
            }

            string sql = "SELECT MaKH, TenKH FROM KhachHang WHERE SDT = @sdt";
            SqlParameter[] pr = { new SqlParameter("@sdt", txtSDT_TimBH.Text.Trim()) };
            DataTable dtKH = data.ExecuteQuery(sql, pr);

            if (dtKH.Rows.Count > 0)
            {
                // Lưu ngầm MaKH vào Tag để lát nữa dùng khi Tạo Phiếu
                txtSDT_TimBH.Tag = dtKH.Rows[0]["MaKH"].ToString();
                lblTenKhachBH.Text = "Khách hàng: " + dtKH.Rows[0]["TenKH"].ToString();

                // Load các sản phẩm khách đã mua vào ComboBox
                string sqlSP = @"SELECT S.[Mã sản phẩm], S.[Tên sản phẩm] 
                         FROM HoaDon H 
                         INNER JOIN ChiTietHoaDon C ON H.MaHD = C.MaHD 
                         INNER JOIN SanPham S ON C.MaSP = S.[Mã sản phẩm]
                         WHERE H.MaKH = @makh";
                SqlParameter[] prSP = { new SqlParameter("@makh", txtSDT_TimBH.Tag.ToString()) };

                DataTable dtSP = data.ExecuteQuery(sqlSP, prSP);
                if (dtSP.Rows.Count > 0)
                {
                    cboSanPhamLoi.DataSource = dtSP;
                    cboSanPhamLoi.DisplayMember = "Tên sản phẩm";
                    cboSanPhamLoi.ValueMember = "Mã sản phẩm";
                }
                else
                {
                    cboSanPhamLoi.DataSource = null;
                    MessageBox.Show("Khách hàng này chưa mua sản phẩm nào tại hệ thống!", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Không tìm thấy khách hàng nào với Số điện thoại này!", "Không thấy");
                lblTenKhachBH.Text = "Khách hàng: Chưa xác định";
                txtSDT_TimBH.Tag = null;
                cboSanPhamLoi.DataSource = null;
            }
        }

        private void btnTaoPhieuBH_Click(object sender, EventArgs e)
        {
            if (txtSDT_TimBH.Tag == null)
            {
                MessageBox.Show("Vui lòng tìm Khách hàng trước!", "Nhắc nhở"); return;
            }
            if (cboSanPhamLoi.SelectedValue == null)
            {
                MessageBox.Show("Vui lòng chọn Sản phẩm bị lỗi!", "Nhắc nhở"); return;
            }
            if (string.IsNullOrEmpty(txtTinhTrangLoi.Text))
            {
                MessageBox.Show("Vui lòng ghi nhận tình trạng lỗi của linh kiện!", "Nhắc nhở"); return;
            }

            string maBH = "BH" + DateTime.Now.ToString("ddMMyyHHmmss");
            string sql = "INSERT INTO PhieuBaoHanh (MaPhieuBH, MaKH, [Mã sản phẩm], TinhTrangLoi, TrangThai) VALUES (@ma, @makh, @masp, @loi, N'Đang xử lý')";

            SqlParameter[] pr = {
        new SqlParameter("@ma", maBH),
        new SqlParameter("@makh", int.Parse(txtSDT_TimBH.Tag.ToString())),
        new SqlParameter("@masp", cboSanPhamLoi.SelectedValue.ToString()),
        new SqlParameter("@loi", txtTinhTrangLoi.Text.Trim())
    };

            if (data.ExecuteNonQuery(sql, pr) > 0)
            {
                MessageBox.Show("Đã tiếp nhận thiết bị bảo hành thành công!", "Hoàn tất");
                LoadBaoHanh(cboLocTrangThai.Text); // Tải lại bảng
                txtTinhTrangLoi.Clear();
            }
        }
        private void cboLocTrangThai_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadBaoHanh(cboLocTrangThai.Text);
        }

        // ==========================================================
        // CỤM 3 NÚT CHỨC NĂNG (CAM - XANH - ĐỎ)
        // ==========================================================

        // Hàm dùng chung để update trạng thái (Viết 1 lần dùng cho cả 2 nút Cam/Xanh)
        void UpdateTrangThaiBaoHanh(string trangThaiMoi)
        {
            if (dgvBaoHanh.CurrentRow != null && dgvBaoHanh.CurrentRow.Index >= 0)
            {
                string maBH = dgvBaoHanh.CurrentRow.Cells["Mã Phiếu"].Value.ToString();
                string sql = "UPDATE PhieuBaoHanh SET TrangThai = @tt WHERE MaPhieuBH = @ma";

                SqlParameter[] pr = {
            new SqlParameter("@tt", trangThaiMoi),
            new SqlParameter("@ma", maBH)
        };

                data.ExecuteNonQuery(sql, pr);
                LoadBaoHanh(cboLocTrangThai.Text); // Tải lại bảng ngay lập tức
            }
            else
            {
                MessageBox.Show("Vui lòng click chọn 1 Phiếu Bảo Hành trong bảng bên trên!", "Nhắc nhở");
            }
        }

        private void cboLocTrangThai_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnDaGuiHang_Click(object sender, EventArgs e)
        {
            UpdateTrangThaiBaoHanh("Đã gửi hãng");
        }

        private void btnHoanThanh_Click(object sender, EventArgs e)
        {
            UpdateTrangThaiBaoHanh("Hoàn thành");
        }

        private void btnHuyPhieu_Click(object sender, EventArgs e)
        {
            if (dgvBaoHanh.CurrentRow != null && dgvBaoHanh.CurrentRow.Index >= 0)
            {
                DialogResult dr = MessageBox.Show("Bạn có chắc chắn muốn Hủy phiếu bảo hành này không?", "Xác nhận Hủy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dr == DialogResult.Yes)
                {
                    UpdateTrangThaiBaoHanh("Đã hủy");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng click chọn 1 Phiếu Bảo Hành trong bảng bên trên!", "Nhắc nhở");
            }
        }

        private void btnLamMoiBH_Click(object sender, EventArgs e)
        {
            txtSDT_TimBH.Clear();
            txtTinhTrangLoi.Clear();
            lblTenKhachBH.Text = "Khách hàng: ";
            txtSDT_TimBH.Tag = null;

            cboSanPhamLoi.DataSource = null; // Xóa list đồ
            cboLocTrangThai.SelectedIndex = -1; // Reset bộ lọc

            LoadBaoHanh(); // Tải lại bảng gốc
            txtSDT_TimBH.Focus();
        }

        private void cboLocTrangThai_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            LoadBaoHanh(cboLocTrangThai.Text);
        }

        private void btnThanhToanNo_Click(object sender, EventArgs e)
        {
            if (dgvNhaCungCap.CurrentRow == null) return;

            string maNCC = dgvNhaCungCap.CurrentRow.Cells["Mã NCC"].Value.ToString();
            decimal noHienTai = Convert.ToDecimal(dgvNhaCungCap.CurrentRow.Cells["Tổng Nợ (VNĐ)"].Value);
            decimal tienTra;

            if (!decimal.TryParse(txtDaThanhToan.Text.Replace(".", "").Replace(",", ""), out tienTra) || tienTra <= 0)
            {
                MessageBox.Show("Số tiền thanh toán không hợp lệ!"); return;
            }

            if (tienTra > noHienTai)
            {
                MessageBox.Show("Số tiền trả lớn hơn số nợ hiện tại!"); return;
            }

            // Logic thanh toán: Trừ dần vào các phiếu còn nợ
            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                conn.Open();
                SqlTransaction tran = conn.BeginTransaction();
                try
                {
                    // Lấy danh sách các phiếu còn nợ của NCC này, phiếu cũ ưu tiên trả trước
                    string sqlPhieu = "SELECT MaPhieuNhap, ConNo FROM PhieuNhap WHERE MaNCC = @maNCC AND ConNo > 0 ORDER BY NgayNhap ASC";
                    SqlCommand cmd = new SqlCommand(sqlPhieu, conn, tran);
                    cmd.Parameters.AddWithValue("@maNCC", maNCC);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dtPhieu = new DataTable();
                    da.Fill(dtPhieu);

                    decimal conLai = tienTra;
                    foreach (DataRow row in dtPhieu.Rows)
                    {
                        if (conLai <= 0) break;

                        string maPN = row["MaPhieuNhap"].ToString();
                        decimal noPhieu = Convert.ToDecimal(row["ConNo"]);
                        decimal thanhToanChoPhieu = Math.Min(conLai, noPhieu);

                        // Cập nhật phiếu đó
                        string sqlUpd = "UPDATE PhieuNhap SET DaThanhToan = DaThanhToan + @tien, ConNo = ConNo - @tien WHERE MaPhieuNhap = @maPN";
                        SqlCommand cmdUpd = new SqlCommand(sqlUpd, conn, tran);
                        cmdUpd.Parameters.AddWithValue("@tien", thanhToanChoPhieu);
                        cmdUpd.Parameters.AddWithValue("@maPN", maPN);
                        cmdUpd.ExecuteNonQuery();

                        conLai -= thanhToanChoPhieu;
                    }

                    tran.Commit();
                    MessageBox.Show("Thanh toán nợ thành công!");
                    txtDaThanhToan.Clear();
                    LoadNhaCungCap();
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    MessageBox.Show("Lỗi thanh toán: " + ex.Message);
                }
            }
        }
    }
}