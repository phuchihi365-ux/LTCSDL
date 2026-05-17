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
    public partial class Main: Form
    {
        string cnStr = @"Data Source=.\SQLEXPRESS;Initial Catalog=LTCSDL;Integrated Security=True";
        string quyenNguoiDung = "";
        string tenNguoiDung = "";
        // SỬA CHỖ NÀY: Thêm string quyen vào trong ngoặc
        public Main(string quyen, string user)
        {
            InitializeComponent();
            this.quyenNguoiDung = quyen; // Cất quyền vào biến
            this.tenNguoiDung = user;   // Cất tên vào biến
        }
       /* public Main()
        {
            InitializeComponent();
        }*/
        string TaoMaTheoThongSo()
        {
            // 1. Xử lý tiền tố Loại
            string loai = cboPhanLoai.Text.Trim();
            string prefixLoai = "SP"; // Mặc định là SP nếu chưa chọn loại
            if (loai == "RAM") prefixLoai = "RAM";
            else if (loai == "Ổ cứng") prefixLoai = "SSD";
            else if (loai == "Thẻ nhớ") prefixLoai = "SD";
            else if (loai == "CPU") prefixLoai = "CPU";
            else if (loai == "Màn hình") prefixLoai = "MH"; // Thêm cái này vì ảnh bạn đang nhập Màn hình

            // 2. Xử lý tiền tố Hãng
            string hang = txtThuongHieu.Text.Trim();
            string prefixHang = "XXX";
            if (hang.Length >= 3)
                prefixHang = hang.Substring(0, 3).ToUpper();
            else if (hang.Length > 0)
                prefixHang = hang.ToUpper();

            // 3. Lấy thời gian
            string code = DateTime.Now.ToString("ssmm");

            return $"{prefixLoai}-{prefixHang}-{code}";
        }
        void LoadData()
        {
            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                string sql = "SELECT * FROM SanPham";
                SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dgvSanPham.DataSource = dt;

                // THÊM DÒNG NÀY: Giúp bảng Quản lý dàn đều khít khung
                dgvSanPham.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
        }
 
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Main_Load(object sender, EventArgs e)
        {
            txtPathAnh.MaxLength = int.MaxValue;
            PhanQuyen();
            LoadData(); // Sản phẩm thì ai cũng phải thấy để bán hàng
            LoadKhachHang();
            LoadDonWeb();
            // --- SỬA Ở ĐÂY: Thêm điều kiện kiểm tra quyền ---
            // Chỉ Quản trị viên mới được phép load dữ liệu Nhân viên và Chấm công
            Timer timerAutoLoad = new Timer();
            timerAutoLoad.Interval = 5000; // Cứ 10.000 mili-giây (tức là 10 giây) tự động load 1 lần
            timerAutoLoad.Tick += TimerAutoLoad_Tick; // Gắn sự kiện chạy ngầm
            timerAutoLoad.Start(); // Bật công tắc cho Timer chạy
            if (quyenNguoiDung != "Nhân viên")
            {
                LoadNhanVien();
                LoadChamCong();
            }

            lblUserHienThi.Text = "Xin chào: " + tenNguoiDung;

            // Các thiết lập giao diện khít giữ nguyên...
            this.StartPosition = FormStartPosition.CenterScreen;
            dgvSanPham.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvKetQuaTim.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvGioHang.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
        }
        void PhanQuyen()
        {
            // 1. Nếu người đăng nhập là Nhân viên
            if (quyenNguoiDung == "Nhân viên")
            {
                // Ẩn tab Quản lý sản phẩm (tabPage1)
                if (tabControl1.TabPages.Contains(tabPage1))
                {
                    tabControl1.TabPages.Remove(tabPage1);
                }

                // Ẩn tab Quản lý nhân viên (tabPage3)
                if (tabControl1.TabPages.Contains(tabPage3))
                {
                    tabControl1.TabPages.Remove(tabPage3);
                }

                // (Nếu bạn đã làm Tab Thống kê doanh thu - tabPage5 thì ẩn luôn)
                 if (tabControl1.TabPages.Contains(tabPage5))
                {
                    tabControl1.TabPages.Remove(tabPage5);
                }

                // ==========================================
                // CỐ TÌNH KHÔNG ẨN TAB KHÁCH HÀNG (tabPage4)
                // ĐỂ NHÂN VIÊN CÓ THỂ BẤM VÀO SỬ DỤNG
                // ==========================================

                // Khóa nút XÓA khách hàng (Nhân viên chỉ được xem và sửa, không được xóa lịch sử)
                btnXoaKH.Enabled = false; // Tính năng bảo mật cực xịn!

                // Mặc định hiện tab Bán hàng đầu tiên
                tabControl1.SelectedTab = tabPage2;

                // Đổi tiêu đề Form
                this.Text = "CỬA HÀNG LINH KIỆN PC - [NHÂN VIÊN: " + DateTime.Now.ToShortDateString() + "]";
            }
            // 2. Nếu người đăng nhập là Admin
            else if (quyenNguoiDung == "Admin" || quyenNguoiDung == "Quản trị viên")
            {
                this.Text = "CỬA HÀNG LINH KIỆN PC - [ADMIN]";

                // Admin thì giữ nguyên toàn bộ Tab, và mở khóa nút Xóa KH
                btnXoaKH.Enabled = true;
            }
        }
        private void btnBanHang_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMaSP.Text)) return;
            try
            {
                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    conn.Open();
                    // Trừ 1 vào tồn kho nếu kho còn hàng (>0)
                    string sql = "UPDATE SanPham SET [Tồn kho] = [Tồn kho] - 1 WHERE [Mã sản phẩm]=@ma AND [Tồn kho] > 0";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@ma", txtMaSP.Text);

                    int rows = cmd.ExecuteNonQuery();
                    if (rows > 0)
                    {
                        LoadData();
                        MessageBox.Show("Đã bán 1 sản phẩm!");
                    }
                    else MessageBox.Show("Hết hàng!");
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message); }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMaSP.Text)) return;

            try
            {
                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    conn.Open();
                    string sql = @"UPDATE SanPham 
                           SET [Tên sản phẩm]=@ten, [Thương hiệu]=@th, [Tồn kho]=@ton, 
                               [Giá thành (VNĐ)]=@gia, [Thông số kỹ thuật]=@ts, 
                               [Bảo hành]=@bh, [Phân loại]=@loai, HinhAnh=@anh 
                           WHERE [Mã sản phẩm]=@ma";

                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@ten", txtTenSP.Text);
                    cmd.Parameters.AddWithValue("@th", txtThuongHieu.Text);
                    cmd.Parameters.AddWithValue("@ton", nmTonKho.Value);
                    cmd.Parameters.AddWithValue("@gia", long.Parse(txtGia.Text.Replace(",", "").Replace(".", "")));
                    cmd.Parameters.AddWithValue("@ts", txtThongSo.Text);
                    cmd.Parameters.AddWithValue("@bh", txtBaoHanh.Text);
                    cmd.Parameters.AddWithValue("@loai", cboPhanLoai.Text);
                    cmd.Parameters.AddWithValue("@anh", txtPathAnh.Text); // Đã bổ sung
                    cmd.Parameters.AddWithValue("@ma", txtMaSP.Text);

                    cmd.ExecuteNonQuery();
                    LoadData();
                    MessageBox.Show("Cập nhật thành công!", "Thông báo");
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi cập nhật: " + ex.Message); }
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            // 1. Kiểm tra các ô bắt buộc
            if (string.IsNullOrEmpty(txtMaSP.Text) || string.IsNullOrEmpty(txtTenSP.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ Mã và Tên sản phẩm!");
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    conn.Open();
                    // 2. Viết lại câu lệnh SQL để khớp với tất cả các ô trên giao diện của bạn
                    string sql = @"INSERT INTO SanPham 
                           ([Mã sản phẩm], [Tên sản phẩm], [Thương hiệu], [Giá thành (VNĐ)], 
                            [Thông số kỹ thuật], [Tồn kho], [Bảo hành], [Phân loại], HinhAnh) 
                           VALUES (@ma, @ten, @th, @gia, @ts, @tk, @bh, @loai, @anh)";

                    SqlCommand cmd = new SqlCommand(sql, conn);

                    // Gán dữ liệu từ tất cả các ô nhập liệu vào tham số
                    cmd.Parameters.AddWithValue("@ma", txtMaSP.Text);
                    cmd.Parameters.AddWithValue("@ten", txtTenSP.Text);
                    cmd.Parameters.AddWithValue("@th", txtThuongHieu.Text); // Thêm Thương hiệu
                    cmd.Parameters.AddWithValue("@gia", txtGia.Text);
                    cmd.Parameters.AddWithValue("@ts", txtThongSo.Text);
                    cmd.Parameters.AddWithValue("@tk", nmTonKho.Value);
                    cmd.Parameters.AddWithValue("@bh", txtBaoHanh.Text);    // Thêm Bảo hành
                    cmd.Parameters.AddWithValue("@loai", cboPhanLoai.Text); // Thêm Phân loại
                    cmd.Parameters.AddWithValue("@anh", txtPathAnh.Text);

                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Thêm sản phẩm mới thành công!");

                    // 3. Làm mới bảng và xóa trắng ô nhập
                    LoadData();
                    ClearInputs();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi thêm: " + ex.Message);
            }
        }
        void ClearInputs()
        {
            txtMaSP.Clear();
            txtTenSP.Clear();
            txtGia.Clear();
            txtThongSo.Clear();
            nmTonKho.Value = 0;
            txtPathAnh.Clear();
            picSanPham_QuanLy.Image = null; // Xóa ảnh hiển thị cũ
        }
        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void dgvSanPham_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 1. Kiểm tra để chắc chắn người dùng bấm vào dòng có dữ liệu, không phải tiêu đề
            if (e.RowIndex >= 0)
            {
                // 2. Lấy ra dòng (row) hiện tại
                DataGridViewRow row = dgvSanPham.Rows[e.RowIndex];

                // 3. Đổ dữ liệu lên các ô TextBox/NumericUpDown/ComboBox
                // Lưu ý: Dùng thêm dấu '?' để tránh lỗi nếu vô tình ô dữ liệu bị rỗng (null)
                txtMaSP.Text = row.Cells["Mã sản phẩm"].Value?.ToString();
                txtTenSP.Text = row.Cells["Tên sản phẩm"].Value?.ToString();
                txtThuongHieu.Text = row.Cells["Thương hiệu"].Value?.ToString();
                txtGia.Text = row.Cells["Giá thành (VNĐ)"].Value?.ToString();
                txtThongSo.Text = row.Cells["Thông số kỹ thuật"].Value?.ToString();
                txtBaoHanh.Text = row.Cells["Bảo hành"].Value?.ToString();

                // Ép kiểu cho NumericUpDown (Tồn kho)
                nmTonKho.Value = Convert.ToDecimal(row.Cells["Tồn kho"].Value ?? 0);

                // Chọn giá trị cho ComboBox (Phân loại)
                cboPhanLoai.Text = row.Cells["Phân loại"].Value?.ToString();

                // ==========================================
                // 4. HIỂN THỊ HÌNH ẢNH (Đã sửa sang Base64)
                // ==========================================

                // Lấy chuỗi base64 từ cột HinhAnh trong Database
                string chuoiAnh = row.Cells["HinhAnh"].Value?.ToString() ?? "";

                // Cất chuỗi vào ô ẩn để dùng khi nhấn nút Sửa
                txtPathAnh.Text = chuoiAnh;

                // Gọi hàm load ảnh Base64 an toàn để hiển thị lên PictureBox
                LoadImageToPictureBox(chuoiAnh, picSanPham_QuanLy);
            }
        }
   

        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            txtMaSP.Clear(); txtTenSP.Clear(); txtThuongHieu.Clear();
            nmTonKho.Value=0 ; txtGia.Clear(); txtThongSo.Clear();
            txtBaoHanh.Clear(); cboPhanLoai.SelectedIndex = -1;
            LoadData();
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Xóa linh kiện này?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    conn.Open();
                    string sql = "DELETE FROM SanPham WHERE [Mã sản phẩm]=@ma";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@ma", txtMaSP.Text);

                    int kq = cmd.ExecuteNonQuery(); // Kiểm tra xem có xóa được dòng nào không

                    if (kq > 0)
                    {
                        LoadData(); // Load lại bảng
                        ClearInputs(); // <-- THÊM DÒNG NÀY: Xóa trắng toàn bộ TextBox
                        MessageBox.Show("Đã xóa món hàng và làm mới ô nhập!");
                    }
                }
            }
        }

        private void UpdateMaSP(object sender, EventArgs e)
        {
            // Chỉ tự sinh mã khi bạn đang ở chế độ "Thêm mới" (ô mã không bị khóa)
            if (txtMaSP.ReadOnly == false)
            {
                string maMoi = TaoMaTheoThongSo();

                // Nếu mã tạo ra chỉ có dấu gạch ngang (do chưa nhập đủ thông tin)
                // thì mình có thể để mặc định là "DRAFT-XXXX" chẳng hạn, hoặc cứ hiện ra.
                txtMaSP.Text = maMoi;
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btnThemGioHang_Click(object sender, EventArgs e)
        {
            // 1. Kiểm tra xem người dùng đã chọn sản phẩm bên bảng tìm kiếm chưa
            if (string.IsNullOrEmpty(lblTenSP.Text) || lblTenSP.Text == "Tên sản phẩm: ")
            {
                MessageBox.Show("Vui lòng chọn một sản phẩm để thêm vào giỏ!", "Nhắc nhở");
                return;
            }

            // 2. Tự động bóc tách số lượng TỒN KHO từ nhãn (Ví dụ từ chữ "Tồn kho: 15" -> lấy số 15)
            string tonKhoRaw = lblThongSo.Text;
            string cleanTonKho = new string(tonKhoRaw.Where(char.IsDigit).ToArray());
            int tonKho = 0;
            int.TryParse(cleanTonKho, out tonKho);

            // 3. Lấy số lượng khách muốn mua
            int soLuong = (int)nmSoLuongMua.Value;

            // =========================================================
            // 4. CHẶN LỖI: KIỂM TRA SỐ LƯỢNG MUA VÀ TỒN KHO
            // =========================================================
            if (soLuong <= 0)
            {
                MessageBox.Show("Số lượng mua phải lớn hơn 0!", "Nhắc nhở");
                return; // Dừng hàm lại ngay, không cho thêm
            }

            if (soLuong > tonKho)
            {
                MessageBox.Show("Không đủ hàng! Trong kho hiện chỉ còn " + tonKho + " sản phẩm.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // Dừng hàm lại ngay, chặn xuất kho âm
            }
            // =========================================================

            // 5. Nếu qua được bước kiểm tra trên, tiến hành thêm vào giỏ hàng bình thường
            string giaRaw = lblGiaBan.Text;
            string cleanGia = new string(giaRaw.Where(char.IsDigit).ToArray());
            long giaBan = 0;

            if (long.TryParse(cleanGia, out giaBan))
            {
                long thanhTien = giaBan * soLuong;

                // Lấy mã SP từ bảng
                string maSP = dgvKetQuaTim.CurrentRow.Cells["Mã sản phẩm"].Value.ToString();

                // Thêm dòng mới vào lưới Giỏ hàng
                dgvGioHang.Rows.Add(maSP, lblTenSP.Text.Replace("Tên sản phẩm: ", ""), soLuong, giaBan, thanhTien);

                // Tính lại tiền
                TinhTongTien();
            }
            else
            {
                MessageBox.Show("Không thể lấy giá bán từ nhãn!");
            }
        }
        void TinhTongTien()
        {
            long tong = 0;
            foreach (DataGridViewRow row in dgvGioHang.Rows)
            {
                if (row.Cells[4].Value != null) // Cột Thành tiền (index 4)
                {
                    tong += Convert.ToInt64(row.Cells[4].Value);
                }
            }
            // Hiển thị lại lên nhãn với định dạng dấu chấm cho đẹp
            lblTongTien.Text = tong.ToString("N0") + " VNĐ";
        }

        private void btnXoaMon_Click(object sender, EventArgs e)
        {
            if (dgvGioHang.CurrentRow != null && dgvGioHang.CurrentRow.Index != -1)
            {
                // 2. Hỏi xác nhận lại cho chắc (tránh bấm nhầm)
                DialogResult dr = MessageBox.Show("Bạn có chắc muốn xóa món này khỏi giỏ hàng?",
                                                 "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dr == DialogResult.Yes)
                {
                    // 3. Thực hiện xóa dòng đang chọn
                    dgvGioHang.Rows.RemoveAt(dgvGioHang.CurrentRow.Index);

                    // 4. QUAN TRỌNG: Gọi lại hàm tính tổng tiền để cập nhật lại con số đúng
                    TinhTongTien();

                    MessageBox.Show("Đã xóa món hàng!");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn món hàng cần xóa trong giỏ!");
            }
        }

        private void btnChonAnh_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif";

            if (open.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // Đọc toàn bộ file ảnh thành mảng Byte, sau đó ép sang Base64
                    byte[] imageBytes = File.ReadAllBytes(open.FileName);

                    // Nối thêm header để giống hệt định dạng của Web
                    string base64String = "data:image/jpeg;base64," + Convert.ToBase64String(imageBytes);

                    // Lưu chuỗi này vào TextBox ẩn để dùng cho việc Thêm/Sửa vào Database
                    txtPathAnh.Text = base64String;

                    // Hiển thị ngay lên PictureBox bằng hàm vừa sửa
                    LoadImageToPictureBox(base64String, picSanPham_QuanLy);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi chuyển đổi file ảnh: " + ex.Message);
                }
            }
        }
        private void LoadImageToPictureBox(string chuoiBase64, PictureBox pic)
        {
            try
            {
                // Giải phóng ảnh cũ để tránh rò rỉ bộ nhớ
                if (pic.Image != null) pic.Image.Dispose();

                // Kiểm tra xem chuỗi có hợp lệ không (có định dạng data:image...)
                if (!string.IsNullOrEmpty(chuoiBase64) && chuoiBase64.Contains(","))
                {
                    // Cắt lấy phần mã thực sự nằm sau dấu phẩy
                    string base64Data = chuoiBase64.Split(',')[1];

                    // Dịch mã thành mảng Byte và load lên hình
                    byte[] imageBytes = Convert.FromBase64String(base64Data);
                    using (MemoryStream ms = new MemoryStream(imageBytes))
                    {
                        pic.Image = Image.FromStream(ms);
                    }
                    pic.SizeMode = PictureBoxSizeMode.Zoom;
                }
                else
                {
                    pic.Image = null; // Bỏ trống nếu sản phẩm chưa có ảnh
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi nạp ảnh Base64: " + ex.Message);
                pic.Image = null;
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

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

        private void dgvSanPham_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // 1. Kiểm tra để chắc chắn người dùng bấm vào dòng có dữ liệu, không phải tiêu đề
            if (e.RowIndex >= 0)
            {
                // 2. Lấy ra dòng (row) hiện tại
                DataGridViewRow row = dgvSanPham.Rows[e.RowIndex];

                // 3. Đổ dữ liệu lên các ô TextBox/NumericUpDown/ComboBox
                // Lưu ý: Dùng thêm dấu '?' để tránh lỗi nếu vô tình ô dữ liệu bị rỗng (null)
                txtMaSP.Text = row.Cells["Mã sản phẩm"].Value?.ToString();
                txtTenSP.Text = row.Cells["Tên sản phẩm"].Value?.ToString();
                txtThuongHieu.Text = row.Cells["Thương hiệu"].Value?.ToString();
                txtGia.Text = row.Cells["Giá thành (VNĐ)"].Value?.ToString();
                txtThongSo.Text = row.Cells["Thông số kỹ thuật"].Value?.ToString();
                txtBaoHanh.Text = row.Cells["Bảo hành"].Value?.ToString();

                // Ép kiểu cho NumericUpDown (Tồn kho)
                nmTonKho.Value = Convert.ToDecimal(row.Cells["Tồn kho"].Value ?? 0);

                // Chọn giá trị cho ComboBox (Phân loại)
                cboPhanLoai.Text = row.Cells["Phân loại"].Value?.ToString();

                // ==========================================
                // 4. HIỂN THỊ HÌNH ẢNH (Đã sửa sang Base64)
                // ==========================================

                // Lấy chuỗi base64 từ cột HinhAnh trong Database
                string chuoiAnh = row.Cells["HinhAnh"].Value?.ToString() ?? "";

                // Cất chuỗi vào ô ẩn để dùng khi nhấn nút Sửa
                txtPathAnh.Text = chuoiAnh;

                // Gọi hàm load ảnh Base64 an toàn để hiển thị lên PictureBox
                LoadImageToPictureBox(chuoiAnh, picSanPham_QuanLy);
            }
        }

        private void txtTimKiemBanHang_TextChanged(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    conn.Open();
                    // Lệnh SQL: Tìm những sản phẩm có Mã hoặc Tên chứa từ khóa vừa gõ
                    string sql = "SELECT [Mã sản phẩm], [Tên sản phẩm], [Giá thành (VNĐ)], [Tồn kho], HinhAnh " +
                                 "FROM SanPham WHERE [Mã sản phẩm] LIKE @key OR [Tên sản phẩm] LIKE @key";

                    SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                    // Dùng dấu % để tìm kiếm kiểu "chứa ký tự", gõ '24UQ' là nó ra cả mã đầy đủ
                    da.SelectCommand.Parameters.AddWithValue("@key", "%" + txtTimKiemBanHang.Text + "%");

                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    // Đổ dữ liệu vào cái bảng xám xám ở trên
                    dgvKetQuaTim.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                // Để trống hoặc MessageBox.Show(ex.Message);
            }
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    conn.Open();
                    // Sử dụng LIKE và dấu % để tìm kiếm gần đúng
                    string sql = "SELECT * FROM SanPham WHERE [Mã sản phẩm] LIKE @key OR [Tên sản phẩm] LIKE @key";

                    SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                    // Thêm dấu % vào trước và sau từ khóa để tìm kiếm "chứa trong"
                    da.SelectCommand.Parameters.AddWithValue("@key", "%" + txtSearch.Text + "%");

                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    // Cập nhật lại nguồn dữ liệu cho bảng
                    dgvSanPham.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                // Tránh hiện MessageBox liên tục khi gõ, chỉ cần log lỗi nếu cần
            }
        }

        private void btnThanhToan_Click(object sender, EventArgs e)
        {
            // 1. Kiểm tra giỏ hàng có đồ hay không
            if (dgvGioHang.Rows.Count == 0 || (dgvGioHang.Rows.Count == 1 && dgvGioHang.Rows[0].IsNewRow))
            {
                MessageBox.Show("Giỏ hàng đang trống, vui lòng thêm sản phẩm!", "Thông báo");
                return;
            }

            // 2. BẮT LỖI THÔNG TIN KHÁCH HÀNG
            if (string.IsNullOrEmpty(txtTenKH_BanHang.Text) || string.IsNullOrEmpty(txtSDT_BanHang.Text))
            {
                MessageBox.Show("Vui lòng nhập Tên và Số điện thoại khách hàng trước khi thanh toán!", "Nhắc nhở");
                txtTenKH_BanHang.Focus();
                return;
            }

            // Lấy tên nhân viên
            string tenNV = lblUserHienThi.Text.Replace("Xin chào: ", "").Trim();

            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                conn.Open();
                SqlTransaction tran = conn.BeginTransaction();

                try
                {
                    // ==========================================================
                    // BƯỚC A: XỬ LÝ KHÁCH HÀNG (Tìm khách cũ hoặc tạo khách mới)
                    // ==========================================================
                    string sdt = txtSDT_BanHang.Text.Trim();
                    string tenKH = txtTenKH_BanHang.Text.Trim();
                    int maKH = 0;

                    // Kiểm tra xem khách đã từng mua chưa
                    string sqlCheckKH = "SELECT MaKH FROM KhachHang WHERE SDT = @sdt";
                    SqlCommand cmdCheck = new SqlCommand(sqlCheckKH, conn, tran);
                    cmdCheck.Parameters.AddWithValue("@sdt", sdt);
                    object resultKH = cmdCheck.ExecuteScalar();

                    if (resultKH != null)
                    {
                        maKH = Convert.ToInt32(resultKH); // Khách cũ, lấy mã
                    }
                    else
                    {
                        // Khách mới: Thêm vào và lấy mã tự tăng (OUTPUT INSERTED.MaKH)
                        string sqlInsertKH = "INSERT INTO KhachHang (TenKH, SDT, NgayMua) OUTPUT INSERTED.MaKH VALUES (@ten, @sdt, @ngay)";
                        SqlCommand cmdInsertKH = new SqlCommand(sqlInsertKH, conn, tran);
                        cmdInsertKH.Parameters.AddWithValue("@ten", tenKH);
                        cmdInsertKH.Parameters.AddWithValue("@sdt", sdt);
                        cmdInsertKH.Parameters.AddWithValue("@ngay", DateTime.Now);
                        maKH = (int)cmdInsertKH.ExecuteScalar();
                    }

                    // ==========================================================
                    // BƯỚC B: TẠO HÓA ĐƠN (Đồng bộ với chuẩn Web)
                    // ==========================================================
                    // Dùng tiền tố POS để phân biệt đơn bán tại cửa hàng và đơn WEB
                    string maHD = "POS" + DateTime.Now.ToString("ddMMyyHHmmss");
                    string cleanTongTien = lblTongTien.Text.Replace(".", "").Replace("VNĐ", "").Replace(" ", "").Trim();
                    long tongTien = long.Parse(cleanTongTien);

                    // Thêm NguonDon và TrangThai để sau này báo cáo dễ dàng
                    string sqlHD = "INSERT INTO HoaDon (MaHD, NgayBan, TongTien, NhanVien, MaKH, NguonDon, TrangThai) " +
                                   "VALUES (@ma, @ngay, @tong, @nv, @makh, 'Tại cửa hàng', 'Hoàn thành')";
                    SqlCommand cmdHD = new SqlCommand(sqlHD, conn, tran);
                    cmdHD.Parameters.AddWithValue("@ma", maHD);
                    cmdHD.Parameters.AddWithValue("@ngay", DateTime.Now);
                    cmdHD.Parameters.AddWithValue("@tong", tongTien);
                    cmdHD.Parameters.AddWithValue("@nv", tenNV);
                    cmdHD.Parameters.AddWithValue("@makh", maKH);
                    cmdHD.ExecuteNonQuery();

                    // ==========================================================
                    // BƯỚC C: LƯU CHI TIẾT HÓA ĐƠN VÀ TRỪ KHO
                    // ==========================================================
                    foreach (DataGridViewRow row in dgvGioHang.Rows)
                    {
                        if (row.IsNewRow || row.Cells[0].Value == null) continue;

                        string maSP = row.Cells[0].Value.ToString();
                        int slMua = Convert.ToInt32(row.Cells[2].Value);
                        long gia = Convert.ToInt64(row.Cells[3].Value);
                        long tt = Convert.ToInt64(row.Cells[4].Value);

                        // Lưu Chi Tiết (Không còn bị dính chùm nữa)
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

                        // Trừ tồn kho
                        string sqlUp = "UPDATE SanPham SET [Tồn kho] = [Tồn kho] - @sl WHERE [Mã sản phẩm] = @masp";
                        SqlCommand cmdUp = new SqlCommand(sqlUp, conn, tran);
                        cmdUp.Parameters.AddWithValue("@sl", slMua);
                        cmdUp.Parameters.AddWithValue("@masp", maSP);
                        cmdUp.ExecuteNonQuery();
                    }

                    tran.Commit(); // Chốt giao dịch thành công

                    // 5. Hỏi xuất hóa đơn (Hàm in hóa đơn của bạn giữ nguyên, không cần sửa)
                    if (MessageBox.Show("Thanh toán thành công! Bạn có muốn xem hóa đơn không?", "Xác nhận",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        InHoaDonChinhChu();
                    }

                    // 6. Làm sạch giao diện sau khi bán
                    dgvGioHang.Rows.Clear();
                    lblTongTien.Text = "0 VNĐ";
                    LoadData();
                    HienThiKetQuaTimKiem("");
                    ResetTabBanHang();
                }
                catch (Exception ex)
                {
                    tran.Rollback(); // Hoàn tác nếu có lỗi
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
            // Thiết lập Font chữ
            Font fCuaHang = new Font("Arial", 16, FontStyle.Bold);
            Font fTieuDe = new Font("Arial", 18, FontStyle.Bold);
            Font fChu = new Font("Arial", 11);
            Font fDam = new Font("Arial", 11, FontStyle.Bold);
            Pen pen = new Pen(Color.Black, 1);

            int y = 40; // Tọa độ dòng bắt đầu
            int x = 80; // Lề trái

            // 1. THÔNG TIN CỬA HÀNG
            g.DrawString("HỒNG PHÚC COMPUTER", fCuaHang, Brushes.Blue, 280, y);
            y += 35;
            g.DrawString("HÓA ĐƠN BÁN HÀNG", fTieuDe, Brushes.Black, 300, y);
            y += 45;

            // 2. THÔNG TIN CHUNG (Mã HD & Ngày)
            string maHD = "HD" + DateTime.Now.ToString("ddMMyyHHmmss");
            g.DrawString("Mã HD: " + maHD, fChu, Brushes.Black, x, y);
            g.DrawString("Ngày: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm"), fChu, Brushes.Black, 500, y);
            y += 35;

            // 3. VẼ TIÊU ĐỀ BẢNG (Kẻ khung bao quanh)
            g.DrawRectangle(pen, x, y, 680, 30);
            g.DrawString("Tên Sản Phẩm", fDam, Brushes.Black, x + 5, y + 7);
            g.DrawString("SL", fDam, Brushes.Black, x + 380, y + 7);
            g.DrawString("Đơn Giá", fDam, Brushes.Black, x + 450, y + 7);
            g.DrawString("Thành Tiền", fDam, Brushes.Black, x + 560, y + 7);
            y += 30;

            // 4. DUYỆT GIỎ HÀNG VÀ IN CHI TIẾT
            foreach (DataGridViewRow row in dgvGioHang.Rows)
            {
                // Bỏ qua dòng trống cuối cùng
                if (row.IsNewRow || row.Cells[0].Value == null) continue;

                string ten = row.Cells[1].Value.ToString();
                string sl = row.Cells[2].Value.ToString();
                string gia = row.Cells[3].Value.ToString();
                string tt = row.Cells[4].Value.ToString();

                // Vẽ nội dung từng cột
                g.DrawString(ten, fChu, Brushes.Black, x + 5, y + 7);
                g.DrawString(sl, fChu, Brushes.Black, x + 380, y + 7);
                g.DrawString(gia, fChu, Brushes.Black, x + 450, y + 7);
                g.DrawString(tt, fChu, Brushes.Black, x + 560, y + 7);

                y += 30;
                // Kẻ đường gạch ngang sau mỗi món hàng cho dễ nhìn
                g.DrawLine(pen, x, y, x + 680, y);
            }

            // 5. TỔNG TIỀN
            y += 15;
            g.DrawString("TỔNG TIỀN THANH TOÁN:", fDam, Brushes.Black, x + 350, y);
            // In màu đỏ cho nổi bật tổng tiền
            g.DrawString(lblTongTien.Text, fDam, Brushes.Red, x + 560, y);

            // 6. PHẦN CHỮ KÝ (Lấy tên người đăng nhập)
            y += 60;
            // Tách tên nhân viên từ Label (Ví dụ: "Xin chào: admin" -> lấy chữ "admin")
            string nhanVien = lblUserHienThi.Text.Replace("Xin chào: ", "").Trim();

            g.DrawString("Người lập hóa đơn", fDam, Brushes.Black, x + 510, y);
            y += 20;
            g.DrawString("(Ký và ghi rõ họ tên)", new Font("Arial", 9, FontStyle.Italic), Brushes.Black, x + 515, y);

            y += 60; // Khoảng trống để ký tên
                     // In tên nhân viên ở dưới cùng chữ ký
            g.DrawString(nhanVien, fDam, Brushes.Black, x + 520, y);

            // Chân trang
            y += 50;
            g.DrawString("--- Cảm ơn quý khách và hẹn gặp lại! ---",
                         new Font("Arial", 10, FontStyle.Italic), Brushes.Gray, 280, y);
        }
        void HienThiKetQuaTimKiem(string keyword)
        {
            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                // Thêm cột [Thông số kỹ thuật] vào câu lệnh SELECT
                string sql = "SELECT [Mã sản phẩm], [Tên sản phẩm], [Giá thành (VNĐ)], [Tồn kho], [Thông số kỹ thuật], HinhAnh " +
                             "FROM SanPham WHERE [Tên sản phẩm] LIKE @key";
                SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                da.SelectCommand.Parameters.AddWithValue("@key", "%" + keyword + "%");
                DataTable dt = new DataTable();
                da.Fill(dt);
                dgvKetQuaTim.DataSource = dt;
            }
        }
        void ResetTabBanHang()
        {
            lblTenSP.Text = "Tên sản phẩm: ";
            lblGiaBan.Text = "Giá bán: ";
            lblThongSo.Text = "Tồn kho: ";
            nmSoLuongMua.Value = 0;
            picSanPham.Image = null;

            // Xóa trắng thông tin khách hàng cho đơn tiếp theo
            txtTenKH_BanHang.Clear();
            txtSDT_BanHang.Clear();
        }
        void LoadChamCong()
        {
            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                try
                {
                    conn.Open();
                    // Lấy toàn bộ lịch sử chấm công, sắp xếp cái mới nhất lên đầu
                    string sql = @"SELECT CC.Username AS [Tài khoản], 
                                  CC.ThoiGianVao AS [Giờ Vào], 
                                  CC.ThoiGianRa AS [Giờ Ra], 
                                  ROUND(CC.TongGio, 2) AS [Tổng Giờ] 
                           FROM ChamCong CC
                           ORDER BY CC.ThoiGianVao DESC";

                    SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    // Đổ dữ liệu vào bảng chấm công bên phải
                    dgvChamCong.DataSource = dt;
                    dgvChamCong.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                }
                catch (Exception ex)
                {
                    // Tránh báo lỗi nếu SQL chưa có bảng ChamCong
                    Console.WriteLine("Chưa có bảng Chấm Công: " + ex.Message);
                }
            }
        }
        void LoadNhanVien()
        {
            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                try
                {
                    conn.Open();
                    // Lấy các cột đúng theo bảng NhanVien của bạn
                    string sql = "SELECT Username, Password, TenNV, ChucVu, SDT, DiaChi FROM NhanVien";
                    SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    dgvNhanVien.DataSource = dt;

                    // THIẾT KẾ KHÍT: Dàn đều các cột ra hết chiều ngang bảng
                    dgvNhanVien.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                    // Đổi tên tiêu đề cột cho đẹp (Tiếng Việt có dấu)
                    dgvNhanVien.Columns["Username"].HeaderText = "Tài khoản";
                    dgvNhanVien.Columns["Password"].HeaderText = "Mật khẩu";
                    dgvNhanVien.Columns["TenNV"].HeaderText = "Họ và Tên";
                    dgvNhanVien.Columns["ChucVu"].HeaderText = "Chức vụ";
                    dgvNhanVien.Columns["SDT"].HeaderText = "Số điện thoại";
                    dgvNhanVien.Columns["DiaChi"].HeaderText = "Địa chỉ";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi load nhân viên: " + ex.Message);
                }
            }
        }
        private void btnLuuAnh_Click(object sender, EventArgs e)
        {
            // 1. Kiểm tra xem đã có Mã sản phẩm chưa (phải biết lưu cho ai chứ)
            if (string.IsNullOrEmpty(txtMaSP.Text))
            {
                MessageBox.Show("Vui lòng chọn một sản phẩm từ bảng trước khi lưu ảnh!");
                return;
            }

            // 2. Kiểm tra xem đã chọn ảnh mới chưa (ô txtPathAnh phải có tên file)
            if (string.IsNullOrEmpty(txtPathAnh.Text))
            {
                MessageBox.Show("Vui lòng nhấn 'Chọn ảnh' để chọn hình trước!");
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    conn.Open();
                    // 3. Câu lệnh SQL chỉ cập nhật duy nhất cột HinhAnh cho đúng Mã SP đó
                    string sql = "UPDATE SanPham SET HinhAnh = @anh WHERE [Mã sản phẩm] = @ma";

                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@anh", txtPathAnh.Text); // Tên file từ txtPathAnh
                    cmd.Parameters.AddWithValue("@ma", txtMaSP.Text);    // Mã sản phẩm đang hiện

                    int rows = cmd.ExecuteNonQuery();

                    if (rows > 0)
                    {
                        MessageBox.Show("Đã lưu ảnh cho sản phẩm: " + txtMaSP.Text, "Thành công");
                        LoadData(); // Load lại bảng để thấy tên file ảnh xuất hiện trong lưới
                    }
                    else
                    {
                        MessageBox.Show("Không tìm thấy sản phẩm này để cập nhật!");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu ảnh: " + ex.Message);
            }
           
        }

        private void dgvNhanVien_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Đảm bảo người dùng click vào một dòng hợp lệ (không click vào thanh tiêu đề)
            if (e.RowIndex >= 0)
            {
                // =========================================================
                // PHẦN 1: ĐỔ DỮ LIỆU TỪ BẢNG LÊN CÁC Ô NHẬP LIỆU BÊN TRÊN
                // =========================================================
                DataGridViewRow row = dgvNhanVien.Rows[e.RowIndex];

                txtUsername.Text = row.Cells["Username"].Value?.ToString();
                txtPassword.Text = row.Cells["Password"].Value?.ToString();
                txtTenNV.Text = row.Cells["TenNV"].Value?.ToString();
                txtSDT.Text = row.Cells["SDT"].Value?.ToString();
                txtDiaChi.Text = row.Cells["DiaChi"].Value?.ToString();
                cboChucVu.Text = row.Cells["ChucVu"].Value?.ToString();

                // Khóa ô Username lại vì đây là khóa chính (ID đăng nhập), không được phép sửa
                txtUsername.ReadOnly = true;


                // =========================================================
                // PHẦN 2: TỰ ĐỘNG LỌC LỊCH SỬ CHẤM CÔNG CỦA NHÂN VIÊN VỪA CHỌN
                // =========================================================
                string userSelected = row.Cells["Username"].Value?.ToString();

                try
                {
                    using (SqlConnection conn = new SqlConnection(cnStr))
                    {
                        conn.Open();
                        // Lấy giờ vào, giờ ra và tổng giờ từ bảng ChamCong theo Username
                        string sql = @"SELECT ThoiGianVao AS [Giờ Vào], 
                                      ThoiGianRa AS [Giờ Ra], 
                                      ROUND(TongGio, 2) AS [Tổng Giờ] 
                               FROM ChamCong 
                               WHERE Username = @user 
                               ORDER BY ThoiGianVao DESC";

                        SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                        da.SelectCommand.Parameters.AddWithValue("@user", userSelected);

                        DataTable dt = new DataTable();
                        da.Fill(dt);

                        // Đổ dữ liệu vào bảng chấm công bên phải
                        dgvChamCong.DataSource = dt;
                    }
                }
                catch (Exception ex)
                {
                    // Nếu chưa có bảng ChamCong hoặc lỗi mạng thì chỉ in ra màn hình Output để khỏi phiền người dùng
                    Console.WriteLine("Lỗi khi lọc lịch sử chấm công riêng: " + ex.Message);
                }
            }
        }

        private void btnThemNV_Click(object sender, EventArgs e)
        {
            // Kiểm tra dữ liệu đầu vào cơ bản
            if (string.IsNullOrEmpty(txtUsername.Text) || string.IsNullOrEmpty(txtPassword.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ Tài khoản và Mật khẩu!");
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    conn.Open();
                    // Câu lệnh SQL thêm mới
                    string sql = "INSERT INTO NhanVien (Username, Password, TenNV, ChucVu, SDT, DiaChi) " +
                                 "VALUES (@user, @pass, @ten, @cv, @sdt, @dc)";

                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@user", txtUsername.Text.Trim());
                    cmd.Parameters.AddWithValue("@pass", txtPassword.Text.Trim());
                    cmd.Parameters.AddWithValue("@ten", txtTenNV.Text.Trim());
                    cmd.Parameters.AddWithValue("@cv", cboChucVu.Text);
                    cmd.Parameters.AddWithValue("@sdt", txtSDT.Text.Trim());
                    cmd.Parameters.AddWithValue("@dc", txtDiaChi.Text.Trim());

                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Thêm nhân viên: " + txtUsername.Text + " thành công!", "Thông báo");

                    LoadNhanVien(); // Load lại bảng để thấy nhân viên mới
                    btnLammoiNV_Click(sender, e); // Xóa trắng các ô nhập
                }
            }
            catch (Exception ex)
            {
                // Lỗi thường gặp: Trùng Username (Khóa chính)
                MessageBox.Show("Lỗi: Tài khoản này đã tồn tại hoặc có lỗi hệ thống! \nChi tiết: " + ex.Message);
            }
        }

        private void btnSuaNV_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtUsername.Text)) return;

            try
            {
                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    conn.Open();
                    // Cập nhật thông tin (Không cho sửa Username vì là khóa chính)
                    string sql = "UPDATE NhanVien SET Password=@pass, TenNV=@ten, ChucVu=@cv, SDT=@sdt, DiaChi=@dc " +
                                 "WHERE Username=@user";

                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@user", txtUsername.Text.Trim());
                    cmd.Parameters.AddWithValue("@pass", txtPassword.Text.Trim());
                    cmd.Parameters.AddWithValue("@ten", txtTenNV.Text.Trim());
                    cmd.Parameters.AddWithValue("@cv", cboChucVu.Text);
                    cmd.Parameters.AddWithValue("@sdt", txtSDT.Text.Trim());
                    cmd.Parameters.AddWithValue("@dc", txtDiaChi.Text.Trim());

                    int kq = cmd.ExecuteNonQuery();
                    if (kq > 0)
                    {
                        MessageBox.Show("Cập nhật thông tin thành công!");
                        LoadNhanVien();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi sửa: " + ex.Message);
            }
        }

        private void btnXoaNV_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtUsername.Text)) return;

            // Không cho phép xóa admin gốc để tránh lỗi hệ thống
            if (txtUsername.Text.ToLower() == "admin")
            {
                MessageBox.Show("Không thể xóa tài khoản Admin hệ thống!");
                return;
            }

            DialogResult dr = MessageBox.Show("Bạn có chắc chắn muốn xóa nhân viên " + txtUsername.Text + "?",
                                             "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (dr == DialogResult.Yes)
            {
                try
                {
                    using (SqlConnection conn = new SqlConnection(cnStr))
                    {
                        conn.Open();
                        string sql = "DELETE FROM NhanVien WHERE Username=@user";
                        SqlCommand cmd = new SqlCommand(sql, conn);
                        cmd.Parameters.AddWithValue("@user", txtUsername.Text.Trim());

                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Đã xóa nhân viên!");
                        LoadNhanVien();
                        btnLammoiNV_Click(sender, e);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa: Nhân viên này có thể đã liên quan đến các hóa đơn cũ, không thể xóa! \nChi tiết: " + ex.Message);
                }
            }
        }

        private void btnLammoiNV_Click(object sender, EventArgs e)
        {
            txtUsername.Clear();
            txtPassword.Clear();
            txtTenNV.Clear();
            txtSDT.Clear();
            txtDiaChi.Clear();
            cboChucVu.SelectedIndex = -1; // Bỏ chọn ComboBox
            txtUsername.ReadOnly = false; // Cho phép nhập Username mới
            txtUsername.Focus();          // Đưa con trỏ vào ô Username
            LoadChamCong();
        }
        void LoadKhachHang()
        {
            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                try
                {
                    conn.Open();
                    string sql = @"
                        SELECT CAST(MaKH AS VARCHAR) AS [Mã KH], 
                               TenKH AS [Họ Tên], 
                               SDT AS [Số Điện Thoại], 
                               ISNULL(MaSP, N'Đơn Web (Nhiều món)') AS [Sản Phẩm], 
                               NgayMua AS [Ngày Mua], 
                               ISNULL(ThoiHanBH, 0) AS [BH (Tháng)], 
                               NgayMua AS [Hết Hạn] 
                        FROM KhachHang 

                        UNION ALL 

                        SELECT 'WEB-' + CAST(Id AS VARCHAR) AS [Mã KH], 
                               HoTen AS [Họ Tên], 
                               SoDienThoai AS [Số Điện Thoại], 
                               DanhSachSanPham AS [Sản Phẩm], 
                               NgayDat AS [Ngày Mua], 
                               0 AS [BH (Tháng)], 
                               NgayDat AS [Hết Hạn] 
                        FROM ChiTietKhachHang 

                        ORDER BY [Ngày Mua] DESC";

                    SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    // THỦ THUẬT: Lấy tên sản phẩm đi tra cứu ngược trong kho
                    foreach (DataRow row in dt.Rows)
                    {
                        string maKH = row["Mã KH"].ToString();
                        DateTime ngayMua = Convert.ToDateTime(row["Ngày Mua"]);
                        int bhThang = 0;

                        if (maKH.StartsWith("WEB-"))
                        {
                            string spStr = row["Sản Phẩm"].ToString();
                            string tenSPLookup = "";

                            // Cắt lấy Tên sản phẩm (Bỏ chuỗi số lượng " (x...)" và giá tiền đi)
                            int pos = spStr.IndexOf(" (x");
                            if (pos == -1) pos = spStr.IndexOf(" x");
                            if (pos == -1) pos = spStr.IndexOf("=");

                            if (pos > 0)
                            {
                                tenSPLookup = spStr.Substring(0, pos).Trim();

                                // Chạy vào kho tìm số bảo hành của món này
                                string sqlLookup = "SELECT [Bảo hành] FROM SanPham WHERE [Tên sản phẩm] LIKE @ten";
                                using (SqlCommand cmdLookup = new SqlCommand(sqlLookup, conn))
                                {
                                    // Dùng LIKE để tìm tương đối cho an toàn
                                    cmdLookup.Parameters.AddWithValue("@ten", "%" + tenSPLookup + "%");
                                    object kq = cmdLookup.ExecuteScalar();

                                    if (kq != null && kq != DBNull.Value)
                                    {
                                        string bhStr = kq.ToString(); // VD: "36 Tháng"
                                        // Bóc tách lấy đúng con số
                                        string numStr = new string(bhStr.Where(char.IsDigit).ToArray());
                                        if (!string.IsNullOrEmpty(numStr))
                                        {
                                            bhThang = int.Parse(numStr);
                                        }
                                    }
                                }
                            }
                            row["BH (Tháng)"] = bhThang;
                        }
                        else
                        {
                            bhThang = Convert.ToInt32(row["BH (Tháng)"]);
                        }

                        row["Hết Hạn"] = ngayMua.AddMonths(bhThang);
                    }

                    dgvKhachHang.DataSource = dt;
                    dgvKhachHang.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                }
                catch (Exception ex) { Console.WriteLine("Lỗi Load KH: " + ex.Message); }
            }
        }
        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                // Kiểm tra xem biến tenNguoiDung có dữ liệu không (đề phòng mở app lên tắt ngay chưa kịp đăng nhập)
                if (!string.IsNullOrEmpty(tenNguoiDung))
                {
                    using (SqlConnection conn = new SqlConnection(cnStr))
                    {
                        conn.Open();

                        // Cập nhật giờ ra và tính tổng giờ (DATEDIFF lấy số giây chia 3600 để ra số giờ)
                        // Điều kiện ThoiGianRa IS NULL đảm bảo nó chỉ tính cho đúng cái ca đang làm hiện tại
                        string sql = @"UPDATE ChamCong 
                               SET ThoiGianRa = @ra, 
                                   TongGio = CAST(DATEDIFF(SECOND, ThoiGianVao, @ra) AS float) / 3600.0 
                               WHERE Username = @user AND ThoiGianRa IS NULL";

                        SqlCommand cmd = new SqlCommand(sql, conn);
                        cmd.Parameters.AddWithValue("@ra", DateTime.Now);
                        cmd.Parameters.AddWithValue("@user", tenNguoiDung);

                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                // Lỗi này cứ để âm thầm, không dùng MessageBox để tránh việc người dùng không tắt được App
                Console.WriteLine("Lỗi chấm công khi thoát: " + ex.Message);
            }
        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        // ==========================================================
        // HÀM CHẠY NGẦM: CỨ 10 GIÂY TỰ ĐỘNG LÀM MỚI (PHIÊN BẢN HOÀN HẢO)
        // ==========================================================
        private void TimerAutoLoad_Tick(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab != null && tabControl1.SelectedTab.Text == "Duyệt đơn Web")
            {
                // 1. Ghi nhớ Mã đơn hàng bạn đang bấm xem (nếu có)
                string maHD_DangXem = "";
                if (dgvDonWeb.CurrentRow != null && dgvDonWeb.CurrentRow.Index >= 0)
                {
                    maHD_DangXem = dgvDonWeb.CurrentRow.Cells["Mã Đơn"].Value?.ToString() ?? "";
                }

                // 2. Tải lại danh sách đơn để "hứng" đơn mới từ Web
                LoadDonWeb();

                // Cập nhật tiêu đề để bạn biết nó vẫn đang chạy
                this.Text = "CỬA HÀNG LINH KIỆN PC - [Đã cập nhật lúc: " + DateTime.Now.ToString("HH:mm:ss") + "]";

                // 3. Tìm và bôi xanh lại đúng cái đơn lúc nãy bạn đang xem
                bool timThay = false;
                if (!string.IsNullOrEmpty(maHD_DangXem))
                {
                    foreach (DataGridViewRow row in dgvDonWeb.Rows)
                    {
                        if (row.Cells["Mã Đơn"].Value?.ToString() == maHD_DangXem)
                        {
                            row.Selected = true;
                            dgvDonWeb.CurrentCell = row.Cells[0]; // Trả lại con trỏ chuột về đúng dòng đó
                            timThay = true;
                            break;
                        }
                    }
                }

                // 4. Nếu lúc nãy bạn chưa chọn gì, thì xóa viền xanh mặc định đi cho đẹp
                if (!timThay)
                {
                    dgvDonWeb.ClearSelection();
                    // Lưu ý: Không thêm dòng dgvChiTietDon.DataSource = null ở đây nữa
                    // để giữ nguyên danh sách món hàng bạn đang xem ở bảng dưới.
                }
            }
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
                else
                    nmThoiHanBH.Value = 0;

                TinhNgayHetHan();
            }
        }
        void TinhNgayHetHan()
        {
            DateTime ngayMua = dtpNgayMua.Value;
            int soThang = (int)nmThoiHanBH.Value; // Lấy con số tháng đã được bóc tách từ bảng lớn đẩy lên

            DateTime ngayHetHan = ngayMua.AddMonths(soThang);
            txtNgayHetHan.Text = ngayHetHan.ToString("dd/MM/yyyy");
        }
        private void btnSuaKH_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMaKH.Text)) return;
            if (txtMaKH.Text.StartsWith("WEB"))
            {
                MessageBox.Show("Đây là lịch sử khách hàng Online cũ. Không thể sửa thông tin này!", "Cảnh báo");
                return;
            }

            // ... (Đoạn code Update SQL cũ của bạn giữ nguyên ở đây) ...
            try
            {
                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    conn.Open();
                    string sql = "UPDATE KhachHang SET TenKH=@ten, SDT=@sdt, MaSP=@masp, NgayMua=@ngay, ThoiHanBH=@bh WHERE MaKH=@ma";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@ten", txtTenKH.Text);
                    cmd.Parameters.AddWithValue("@sdt", txtSDT_KH.Text);
                    cmd.Parameters.AddWithValue("@masp", txtMaSP_Mua.Text);
                    cmd.Parameters.AddWithValue("@ngay", dtpNgayMua.Value);
                    cmd.Parameters.AddWithValue("@bh", nmThoiHanBH.Value);
                    cmd.Parameters.AddWithValue("@ma", txtMaKH.Text);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Đã cập nhật thông tin khách hàng!");
                    LoadKhachHang();
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi sửa: " + ex.Message); }
        }

        private void btnXoaKH_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMaKH.Text)) return;
            if (txtMaKH.Text.StartsWith("WEB"))
            {
                MessageBox.Show("Đây là lịch sử khách hàng Online cũ. Không thể xóa!", "Cảnh báo");
                return;
            }

            // ... (Đoạn code Delete SQL cũ của bạn giữ nguyên ở đây) ...
            if (MessageBox.Show("Xóa lịch sử bảo hành của khách này?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    conn.Open();
                    string sql = "DELETE FROM KhachHang WHERE MaKH=@ma";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@ma", txtMaKH.Text);
                    cmd.ExecuteNonQuery();
                    LoadKhachHang();
                    btnLamMoiKH_Click(sender, e);
                }
            }
        }

        private void btnLamMoiKH_Click(object sender, EventArgs e)
        {
            txtMaKH.Clear(); txtTenKH.Clear(); txtSDT_KH.Clear(); txtMaSP_Mua.Clear();
            dtpNgayMua.Value = DateTime.Now; nmThoiHanBH.Value = 0; txtNgayHetHan.Clear();
            LoadKhachHang();
        }

        private void dtpNgayMua_ValueChanged(object sender, EventArgs e)
        {
            TinhNgayHetHan();
        }

        private void nmThoiHanBH_ValueChanged(object sender, EventArgs e)
        {
            TinhNgayHetHan();
        }

        private void txtTimKiemKH_TextChanged(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    conn.Open();
                    string sql = @"
                        SELECT CAST(MaKH AS VARCHAR) AS [Mã KH], TenKH AS [Họ Tên], SDT AS [Số Điện Thoại], ISNULL(MaSP, N'Đơn Web (Nhiều món)') AS [Sản Phẩm], NgayMua AS [Ngày Mua], ISNULL(ThoiHanBH, 0) AS [BH (Tháng)], NgayMua AS [Hết Hạn] 
                        FROM KhachHang 
                        WHERE TenKH LIKE @key OR SDT LIKE @key
                        
                        UNION ALL 
                        
                        SELECT 'WEB-' + CAST(Id AS VARCHAR) AS [Mã KH], HoTen AS [Họ Tên], SoDienThoai AS [Số Điện Thoại], DanhSachSanPham AS [Sản Phẩm], NgayDat AS [Ngày Mua], 0 AS [BH (Tháng)], NgayDat AS [Hết Hạn] 
                        FROM ChiTietKhachHang 
                        WHERE HoTen LIKE @key OR SoDienThoai LIKE @key
                        
                        ORDER BY [Ngày Mua] DESC";

                    SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                    da.SelectCommand.Parameters.AddWithValue("@key", "%" + txtTimKiemKH.Text.Trim() + "%");

                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    foreach (DataRow row in dt.Rows)
                    {
                        string maKH = row["Mã KH"].ToString();
                        DateTime ngayMua = Convert.ToDateTime(row["Ngày Mua"]);
                        int bhThang = 0;

                        if (maKH.StartsWith("WEB-"))
                        {
                            string spStr = row["Sản Phẩm"].ToString();
                            string tenSPLookup = "";
                            int pos = spStr.IndexOf(" (x");
                            if (pos == -1) pos = spStr.IndexOf(" x");
                            if (pos == -1) pos = spStr.IndexOf("=");

                            if (pos > 0)
                            {
                                tenSPLookup = spStr.Substring(0, pos).Trim();
                                string sqlLookup = "SELECT [Bảo hành] FROM SanPham WHERE [Tên sản phẩm] LIKE @ten";
                                using (SqlCommand cmdLookup = new SqlCommand(sqlLookup, conn))
                                {
                                    cmdLookup.Parameters.AddWithValue("@ten", "%" + tenSPLookup + "%");
                                    object kq = cmdLookup.ExecuteScalar();
                                    if (kq != null && kq != DBNull.Value)
                                    {
                                        string numStr = new string(kq.ToString().Where(char.IsDigit).ToArray());
                                        if (!string.IsNullOrEmpty(numStr)) bhThang = int.Parse(numStr);
                                    }
                                }
                            }
                            row["BH (Tháng)"] = bhThang;
                        }
                        else
                        {
                            bhThang = Convert.ToInt32(row["BH (Tháng)"]);
                        }
                        row["Hết Hạn"] = ngayMua.AddMonths(bhThang);
                    }
                    dgvKhachHang.DataSource = dt;
                }
            }
            catch (Exception ex) { Console.WriteLine("Lỗi khi tìm kiếm KH: " + ex.Message); }
        }
        

        private void lblTongSoDon_Click(object sender, EventArgs e)
        {

        }
        void LoadDonWeb()
        {
            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                try
                {
                    conn.Open();
                    // Lấy đơn Web, đưa đơn "Chờ xác nhận" lên ưu tiên trên cùng
                    string sql = @"SELECT H.MaHD AS [Mã Đơn], 
                                          K.TenKH AS [Khách Hàng], 
                                          K.SDT AS [SĐT], 
                                          H.NgayBan AS [Ngày Đặt], 
                                          H.TongTien AS [Tổng Tiền], 
                                          H.TrangThai AS [Trạng Thái]
                                   FROM HoaDon H
                                   INNER JOIN KhachHang K ON H.MaKH = K.MaKH
                                   WHERE H.NguonDon = 'Web Online'
                                   ORDER BY CASE WHEN H.TrangThai = N'Chờ xác nhận' THEN 1 ELSE 2 END, H.NgayBan DESC";
                    SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dgvDonWeb.DataSource = dt;
                    dgvDonWeb.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    dgvDonWeb.ClearSelection();
                }
                catch (Exception ex) { Console.WriteLine("Lỗi Load đơn web: " + ex.Message); }
            }
        }
        private void btnThongKe_Click(object sender, EventArgs e)
        {
            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                try
                {
                    conn.Open();

                    // 1. Lấy mốc thời gian Admin chọn
                    // (Lưu ý: Mốc 'Đến ngày' phải cộng thêm 1 ngày để lấy trọn vẹn dữ liệu của ngày hôm đó)
                    DateTime tuNgay = dtpTuNgay.Value.Date;
                    DateTime denNgay = dtpDenNgay.Value.Date.AddDays(1);

                    // 2. Tải danh sách Hóa Đơn trong khoảng thời gian này
                    string sql = @"SELECT MaHD AS [Mã Hóa Đơn], 
                                  NgayBan AS [Ngày Bán], 
                                  TongTien AS [Tổng Tiền (VNĐ)], 
                                  NhanVien AS [Nhân Viên Bán] 
                           FROM HoaDon 
                           WHERE NgayBan >= @tuNgay AND NgayBan < @denNgay 
                           ORDER BY NgayBan DESC";

                    SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                    da.SelectCommand.Parameters.AddWithValue("@tuNgay", tuNgay);
                    da.SelectCommand.Parameters.AddWithValue("@denNgay", denNgay);

                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    // Đổ lên DataGridView
                    dgvDoanhThu.DataSource = dt;
                    dgvDoanhThu.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                    // 3. Tính toán Tổng Doanh Thu và Tổng Số Đơn từ cái bảng vừa tải về
                    long tongTien = 0;
                    int soDon = dt.Rows.Count; // Số dòng chính là số đơn hàng

                    foreach (DataRow row in dt.Rows)
                    {
                        tongTien += Convert.ToInt64(row["Tổng Tiền (VNĐ)"]);
                    }

                    // 4. Hiển thị con số lên Label với định dạng có dấu phẩy cho đẹp
                    lblTongDoanhThu.Text = tongTien.ToString("N0") + " VNĐ";
                    lblTongSoDon.Text = soDon.ToString() + " đơn";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi thống kê doanh thu: " + ex.Message);
                }
            }
        }

        private void btnDuyetDon_Click(object sender, EventArgs e)
        {
            if (dgvDonWeb.CurrentRow != null && dgvDonWeb.CurrentRow.Index >= 0)
            {
                string maHD = dgvDonWeb.CurrentRow.Cells["Mã Đơn"].Value.ToString();
                string trangThai = dgvDonWeb.CurrentRow.Cells["Trạng Thái"].Value.ToString();

                // Nếu đơn đã duyệt rồi thì chặn lại không cho trừ kho nữa
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
                            // A. Cập nhật trạng thái thành Hoàn thành
                            string sqlUpd = "UPDATE HoaDon SET TrangThai = N'Hoàn thành' WHERE MaHD = @ma";
                            SqlCommand cmdUpd = new SqlCommand(sqlUpd, conn, tran);
                            cmdUpd.Parameters.AddWithValue("@ma", maHD);
                            cmdUpd.ExecuteNonQuery();

                            // B. Trừ tồn kho liên hoàn cho tất cả món trong đơn
                            string sqlTruKho = @"UPDATE SanPham 
                                                 SET [Tồn kho] = [Tồn kho] - C.SoLuong
                                                 FROM SanPham S
                                                 INNER JOIN ChiTietHoaDon C ON S.[Mã sản phẩm] = C.MaSP
                                                 WHERE C.MaHD = @ma";
                            SqlCommand cmdTruKho = new SqlCommand(sqlTruKho, conn, tran);
                            cmdTruKho.Parameters.AddWithValue("@ma", maHD);
                            cmdTruKho.ExecuteNonQuery();

                            tran.Commit();

                            MessageBox.Show("Đã duyệt đơn và xuất kho thành công!", "Hoàn tất");

                            // Cập nhật lại giao diện ngay lập tức
                            LoadDonWeb();
                            dgvChiTietDon.DataSource = null; // Xóa trắng bảng bên phải
                            LoadData(); // Load lại kho để thấy số lượng đã giảm
                        }
                        catch (Exception ex)
                        {
                            tran.Rollback();
                            MessageBox.Show("Lỗi duyệt đơn: " + ex.Message);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn 1 đơn hàng bên danh sách để duyệt!");
            }
        }

        private void dgvDonWeb_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                string maHD = dgvDonWeb.Rows[e.RowIndex].Cells["Mã Đơn"].Value.ToString();

                using (SqlConnection conn = new SqlConnection(cnStr))
                {
                    try
                    {
                        conn.Open();
                        // Lấy danh sách sản phẩm khách đã mua trong đơn đó
                        string sql = @"SELECT S.[Tên sản phẩm] AS [Sản Phẩm], 
                                              C.SoLuong AS [Số Lượng], 
                                              C.DonGia AS [Đơn Giá], 
                                              C.ThanhTien AS [Thành Tiền]
                                       FROM ChiTietHoaDon C
                                       INNER JOIN SanPham S ON C.MaSP = S.[Mã sản phẩm]
                                       WHERE C.MaHD = @ma";
                        SqlCommand cmd = new SqlCommand(sql, conn);
                        cmd.Parameters.AddWithValue("@ma", maHD);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);

                        dgvChiTietDon.DataSource = dt;
                        dgvChiTietDon.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    }
                    catch (Exception ex) { Console.WriteLine("Lỗi Load chi tiết: " + ex.Message); }
                }
            }
        }
    }

}
