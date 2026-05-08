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
        string cnStr = @"Data Source=.\SQLEXPRESS;Initial Catalog=LTCSDL2;Integrated Security=True";
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
            PhanQuyen();
            LoadData();
            lblUserHienThi.Text = "Xin chào: " + tenNguoiDung;

            // 1. Để Form hiện ra ở giữa màn hình
            this.StartPosition = FormStartPosition.CenterScreen;

            // 2. Ép các bảng phải dàn đều các cột cho khít khung
            dgvSanPham.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvKetQuaTim.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvGioHang.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // 3. (Tùy chọn) Khóa khung không cho người dùng kéo dãn làm lệch giao diện
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false; // Tắt nút phóng to để giữ nguyên tỉ lệ đẹp
        }
        void PhanQuyen()
        {
            // Kiểm tra xem tên biến có đúng là tabControl1 không (xem trong Properties)
            if (quyenNguoiDung == "Nhân viên")
            {
                // 1. Ẩn tab Quản lý đi, nhân viên chỉ được thấy tab Bán hàng
                if (tabControl1.TabPages.Contains(tabPage1))
                {
                    tabControl1.TabPages.Remove(tabPage1);
                }

                // 2. Chốt tab mặc định hiện lên là Bán hàng
                tabControl1.SelectedTab = tabPage2;

                this.Text = "CỬA HÀNG LINH KIỆN PC - [NHÂN VIÊN: " + DateTime.Now.ToShortDateString() + "]";
            }
            else // Là Admin
            {
                this.Text = "CỬA HÀNG LINH KIỆN PC - [QUẢN TRỊ VIÊN]";
                // Admin thấy hết, không cần xóa gì cả
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
                // Lưu ý: Tên trong ngoặc ["..."] phải khớp 100% với tên cột trong SQL của bạn
                txtMaSP.Text = row.Cells["Mã sản phẩm"].Value.ToString();
                txtTenSP.Text = row.Cells["Tên sản phẩm"].Value.ToString();
                txtThuongHieu.Text = row.Cells["Thương hiệu"].Value.ToString();
                txtGia.Text = row.Cells["Giá thành (VNĐ)"].Value.ToString();
                txtThongSo.Text = row.Cells["Thông số kỹ thuật"].Value.ToString();
                txtBaoHanh.Text = row.Cells["Bảo hành"].Value.ToString();

                // Ép kiểu cho NumericUpDown (Tồn kho)
                nmTonKho.Value = Convert.ToDecimal(row.Cells["Tồn kho"].Value);

                // Chọn giá trị cho ComboBox (Phân loại)
                cboPhanLoai.Text = row.Cells["Phân loại"].Value.ToString();

                // 4. HIỂN THỊ HÌNH ẢNH
                // Lấy tên file từ cột HinhAnh trong Database
                string tenFile = row.Cells["HinhAnh"].Value.ToString();
                txtPathAnh.Text = tenFile; // Cất tên file vào ô ẩn để dùng khi nhấn nút Sửa

                // Tạo đường dẫn đầy đủ đến thư mục Images của bạn
                string pathFull = Application.StartupPath + "\\Images\\" + tenFile;

                if (!string.IsNullOrEmpty(tenFile) && System.IO.File.Exists(pathFull))
                {
                    // Giải phóng ảnh cũ để không bị lỗi "file đang mở" khi bạn muốn chọn ảnh khác
                    if (picSanPham_QuanLy.Image != null) picSanPham_QuanLy.Image.Dispose();

                    picSanPham_QuanLy.Image = Image.FromFile(pathFull);
                    picSanPham_QuanLy.SizeMode = PictureBoxSizeMode.Zoom;
                }
                else
                {
                    picSanPham_QuanLy.Image = null; // Nếu không tìm thấy file thì xóa trắng khung ảnh
                }
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
            // 1. Lấy chuỗi từ Label (Ví dụ: "Giá bán: 2290000")
            string giaRaw = lblGiaBan.Text;

            // 2. Lọc bỏ tất cả chữ cái, chỉ giữ lại các con số
            string cleanGia = new string(giaRaw.Where(char.IsDigit).ToArray());

            long giaBan = 0;
            if (long.TryParse(cleanGia, out giaBan))
            {
                int soLuong = (int)nmSoLuongMua.Value;
                long thanhTien = giaBan * soLuong;

                // Thêm vào giỏ hàng (Giả sử dgvGioHang của bạn có các cột tương ứng)
                string maSP = dgvKetQuaTim.CurrentRow.Cells["Mã sản phẩm"].Value.ToString();
                dgvGioHang.Rows.Add(maSP, lblTenSP.Text.Replace("Tên sản phẩm: ", ""), soLuong, giaBan, thanhTien);

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
                string fileName = Path.GetFileName(open.FileName);
                string folderPath = Path.Combine(Application.StartupPath, "Images");
                string destPath = Path.Combine(folderPath, fileName);

                if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);

                // Copy file (ghi đè nếu trùng tên)
                File.Copy(open.FileName, destPath, true);

                // Hiển thị lên PictureBox bằng cách an toàn
                txtPathAnh.Text = fileName;
                LoadImageToPictureBox(fileName, picSanPham_QuanLy);
            }
        }
        private void LoadImageToPictureBox(string tenFile, PictureBox pic)
        {
            try
            {
                string pathFull = Path.Combine(Application.StartupPath, "Images", tenFile);

                // Kiểm tra nếu tên file thiếu đuôi thì bổ sung (xử lý tình huống trong video của bạn)
                if (!tenFile.ToLower().EndsWith(".jpg") && !tenFile.ToLower().EndsWith(".png"))
                {
                    if (File.Exists(pathFull + ".jpg")) pathFull += ".jpg";
                }

                if (File.Exists(pathFull))
                {
                    // Giải phóng ảnh cũ để tránh rò rỉ bộ nhớ
                    if (pic.Image != null) pic.Image.Dispose();

                    using (FileStream fs = new FileStream(pathFull, FileMode.Open, FileAccess.Read))
                    {
                        pic.Image = Image.FromStream(fs);
                    }
                    pic.SizeMode = PictureBoxSizeMode.Zoom;
                }
                else
                {
                    pic.Image = null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi nạp ảnh: " + ex.Message);
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

                string tenFile = row.Cells["HinhAnh"].Value?.ToString() ?? "";
                txtPathAnh.Text = tenFile;

                // Gọi hàm nạp ảnh an toàn
                LoadImageToPictureBox(tenFile, picSanPham_QuanLy);
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

            // Lấy tên nhân viên từ nhãn hiển thị (loại bỏ chữ "Xin chào: ")
            string tenNV = lblUserHienThi.Text.Replace("Xin chào: ", "").Trim();

            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                conn.Open();
                SqlTransaction tran = conn.BeginTransaction(); // Bắt đầu giao dịch để đảm bảo an toàn dữ liệu

                try
                {
                    // 2. TẠO HÓA ĐƠN MỚI
                    // 2. TẠO HÓA ĐƠN MỚI
                    string maHD = "HD" + DateTime.Now.ToString("ddMMyyHHmmss");

                    // --- ĐÂY LÀ ĐOẠN CẦN DÁN/SỬA ---
                    // Xóa sạch dấu chấm, chữ VNĐ và khoảng trắng để máy hiểu là số thuần túy
                    string cleanTongTien = lblTongTien.Text.Replace(".", "").Replace("VNĐ", "").Replace(" ", "").Trim();

                    // Ép kiểu số một cách an toàn
                    long tongTien = 0;
                    if (!long.TryParse(cleanTongTien, out tongTien))
                    {
                        MessageBox.Show("Số tiền không hợp lệ, vui lòng kiểm tra lại!");
                        return;
                    }
                    // -------------------------------

                    string sqlHD = "INSERT INTO HoaDon (MaHD, NgayBan, TongTien, NhanVien) VALUES (@ma, @ngay, @tong, @nv)";
                    SqlCommand cmdHD = new SqlCommand(sqlHD, conn, tran);
                    cmdHD.Parameters.AddWithValue("@ma", maHD);
                    cmdHD.Parameters.AddWithValue("@ngay", DateTime.Now);
                    cmdHD.Parameters.AddWithValue("@tong", tongTien);
                    cmdHD.Parameters.AddWithValue("@nv", tenNV);
                    cmdHD.ExecuteNonQuery();

                    // 3. DUYỆT GIỎ HÀNG ĐỂ LƯU CHI TIẾT VÀ TRỪ KHO
                    foreach (DataGridViewRow row in dgvGioHang.Rows)
                    {
                        if (row.IsNewRow || row.Cells[0].Value == null) continue;

                        string maSP = row.Cells[0].Value.ToString();
                        int slMua = Convert.ToInt32(row.Cells[2].Value);
                        long gia = Convert.ToInt64(row.Cells[3].Value);
                        long tt = Convert.ToInt64(row.Cells[4].Value);

                        // Lưu vào bảng ChiTietHoaDon (Có kèm tên nhân viên)
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

                        // Cập nhật giảm Tồn kho trong bảng SanPham
                        string sqlUp = "UPDATE SanPham SET [Tồn kho] = [Tồn kho] - @sl WHERE [Mã sản phẩm] = @masp";
                        SqlCommand cmdUp = new SqlCommand(sqlUp, conn, tran);
                        cmdUp.Parameters.AddWithValue("@sl", slMua);
                        cmdUp.Parameters.AddWithValue("@masp", maSP);
                        cmdUp.ExecuteNonQuery();
                    }

                    tran.Commit(); // Lưu mọi thay đổi vào SQL

                    // 4. HỎI XUẤT HÓA ĐƠN
                    if (MessageBox.Show("Thanh toán thành công! Bạn có muốn xem hóa đơn không?", "Xác nhận",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        InHoaDonChinhChu(); // Gọi hàm vẽ hóa đơn đã hướng dẫn trước đó
                    }

                    // 5. LÀM SẠCH GIAO DIỆN SAU KHI BÁN
                    dgvGioHang.Rows.Clear();
                    lblTongTien.Text = "0 VNĐ";
                    LoadData(); // Cập nhật lại bảng bên tab Quản lý
                    HienThiKetQuaTimKiem(""); // Cập nhật lại bảng tìm kiếm bên tab Bán hàng
                    ResetTabBanHang(); // Xóa các label tên sản phẩm, giá... đang hiện
                }
                catch (Exception ex)
                {
                    tran.Rollback(); // Nếu có bất kỳ lỗi nào, hủy bỏ toàn bộ thao tác để tránh sai sót dữ liệu
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
    }
}
