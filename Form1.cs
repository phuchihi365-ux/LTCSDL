using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace LTCSDL
{
    public partial class Form1 : Form
    {
        // Thay chuỗi kết nối này bằng của máy bạn
        string connectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=LTCSDL2;Integrated Security=True";

        public Form1()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen; // Hiện ở giữa màn hình
            txtPassword.PasswordChar = '●'; // Ẩn mật khẩu
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit(); // Thoát hẳn chương trình
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btnLogin_Click_1(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    // Truy vấn lấy Chức vụ dựa trên Username và Password
                    string sql = "SELECT ChucVu FROM NhanVien WHERE Username=@user AND Password=@pass";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@user", txtUsername.Text);
                    cmd.Parameters.AddWithValue("@pass", txtPassword.Text);

                    object result = cmd.ExecuteScalar();

                    if (result != null && result != DBNull.Value)
                    {
                        string chucVu = result.ToString();
                        string tenUser = txtUsername.Text; // Lấy tên người dùng

                        MessageBox.Show("Chào " + tenUser + " (" + chucVu + ")", "Thông báo");

                        // Mở Form Main với 2 tham số
                        Main f = new Main(chucVu, tenUser);
                        this.Hide();
                        f.ShowDialog();
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Tài khoản hoặc mật khẩu không đúng!", "Lỗi");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối: " + ex.Message);
            }
        }

        private void btnExit_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
