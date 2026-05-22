
using System;
using System.Data;
using System.Data.SqlClient;

namespace LTCSDL
{
    public class DataProvider
    {
    
        private string cnStr = @"Data Source=HONGPHUC\SQLEXPRESS;Initial Catalog=LTCSDL;Integrated Security=True;TrustServerCertificate=True;";

        // 1. Hàm chạy lệnh SELECT (Lấy dữ liệu đổ vào DataGridView hoặc xử lý)
        public DataTable ExecuteQuery(string query, SqlParameter[] parameters = null)
        {
            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                try
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(query, conn);
                    if (parameters != null) cmd.Parameters.AddRange(parameters);

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                }
                catch (Exception ex) { Console.WriteLine("Lỗi ExecuteQuery: " + ex.Message); }
            }
            return dt;
        }

        // 2. Hàm lấy 1 giá trị duy nhất (Dùng tra cứu nhanh như ô Bảo Hành)
        public object ExecuteScalar(string query, SqlParameter[] parameters = null)
        {
            object result = null;
            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                try
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(query, conn);
                    if (parameters != null) cmd.Parameters.AddRange(parameters);

                    result = cmd.ExecuteScalar();
                }
                catch (Exception ex) { Console.WriteLine("Lỗi ExecuteScalar: " + ex.Message); }
            }
            return result;
        }

        // 3. Hàm Thêm, Sửa, Xóa (Để dành dùng cho các nút chức năng khác)
        public int ExecuteNonQuery(string query, SqlParameter[] parameters = null)
        {
            int result = 0;
            using (SqlConnection conn = new SqlConnection(cnStr))
            {
                try
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(query, conn);
                    if (parameters != null) cmd.Parameters.AddRange(parameters);

                    result = cmd.ExecuteNonQuery();
                }
                catch (Exception ex) { Console.WriteLine("Lỗi ExecuteNonQuery: " + ex.Message); }
            }
            return result;
        }
    }
}