using System;
using System.Data;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Novacode;

namespace query
{
    public partial class Form1 : Form
    {
        DataTable dt = new DataTable();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ip=textBox1.Text.Trim();
            string user = textBox3.Text.Trim();
            string password = textBox4.Text.Trim();
            string db = textBox2.Text.Trim();
            string sql = textBox5.Text;
            login(ip, user, password, db,sql);
        }
        public void login(string ip, string user, string password, string db,string sql)
        {
            
            string connect="Data Source="+ip+";Initial Catalog="+db+";uid="+user+";pwd="+password+"";
            if (user == "")
            {
                connect = "Data Source=" + ip + ";Initial Catalog=" + db + ";Integrated Security=True";
            }
            SqlConnection sqlCnt = new SqlConnection(connect);
            try
            {
                sqlCnt.Open();
            }
            catch
            {
                MessageBox.Show("数据库连接失败");
                return ;
            }
            SqlCommand cmd = sqlCnt.CreateCommand();
            cmd.CommandText = sql;
            SqlDataReader dr = cmd.ExecuteReader();
            dt.Load(dr);
            dr.Close();
            sqlCnt.Close();

            putword(dt);
            MessageBox.Show("word 输出成功");
            System.Diagnostics.Process.Start(@"结果.docx");
        }
        public void putword (DataTable dt)
        { 
        
            int x, y;
            x = dt.Rows.Count;
            y = dt.Columns.Count;
            DocX doc = DocX.Create(@"结果.docx");
            doc.InsertTable(x,y);
            Table table = doc.Tables[0];
            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    table.Rows[i].Cells[j].Paragraphs[0].Append(dt.Rows[i][j].ToString());
                    table.Rows[i].Cells[j].Paragraphs[0].Alignment = Alignment.center;
                    table.Rows[i].Cells[j].VerticalAlignment = VerticalAlignment.Center;
                    table.Rows[i].Cells[j].Paragraphs[0].FontSize(8);
                }
            
            }
            doc.Save();
        }
    }
}
