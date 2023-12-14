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
using excel = Microsoft.Office.Interop.Excel;
namespace TedarikKayit
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        SqlConnection conn = new SqlConnection("data source=(LocalDB)\\MSSQLLocalDB;attachdbfilename=|DataDirectory|\\database.mdf;integrated security=True;MultipleActiveResultSets=True;App=EntityFramework");
        DataTable dt = new DataTable();
        void veri()
        {
            SqlDataAdapter adapter = new SqlDataAdapter("SELECT *FROM tblurun", conn);
            DataTable tablo = new DataTable();
            conn.Open();
            adapter.Fill(tablo);
            dataGridView1.DataSource = tablo;
            conn.Close();
        }


        private void ExportToExcel(DataGridView dataGridView)// excele aktarma 
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            excel.Visible = true;

            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);

            Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                sheet1.Cells[1, i + 1] = dataGridView.Columns[i].HeaderText;
            }

            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView.Columns.Count; j++)
                {
                    sheet1.Cells[i + 2, j + 1] = dataGridView.Rows[i].Cells[j].Value.ToString();
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            veri();
            dataGridView1.AllowUserToAddRows = false;//kullanıcının mevcut kayıtları ekleyep ekleyemeceğini belirler belirler
            dataGridView1.AllowUserToDeleteRows = false;//kullanıcının mevcut kayıtları silip silemeyeceğini belirler 
            dataGridView1.ReadOnly = true;//yalnızca okunabilir
            dataGridView1.AllowUserToResizeRows = false;//kullanıcı sürunları yeniden boyutlandıramaz
            dataGridView1.AllowUserToResizeColumns = false;//kullanıcı sürunları yeniden boyutlandıramaz
            textBox1.Enabled = false;//labellere isim verilmesi

            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            maskedTextBox1.Text = "";//textlerin temizlenmesi

            maskedTextBox1.Text = DateTime.Today.ToString();//maskedtextboxa tarihin çekilmesi
        }

        private void btnekle_Click(object sender, EventArgs e)
        {
            try
            {
                string sorgu = "Insert into tblurun (alinanurun,urunfiyat,talepedenkisi,tarih) values (@urn,@fyt,@kisi,@tarih)";
                SqlCommand komut = new SqlCommand(sorgu, conn);
                komut.Parameters.AddWithValue("@urn", (textBox2.Text));
                komut.Parameters.AddWithValue("@fyt", Convert.ToDecimal(textBox3.Text));
                komut.Parameters.AddWithValue("@kisi", textBox4.Text);
                komut.Parameters.AddWithValue("@tarih", Convert.ToDateTime(maskedTextBox1.Text));
                conn.Open();
                komut.ExecuteNonQuery();
                conn.Close();
                veri();
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                maskedTextBox1.Text = "";
            }
            catch
            {
                MessageBox.Show("Herhangi Bir Veri Girişi Yapılması Gerekmektedir", "Hata!");
            }
        }

        private void btnsil_Click(object sender, EventArgs e)
        {
            try //Hata oluşabilecek kodlar 
            {
                string sorgu = "Delete From tblurun Where ID=@ID";
                SqlCommand komut = new SqlCommand(sorgu, conn);
                komut.Parameters.AddWithValue("@ID", dataGridView1.CurrentRow.Cells[0].Value);
                conn.Open();
                komut.ExecuteNonQuery();
                conn.Close();
                veri();
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                maskedTextBox1.Text = DateTime.Today.ToString();
            }
            catch //Hata oluştuğunda çalışacak kodlar 
            {
                MessageBox.Show("Silmek İçin Veri Seçilmesi Gerekmektedir", "Hata!");
            }
        }

        private void btngnclle_Click(object sender, EventArgs e)
        {
            try
            {
                string sorgu = "Update tblurun Set  alinanurun=@urun,urunfiyat=@fyt,talepedenkisi=@kisi,tarih=@tarih Where ID=@ıd";
                SqlCommand komut = new SqlCommand(sorgu, conn);
                komut.Parameters.AddWithValue("@ıd", textBox1.Text.ToString());
                komut.Parameters.AddWithValue("@urun", textBox2.Text);
                komut.Parameters.AddWithValue("@fyt", Convert.ToDecimal(textBox3.Text));
                komut.Parameters.AddWithValue("@kisi", textBox4.Text);
                komut.Parameters.AddWithValue("@tarih", Convert.ToDateTime(maskedTextBox1.Text));
                conn.Open();
                komut.ExecuteNonQuery();
                conn.Close();
                veri();
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                maskedTextBox1.Text = DateTime.Today.ToString();
            }
            catch
            {
                MessageBox.Show("Güncellemek İçin Herhangi Bir Veri Seçilmesi Gerekmektedir", "Hata!");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            maskedTextBox1.Text = DateTime.Today.ToString();
            MessageBox.Show("Temizlendi", "Bilgi");

        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {

            try
            {
                ExportToExcel(dataGridView1);
            }
            catch
            {
                MessageBox.Show("Veriler Başarıyla Excele Yazıldı");
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            dt.Clear();
            conn.Open();
            SqlDataAdapter adtr = new SqlDataAdapter("select  *from tblurun where alinanurun like '%" + textBox5.Text + "%'", conn);
            adtr.Fill(dt);
            dataGridView1.DataSource = dt;
            conn.Close();
            if (textBox5.Text == "")
            {
                dataGridView1.Show();
            }
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            maskedTextBox1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
        }
    }
}
