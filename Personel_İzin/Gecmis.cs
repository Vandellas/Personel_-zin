using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.AccessControl;
namespace Personel_İzin
{
    public partial class Gecmis : Form
    {
        int Personel_id;
        int raporid;
        string ServerName;
        public Gecmis(int Personel_id,int raporid,string ServerName)
        {
            this.ServerName = ServerName;
            this.Personel_id = Personel_id;
            this.raporid = raporid;
            InitializeComponent();
        }
        public  string BaglantiOlustur()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = ServerName;
            builder.InitialCatalog = "personel_izin";
        
            builder.IntegratedSecurity = true;
            return builder.ConnectionString;
        }
        private void Gecmis_Load(object sender, EventArgs e)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            try
            {
                myCon.Open();
            }
            catch (Exception hataDB)
            {

                MessageBox.Show("Hata DB bağlantı kurulamadı !" + hataDB.ToString());
            }
            string command = "select izin_id,rapor_türü,isim,soyad,tc_no,bas_tar,rap_sür,bitis_tar,adres from izinBilgileri where personel_id =" + Personel_id + "and rapor_id=" + raporid;
            SqlCommand cmd = new SqlCommand(command, myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                DatagridLoad(command);
            }
        }
        public void DatagridLoad(string command)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            SqlDataAdapter cmd = new SqlDataAdapter(command, myCon);
            DataSet ds = new DataSet();
            ds.Clear();
            cmd.Fill(ds, "izinBilgileri");
            dataGridView1.DataSource = ds.Tables["izinBilgileri"];
            dataGridView1.Columns[0].HeaderText = "izin_id";
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "İZİN TÜRÜ";
            dataGridView1.Columns[2].HeaderText = "İSİM";
            dataGridView1.Columns[3].HeaderText = "SOYİSİM";
            dataGridView1.Columns[4].HeaderText = "TC NUMARASİ";
            dataGridView1.Columns[5].HeaderText = "BASLAMA TARİHİ";
            dataGridView1.Columns[6].HeaderText = "RAPOR SÜRESİ";
            dataGridView1.Columns[7].HeaderText = "BİTİS TARİHİ";
            dataGridView1.Columns[8].HeaderText = "ADRES";
            cmd.Dispose();
            myCon.Close();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            izinGüncelle güncelle = new izinGüncelle(Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value), ServerName);
            güncelle.Show();
        }

        
    }
}
