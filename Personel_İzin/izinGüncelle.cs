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

namespace Personel_İzin
{
    public partial class izinGüncelle : Form
    {
        int izinid;
        string ServerName;
        public izinGüncelle(int izinid,string ServerName)
        {
            this.ServerName = ServerName;
            this.izinid = izinid;
            InitializeComponent();
        }
        public string BaglantiOlustur()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = ServerName;
            builder.InitialCatalog = "personel_izin";
            builder.IntegratedSecurity = true;
            return builder.ConnectionString;
        }
       
        int rapor_id = 12, personel_id = -2; string rapor_süresi = "";
        string raportürü = "";
        public void Readİzin()
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlCommand cmd = new SqlCommand("select * from izinBilgileri where izin_id="+izinid, myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label35.Text = dr["isim"].ToString();
                label34.Text = dr["soyad"].ToString();
                label33.Text = dr["tc_no"].ToString();
                label27.Text = dr["isegiris_tar"].ToString();
                dateTimePicker5.Value = Convert.ToDateTime(dr["bas_tar"]);
                rapor_süresi=textBox8.Text = dr["rap_sür"].ToString();
                dateTimePicker4.Value = Convert.ToDateTime(dr["bitis_tar"]);
                comboBox3.Text = dr["adres"].ToString();
                label32.Text = dr["tel_no"].ToString();
                raportürü=this.Text = dr["rapor_türü"].ToString();
                this.Text +=" Güncelle";
                rapor_id = Convert.ToInt16(dr["rapor_id"]);
                personel_id = Convert.ToInt16(dr["personel_id"]);


            }
            myCon.Close();
           
        }
        private void Güncelle_Load(object sender, EventArgs e)
        {
            Readİzin();
        }

        private void pictureBox12_Click(object sender, EventArgs e)
        {
            if (rapor_id != 3)
            {
                Updateizin();
                
            }
            else
            {
                Updateizin();
                ReadPersonel();
                UpdatePersonel();
                Form2 frm1 = (Form2)Application.OpenForms["Form2"];
                string command = "select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi,Sil_id from PersonelBilgileri where Sil_id=0";
                frm1.DatagridLoad(command);
            }
        }
        int kullanılan = 0;
        int kalan = 0;
        public void ReadPersonel()
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlCommand cmd = new SqlCommand("select * from dbo.PersonelBilgileri where Personel_id=" + personel_id, myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                kullanılan = Convert.ToInt16(dr["Kullanılan_Süre"]);
                kalan = Convert.ToInt16(dr["Yıllıkizin_Süresi"]);
            }
            myCon.Close();
        }
        public void Updateizin()
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            string vsql = string.Format("UPDATE  izinBilgileri set bas_tar = @bastar, rap_sür = @sür,bitis_tar=@bittar,adres=@adres where izin_id=" + izinid);
            myCon.Open();
            SqlCommand azert = new SqlCommand(vsql, myCon);
            azert.Parameters.Add("@bastar", string.Format("{0:dd/MM/yyyy}", dateTimePicker5.Value));
            azert.Parameters.Add("@sür", textBox8.Text);
            azert.Parameters.Add("@bittar", string.Format("{0:dd/MM/yyyy}", dateTimePicker4.Value));
            azert.Parameters.Add("@adres", comboBox3.Text);
            azert.ExecuteNonQuery();
            MessageBox.Show("Basariyla Güncelendi");
            myCon.Close();
            Form2 frm1 = (Form2)Application.OpenForms["Form2"];
            frm1.izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by rapor_türü ");
            frm1.izinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where  Sil_id like 0 and rapor_id!=4 order by  rapor_türü");
            frm1.AddCbxAyYıl();
        }
        public void UpdatePersonel()
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            kullanılan = Convert.ToInt16(textBox8.Text) - Convert.ToInt16(rapor_süresi)+kullanılan;
            kalan = Convert.ToInt16(rapor_süresi) - Convert.ToInt16(textBox8.Text) + kalan;
            string vsql = string.Format("UPDATE  PersonelBilgileri set Kullanılan_Süre = @kulsüre, Yıllıkizin_Süresi = @yılsür where Personel_id="+personel_id);
            myCon.Open();
            SqlCommand azert = new SqlCommand(vsql, myCon);
            azert.Parameters.Add("@kulsüre", kullanılan);
            azert.Parameters.Add("@yılsür",kalan);
            azert.ExecuteNonQuery();
        }
        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            if (textBox8.Text != "")
                dateTimePicker4.Value = dateTimePicker5.Value.AddDays(Convert.ToInt16(textBox8.Text));
            else
                dateTimePicker4.Value = dateTimePicker5.Value;

        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {
            Yazdır yaz = new Yazdır(label35.Text + " " + label34.Text, label27.Text, raportürü, dateTimePicker5.Value.Year.ToString(), string.Format("{0:dd/MM/yyyy}", dateTimePicker5.Value), string.Format("{0:dd/MM/yyyy}", dateTimePicker4.Value), textBox8.Text, comboBox3.Text, label32.Text, Convert.ToInt16(rapor_id), ServerName );
            yaz.Show();
        }

        private void label32_Click(object sender, EventArgs e)
        {

        }
    }
}
