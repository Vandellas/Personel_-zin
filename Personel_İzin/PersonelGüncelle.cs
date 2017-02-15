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
using System.Diagnostics;  
using Microsoft.Win32;
namespace Personel_İzin
{
    public partial class PersonelGüncelle : Form
    {
        int Personel_id;
        string kullanıcıismi;
        int Update_id;
        string ServerName;
        public PersonelGüncelle(int Personel_id,string kullanıcıismi,int Update_id,string ServerName)
        {
            this.ServerName = ServerName;
            this.Update_id = Update_id;
            this.Personel_id = Personel_id;
            this.kullanıcıismi = kullanıcıismi;
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
       
        SqlCommand sorgu = new SqlCommand();
        private void btngüncelle_Click(object sender, EventArgs e)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlCommand sorgu = new SqlCommand(" Update PersonelBilgileri set ÖzürlülükYüzdesi='"+txtYüzde.Text+"',ÖzürlülükDurumu='" + cbxÖzürlülükDurumu.Text + "',ÇocukSayisi='" + txtSayisi.Text + "',EvlilikDurumu='" + cbxEvlilikDurumu.Text + "' ,Personel_Tc='" + txttcno.Text + "',Personel_Name='" + txtisim.Text + "',EsDurumu='" + cbxEsdurumu.Text + "',Personel_Surname=" + "'" + txtsoyisim.Text + "'" + ",Personel_SGK='" + txtsgk.Text + "',Personel_FatherName='" + txtbabaisim.Text + "',Personel_Birthday='" + string.Format("{0:dd/MM/yyyy}", DateDogumtar.Value) + "',Personel_Hometown='" + comboBox6.Text + "',Personel_BasTarihi='" + string.Format("{0:dd/MM/yyyy}", datebaslamatar.Value) + "',YıllıkizinBaslama_Tar='" + string.Format("{0:dd/MM/yyyy}", datebaslamatar.Value) + "',Personel_TelNumber='" + txttelno.Text + "',Personel_Adres='" + txtadres.Text + "',Personel_Göreviyeri='" + txtgörevyeri.Text + "',Update_Date='" + DateTime.Now.Date + "',Update_User='" + kullanıcıismi + "',Update_id=" + Update_id + ",HesapNo='" + txthesapno.Text + "',Banka='" + txtbanka.Text + "',Mail='" + txtmail.Text + "' where Personel_id=" + Personel_id, myCon);
            sorgu.ExecuteNonQuery();
            myCon.Close();
            myCon.Open();
            SqlCommand sorgu2 = new SqlCommand("Update PersonelBilgileri set Update_User=Personel_Name where Update_id="+Personel_id, myCon);
            sorgu2.ExecuteNonQuery();
            myCon.Close();
            myCon.Open();
            SqlCommand sorgu3 = new SqlCommand("Update PersonelBilgileri set Recort_User=Personel_Name where Recort_id="+Personel_id, myCon);
            sorgu3.ExecuteNonQuery();
            myCon.Close();
            myCon.Open();
            SqlCommand sorgu4 = new SqlCommand("Update KullaniciBilgileri set PersonelName =  PersonelBilgileri.Personel_Name from PersonelBilgileri where PersonelBilgileri.Personel_id=KullaniciBilgileri.Personelid", myCon);
            sorgu4.ExecuteNonQuery();
            myCon.Close();
            myCon.Open();
            SqlCommand cmd = new SqlCommand("Update izinBilgileri set isegiris_tar='" + string.Format("{0:dd/MM/yyyy}", datebaslamatar.Value) +"',isim='" + txtisim.Text + "',soyad='" + txtsoyisim.Text + "',tc_no='" + txttcno.Text + "'  where personel_id=" + Personel_id, myCon);
            cmd.ExecuteNonQuery();
            myCon.Close();
            MessageBox.Show("Kayıt güncelendi");
            Form2 frm1 = (Form2)Application.OpenForms["Form2"];
            string command = "select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi, Sil_id from PersonelBilgileri where Sil_id=0";
            frm1.DatagridLoad(command);
            frm1.izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0  order by  rapor_türü");
            frm1.izinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0  order by  rapor_türü");
            frm1.PersonelRaporlama("select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,Personel_Göreviyeri from PersonelBilgileri where Sil_id=0 ", 0);
            this.Close();
           
        }

        private void PersonelGüncelle_Load(object sender, EventArgs e)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            string[][] str = new string[16][];
            myCon.Open();
            SqlCommand cmd = new SqlCommand("select * from dbo.PersonelBilgileri where Personel_id="+Personel_id, myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {

                txtagitc.Text= txttcno.Text = dr["Personel_Tc"].ToString();
                txtagiad.Text= txtisim.Text = dr["Personel_Name"].ToString();
                txtagisoyad.Text=txtsoyisim.Text = dr["Personel_Surname"].ToString();
                txtsgk.Text = dr["Personel_SGK"].ToString();
                txtbabaisim.Text = dr["Personel_FatherName"].ToString();
                DateDogumtar.Text = dr["Personel_Birthday"].ToString();
                comboBox6.Text = dr["Personel_Hometown"].ToString();
                datebaslamatar.Text = dr["Personel_BasTarihi"].ToString();
                txttelno.Text = dr["Personel_TelNumber"].ToString();
                txtadres.Text= dr["Personel_Adres"].ToString();
                txtgörevyeri.Text= dr["Personel_Göreviyeri"].ToString();
                txtmail.Text = dr["Mail"].ToString();
                txthesapno.Text = dr["HesapNo"].ToString();
                txtbanka.Text = dr["Banka"].ToString();
                cbxagimedenidurumu.Text=cbxEvlilikDurumu.Text = dr["EvlilikDurumu"].ToString();
                txtagicocuksayisi.Text=txtSayisi.Text = dr["ÇocukSayisi"].ToString();
                cbxagiözürlülük.Text=cbxÖzürlülükDurumu.Text = dr["ÖzürlülükDurumu"].ToString();
                txtagiözürlülükyüz.Text=txtYüzde.Text = dr["ÖzürlülükYüzdesi"].ToString();
                cbxagiesdurumu.Text=cbxEsdurumu.Text = dr["EsDurumu"].ToString();
            }
            myCon.Close();
        }

        private void cbxEvlilikDurumu_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxEvlilikDurumu.SelectedIndex == 0 || cbxEvlilikDurumu.SelectedIndex == 2)
            {
                txtSayisi.Text = "";
                txtSayisi.Enabled = true;
                
            }
            else
            {
                txtSayisi.Text = "X";
                txtSayisi.Enabled = false;
            }
            if (cbxEvlilikDurumu.SelectedIndex == 1 || cbxEvlilikDurumu.SelectedIndex == 2)
            {
                cbxEsdurumu.Enabled = false;
                cbxEsdurumu.Text = "X";
            }
            else if (cbxEvlilikDurumu.SelectedIndex == 0)
            {
                cbxEsdurumu.Text = "Çalışmıyor";
                cbxEsdurumu.Enabled = true;
            }
        }

        private void cbxÖzürlülükDurumu_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxÖzürlülükDurumu.Text == "YOK")
            {
                txtYüzde.Enabled = false;
                txtYüzde.Text = "X";
            }
            else if (cbxÖzürlülükDurumu.Text == "VAR")
            {
                txtYüzde.Text = "";
                txtYüzde.Enabled = true;
            }

        }
    }
}
