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
    public partial class GünüBrilikGüncelleme : Form
    {
        int izinid;
        string ServerName;
        public GünüBrilikGüncelleme(int izinid,string ServerName)
        {
            this.ServerName = ServerName;
            this.izinid = izinid;
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
        
        private void GünüBrilikGüncelleme_Load(object sender, EventArgs e)
        {
            Readİzin();
        }
        public void Readİzin()
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlCommand cmd = new SqlCommand("select * from izinBilgileri where izin_id=" + izinid, myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label16.Text = dr["isim"].ToString();
                label15.Text = dr["soyad"].ToString();
                label19.Text = dr["tc_no"].ToString();
                dateTimePicker3.Value = Convert.ToDateTime(dr["bas_tar"]);
                string[] word = dr["Baslama_Saati"].ToString().Split(':');
                textBox2.Text = word[0];
                textBox7.Text = word[1];
                string[] word2 = dr["Kac_Saat"].ToString().Split(':');
                textBox4.Text = word2[0];
                textBox6.Text = word2[1];
                textBox5.Text = dr["Bitis_Saati"].ToString();
                textBox1.Text = dr["Mazeret"].ToString();


            }
            myCon.Close();

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string hour = "";
            string minute = "";
            string text2 = "";
            if (textBox2.Text == "")
                text2 = "00";
            else
                text2 = textBox2.Text;
            string sonucminute = "";
            if (textBox6.Text == "")
                minute = "00";
            else
                minute = textBox6.Text;
            if (textBox4.Text == "")
                hour = "00";
            else
                hour = textBox4.Text;

            TimeSpan saat1 = new TimeSpan(Convert.ToInt16(text2), Convert.ToInt16(textBox7.Text), 0);
            TimeSpan saat2 = new TimeSpan(Convert.ToInt16(hour), Convert.ToInt16(minute), 0);
            TimeSpan sonuc = new TimeSpan();
            sonuc = saat1 + saat2;
            if (sonuc.Minutes < 10)
                sonucminute = "0" + sonuc.Minutes;
            else
                sonucminute = sonuc.Minutes.ToString();

            textBox5.Text = sonuc.Hours + ":" + sonucminute;
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            string hour = "";
            string sonucminute = "";
            if (textBox4.Text == "")
                hour = "00";
            else
                hour = textBox4.Text;
            string minunte;
            if (textBox6.Text == "")
                minunte = "00";
            else
                minunte = textBox6.Text;
            string text7;
            if (textBox7.Text == "")
                text7 = "00";
            else
                text7 = textBox7.Text;
            TimeSpan saat1 = new TimeSpan(Convert.ToInt16(textBox2.Text), Convert.ToInt16(text7), 0);
            TimeSpan saat2 = new TimeSpan(Convert.ToInt16(hour), Convert.ToInt16(minunte), 0);
            TimeSpan sonuc = new TimeSpan();
            sonuc = saat1 + saat2;
            if (sonuc.Minutes < 10)
                sonucminute = "0" + sonuc.Minutes;
            else
                sonucminute = sonuc.Minutes.ToString();
            textBox5.Text = sonuc.Hours + ":" + sonucminute;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            string hour = "";
            string minute = "";
            string sonucminute = "";
            if (textBox6.Text == "")
                minute = "00";
            else
                minute = textBox6.Text;
            if (textBox4.Text == "")
                hour = "00";
            else
                hour = textBox4.Text;
            TimeSpan saat1 = new TimeSpan(Convert.ToInt16(textBox2.Text), Convert.ToInt16(textBox7.Text), 0);
            TimeSpan saat2 = new TimeSpan(Convert.ToInt16(hour), Convert.ToInt16(minute), 0);
            TimeSpan sonuc = new TimeSpan();
            sonuc = saat1 + saat2;
            if (sonuc.Minutes < 10)
                sonucminute = "0" + sonuc.Minutes;
            else
                sonucminute = sonuc.Minutes.ToString();

            textBox5.Text = sonuc.Hours + ":" + sonucminute;
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            string hour = "";
            string sonucminute = "";
            if (textBox4.Text == "")
                hour = "00";
            else
                hour = textBox4.Text;
            string minunte;
            if (textBox6.Text == "")
                minunte = "00";
            else
                minunte = textBox6.Text;
            TimeSpan saat1 = new TimeSpan(Convert.ToInt16(textBox2.Text), Convert.ToInt16(textBox7.Text), 0);
            TimeSpan saat2 = new TimeSpan(Convert.ToInt16(hour), Convert.ToInt16(minunte), 0);
            TimeSpan sonuc = new TimeSpan();
            sonuc = saat1 + saat2;
            if (sonuc.Minutes < 10)
                sonucminute = "0" + sonuc.Minutes;
            else
                sonucminute = sonuc.Minutes.ToString();

            textBox5.Text = sonuc.Hours + ":" + sonucminute;
        }

        private void pictureBox12_Click(object sender, EventArgs e)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            string vsql = string.Format("UPDATE  izinBilgileri  set bas_tar = @bastar, Baslama_Saati = @basaat,Kac_Saat=@kacsaat,Bitis_Saati=@bitsaat,Mazeret=@mazeret where izin_id="+izinid);
            myCon.Open();
            SqlCommand azert = new SqlCommand(vsql, myCon);
            azert.Parameters.Add("@bastar", string.Format("{0:dd/MM/yyyy}",dateTimePicker3.Value));
            azert.Parameters.Add("@basaat", textBox2.Text+":"+textBox7.Text);
            azert.Parameters.Add("@kacsaat", textBox4.Text+":"+textBox6.Text);
            azert.Parameters.Add("@bitsaat", textBox5.Text);
            azert.Parameters.Add("@mazeret", textBox1.Text);
            azert.ExecuteNonQuery();
            MessageBox.Show("Basariyla Güncelendi");
            myCon.Close();
            Form2 frm1 = (Form2)Application.OpenForms["Form2"];
            frm1.GünlükizinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id from izinBilgileri  where Sil_id=0 and rapor_id=4");
            frm1.AddCbxAyYıl();
            frm1.Günübirlikizindataload("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id  from izinBilgileri where Sil_id=0 and rapor_id=4");
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            GünlükYazdır yaz = new GünlükYazdır(dateTimePicker3.Value.Year.ToString(), textBox2.Text + ":" + textBox7.Text, textBox5.Text, label16.Text + " " + label15.Text, textBox1.Text, ServerName );
            yaz.Show();
        }
    }
}
