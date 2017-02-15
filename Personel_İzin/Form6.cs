using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.AccessControl;

namespace Personel_İzin
{
    public partial class Form6 : Form
    {
        int Personel_id;
        string ServerName;
        public Form6(int Personel_id,string ServerName)
        {
            this.ServerName = ServerName;
            this.Personel_id=Personel_id;
            InitializeComponent();
        }
        public void DatagridLoad(string command)
        {

            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            SqlDataAdapter cmd = new SqlDataAdapter(command, myCon);
            DataSet ds = new DataSet();
            ds.Clear();
            cmd.Fill(ds, "PersonelBilgileri");
            dataGridView1.DataSource = ds.Tables["PersonelBilgileri"];
            dataGridView1.Columns[0].HeaderText = "Personel id";
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "DEVREDEN";
            dataGridView1.Columns[2].HeaderText = "BU YIL İZİNLERİ";
            dataGridView1.Columns[3].HeaderText = "TOPLAM";
            dataGridView1.Columns[5].HeaderText = "KALAN";
            dataGridView1.Columns[4].HeaderText = "KULLANILAN";
            for (int i = 0; i < 5; i++)
            {
                dataGridView1.Columns[i].Width = 174;

            }
            cmd.Dispose();
            myCon.Close();

        }
        public void izinDataLoad(string command)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlDataAdapter cmd = new SqlDataAdapter(command, myCon);
            DataSet ds = new DataSet();
            ds.Clear();
            cmd.Fill(ds, "izinBilgileri");
            dataGridView2.DataSource = ds.Tables["izinBilgileri"];
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[0].HeaderText = "izin_id";
            dataGridView2.Columns[4].HeaderText = "İZİN TÜRÜ";
            dataGridView2.Columns[2].HeaderText = "İSİM";
            dataGridView2.Columns[3].HeaderText = "SOYİSİM";
            dataGridView2.Columns[1].HeaderText = "TC NUMARASİ";
            dataGridView2.Columns[5].HeaderText = "BASLAMA TARİHİ";
            dataGridView2.Columns[6].HeaderText = "RAPOR SÜRESİ";
            dataGridView2.Columns[7].HeaderText = "BİTİS TARİHİ";
            dataGridView2.Columns[8].HeaderText = "ADRES";
            dataGridView2.Columns[9].Visible = false;
            dataGridView2.Columns[10].Visible = false;
            dataGridView2.Columns[11].Visible = false;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                if (dataGridView2.Rows[i].Cells[9].Value.ToString() == "1")
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.Orange;
            }
            myCon.Close();
        }
        private void Form6_Load(object sender, EventArgs e)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            tabControl1.SelectedIndex = 6;   
            timer1.Start();
            izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id,rapor_id,personel_id from izinBilgileri where Sil_id=0 and rapor_id!=4 and personel_id=" + Personel_id + " order by rapor_türü");
            string command = "select Personel_id,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi from PersonelBilgileri where Sil_id=0  and Personel_id=" + Personel_id;
            DatagridLoad(command);
            label82.Text = label82.Text.ToUpper();
            comboBox1.SelectedIndex = 59;
            comboBox3.SelectedIndex = 59;
            comboBox4.SelectedIndex = 59;
            comboBox5.SelectedIndex = 59;
            comboBox6.SelectedIndex = 59;
            myCon.Open();
            SqlCommand cmd = new SqlCommand("select * from dbo.PersonelBilgileri where Personel_id="+Personel_id, myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
              label77.Text=label63.Text=label49.Text= label35.Text=label16.Text= label5.Text = dr["Personel_Name"].ToString();
              label76.Text=label62.Text=label48.Text=label34.Text= label15.Text= label6.Text = dr["Personel_Surname"].ToString();
              label75.Text=label61.Text=label47.Text=label33.Text=label19.Text=label7.Text = dr["Personel_Tc"].ToString();
              label74.Text=label60.Text=label46.Text=label32.Text=label8.Text = dr["Personel_TelNumber"].ToString();
              label69.Text=label55.Text=label41.Text= label27.Text=label12.Text = dr["Personel_BasTarihi"].ToString();
              dateTimePicker11.Value=dateTimePicker9.Value=dateTimePicker7.Value=dateTimePicker5.Value=dateTimePicker3.Value = DateTime.Now.Date;
              string minunte = "";
              if (Convert.ToInt16(DateTime.Now.Minute) < 10)
                  minunte = "0" + DateTime.Now.Minute;
              else
                  minunte = DateTime.Now.Minute.ToString();
              textBox2.Text = DateTime.Now.Hour.ToString();
              textBox7.Text = minunte;
               dateTimePicker1.Value = DateTime.Now.Date;
            }
            label82.Text = label77.Text + " " + label76.Text + " Tüm İzinleri";
        }
        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {
            if (textBox3.Text != "")
            {
                dateTimePicker2.Value = dateTimePicker1.Value.AddDays(Convert.ToInt16(textBox3.Text));
            }
            else
                dateTimePicker2.Value = dateTimePicker1.Value;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            string hour = "";
            string minute = "";
            string sonucminute="";
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

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            if (textBox8.Text != "")
                dateTimePicker4.Value = dateTimePicker1.Value.AddDays(Convert.ToInt16(textBox8.Text));
            else
                dateTimePicker4.Value = dateTimePicker1.Value;
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            if (textBox10.Text != "")
            dateTimePicker6.Value = dateTimePicker7.Value.AddDays(Convert.ToInt16(textBox10.Text));
            else
                dateTimePicker6.Value = dateTimePicker7.Value;
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            if (textBox12.Text != "")
            dateTimePicker8.Value = dateTimePicker9.Value.AddDays(Convert.ToInt16(textBox12.Text));
            else
                dateTimePicker8.Value = dateTimePicker9.Value;
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            if(textBox14.Text!="")
            dateTimePicker10.Value = dateTimePicker11.Value.AddDays(Convert.ToInt16(textBox14.Text));
            else
                dateTimePicker10.Value = dateTimePicker11.Value;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Yazdır yaz = new Yazdır(label5.Text + " " + label6.Text, label12.Text, "RAPOR", dateTimePicker1.Value.Year.ToString(), string.Format("{0:dd/MM/yyyy}", dateTimePicker1.Value), string.Format("{0:dd/MM/yyyy}", dateTimePicker2.Value), textBox3.Text, comboBox1.Text, label8.Text, 1, ServerName );
           yaz.Show();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            Yazdır yaz = new Yazdır(label49.Text + " " + label48.Text, label41.Text, "İDARi İZİN", dateTimePicker5.Value.Year.ToString(), string.Format("{0:dd/MM/yyyy}", dateTimePicker5.Value), string.Format("{0:dd/MM/yyyy}", dateTimePicker4.Value), textBox8.Text, comboBox3.Text, label46.Text, 5, ServerName );
            yaz.Show();
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Yazdır yaz = new Yazdır(label49.Text + " " + label48.Text, label41.Text, "YILLIK İZİN", dateTimePicker7.Value.Year.ToString(), string.Format("{0:dd/MM/yyyy}", dateTimePicker7.Value), string.Format("{0:dd/MM/yyyy}", dateTimePicker6.Value), textBox10.Text, comboBox4.Text, label46.Text, 3, ServerName );
            yaz.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Yazdır yaz = new Yazdır(label49.Text + " " + label48.Text, label41.Text, "MAZERET İZNİ", dateTimePicker9.Value.Year.ToString(), string.Format("{0:dd/MM/yyyy}", dateTimePicker9.Value), string.Format("{0:dd/MM/yyyy}", dateTimePicker8.Value), textBox12.Text, comboBox5.Text, label46.Text, 4, ServerName );
            yaz.Show();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Yazdır yaz = new Yazdır(label49.Text + " " + label48.Text, label41.Text, "SEVK", dateTimePicker11.Value.Year.ToString(), string.Format("{0:dd/MM/yyyy}", dateTimePicker11.Value), string.Format("{0:dd/MM/yyyy}", dateTimePicker10.Value), textBox14.Text, comboBox6.Text, label46.Text, 2, ServerName );
            yaz.Show();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            pictureBox8.Enabled = true;
            SqlBaglanti sql = new SqlBaglanti(ServerName );
            sql.İzinEkle(1,Personel_id,"RAPOR", label5.Text, label6.Text, label7.Text, label8.Text, label12.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker1.Value), textBox3.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker2.Value), comboBox1.Text);
            MessageBox.Show("Kaydedildi");
            Form2 frm1 = (Form2)Application.OpenForms["Form2"];
            Form6 frm6 = (Form6)Application.OpenForms[this.Name];
            frm6.izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id,rapor_id,personel_id from izinBilgileri where Sil_id=0 and rapor_id!=4 and personel_id=" + Personel_id + " order by rapor_türü");
            string command = "select Personel_id,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi from PersonelBilgileri where Sil_id=0  and Personel_id=" + Personel_id;
            frm6.DatagridLoad(command);
            frm1.izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by  rapor_türü");
            frm1.izinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by  rapor_türü");
            frm1.AddCbxAyYıl();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            pictureBox11.Enabled = true;
            SqlBaglanti sql = new SqlBaglanti(ServerName);
            sql.İzinEkle(5,Personel_id,"İDARi İZİN",label5.Text, label6.Text, label7.Text, label8.Text, label12.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker5.Value), textBox8.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker4.Value), comboBox3.Text);
            MessageBox.Show("Kaydedildi");
            Form2 frm1 = (Form2)Application.OpenForms["Form2"];
            frm1.izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by  rapor_türü");
            frm1.izinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by  rapor_türü");
            Form6 frm6 = (Form6)Application.OpenForms[this.Name];
            frm6.izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id,rapor_id,personel_id from izinBilgileri where Sil_id=0 and rapor_id!=4 and personel_id=" + Personel_id + " order by rapor_türü");
            string command = "select Personel_id,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi from PersonelBilgileri where Sil_id=0  and Personel_id=" + Personel_id;
            frm6.DatagridLoad(command);
            frm1.AddCbxAyYıl();
        }
        public  string BaglantiOlustur()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = ServerName;
            builder.InitialCatalog = "personel_izin";
            builder.IntegratedSecurity = true;
            return builder.ConnectionString;
        }
       
        private void button8_Click(object sender, EventArgs e)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            pictureBox13.Enabled = true;
            myCon.Close();
            myCon.Open();
            int KullanılanGün=Convert.ToInt16(textBox10.Text);
            int ToplamGün=0;
            SqlCommand cmd = new SqlCommand("select * from dbo.PersonelBilgileri where Personel_id="+Personel_id, myCon);
            SqlDataReader dr = cmd.ExecuteReader();;
            while (dr.Read())
            {
                ToplamGün = Convert.ToInt16(dr["Yıllıkizin_Süresi"]) - KullanılanGün;
                KullanılanGün += Convert.ToInt16(dr["Kullanılan_Süre"]);
            }
            myCon.Close();


            SqlBaglanti sql = new SqlBaglanti(ServerName );
                sql.İzinEkle(3, Personel_id, "YILLIK İZİN", label5.Text, label6.Text, label7.Text, label8.Text, label12.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker7.Value), textBox10.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker6.Value), comboBox4.Text);
                MessageBox.Show("Kaydedildi");
                Form2 frm1 = (Form2)Application.OpenForms["Form2"];
                frm1.izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by  rapor_türü");
                frm1.izinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by  rapor_türü");
                frm1.AddCbxAyYıl();
                myCon.Open();
                SqlCommand cmd2 = new SqlCommand("  Update PersonelBilgileri set Yıllıkizin_Süresi=" + ToplamGün + ",Kullanılan_Süre=" + KullanılanGün + " where personel_id=" + Personel_id, myCon);
                cmd2.ExecuteNonQuery();
                myCon.Close();
                string command = "select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi,Sil_id from PersonelBilgileri where Sil_id=0";
                frm1.DatagridLoad(command);
                Form6 frm6 = (Form6)Application.OpenForms[this.Name];
                frm6.izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id,rapor_id,personel_id from izinBilgileri where Sil_id=0 and rapor_id!=4 and personel_id=" + Personel_id + " order by rapor_türü");
                string command2 = "select Personel_id,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi from PersonelBilgileri where Sil_id=0  and Personel_id=" + Personel_id;
                frm6.DatagridLoad(command2);
            
            if (ToplamGün < 0)
                MessageBox.Show("Kullanabilceginiz izni aştınız");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            pictureBox15.Enabled = true;
            SqlBaglanti sql = new SqlBaglanti(ServerName );
            sql.İzinEkle(0, Personel_id,"MAZERET İZNİ",label5.Text, label6.Text, label7.Text, label8.Text, label12.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker9.Value), textBox12.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker8.Value), comboBox5.Text);
            MessageBox.Show("Kaydedildi");
            Form2 frm1 = (Form2)Application.OpenForms["Form2"];
            frm1.izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by  rapor_türü");
            frm1.izinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by  rapor_türü");
            Form6 frm6 = (Form6)Application.OpenForms[this.Name];
            frm6.izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id,rapor_id,personel_id from izinBilgileri where Sil_id=0 and rapor_id!=4 and personel_id=" + Personel_id + " order by rapor_türü");
            string command = "select Personel_id,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi from PersonelBilgileri where Sil_id=0  and Personel_id=" + Personel_id;
            frm6.DatagridLoad(command);
            frm1.AddCbxAyYıl();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            pictureBox17.Enabled = true;
            SqlBaglanti sql = new SqlBaglanti(ServerName );
            sql.İzinEkle(2, Personel_id,"SEVK", label5.Text, label6.Text, label7.Text, label8.Text, label12.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker11.Value), textBox14.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker10.Value), comboBox6.Text);
            MessageBox.Show("Kaydedildi");
            Form2 frm1 = (Form2)Application.OpenForms["Form2"];
            frm1.izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by  rapor_türü");
            frm1.izinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by  rapor_türü");
            Form6 frm6 = (Form6)Application.OpenForms[this.Name];
            frm6.izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id,rapor_id,personel_id from izinBilgileri where Sil_id=0 and rapor_id!=4 and personel_id=" + Personel_id + " order by rapor_türü");
            string command = "select Personel_id,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi from PersonelBilgileri where Sil_id=0  and Personel_id=" + Personel_id;
            frm6.DatagridLoad(command);
            frm1.AddCbxAyYıl();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Gecmis gec = new Gecmis(Personel_id, 1, ServerName);
            gec.Text = "RAPOR GECMİSİ";
            gec.Show();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            Gecmis gec = new Gecmis(Personel_id, 5, ServerName );
            gec.Text = "İDARİ İZİN GECMİSİ";
            gec.Show();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            Gecmis gec = new Gecmis(Personel_id, 3, ServerName );
            gec.Text = "YILLIK İZİN GECMİSİ";
            gec.Show();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            Gecmis gec = new Gecmis(Personel_id, 0, ServerName );
            gec.Text = "MAZERET İZNİ GECMİSİ";
            gec.Show();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            Gecmis gec = new Gecmis(Personel_id, 2, ServerName );
            gec.Text = "SEVK GECMİS";
            gec.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            GünlükYazdır yaz = new GünlükYazdır(string.Format("{0:dd/MM/yyyy}", dateTimePicker3.Value), textBox2.Text + ":" + textBox7.Text, textBox5.Text, label16.Text + " " + label15.Text, textBox1.Text, ServerName );
            yaz.Show();
     
        }

        private void button4_Click(object sender, EventArgs e)
        {
            pictureBox9.Enabled = true;
            SqlBaglanti sql = new SqlBaglanti(ServerName );
            sql.GünlükizinEkle(4, Personel_id, "GÜNÜBİRLİK İZİN", label16.Text, label15.Text, label19.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker3.Value), textBox2.Text + ":" + textBox7.Text, textBox5.Text, textBox4.Text + ":" + textBox6.Text, textBox1.Text);
            Form2 frm1 = (Form2)Application.OpenForms["Form2"];
            frm1.GünlükizinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id  from izinBilgileri where  Sil_id=0 and rapor_id=4");
            frm1.Günübirlikizindataload("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id  from izinBilgileri where  Sil_id=0 and rapor_id=4");
            MessageBox.Show("Basariyla Eklendi");
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 6)
            {
                this.WindowState = FormWindowState.Maximized;
                dataGridView1.Width = this.Width;
                dataGridView2.Width = this.Width;
                dataGridView2.Height = this.Height - 300;
                tabControl1.Size = this.Size;
                this.StartPosition = FormStartPosition.CenterScreen;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
                this.StartPosition = FormStartPosition.CenterScreen;
            }
        }
        int i = 0;
        int x = -200;
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Interval = 1;
            int y = label82.Location.Y;
            i=i+2;
            i = i % 1400;
            label82.Location = new Point(x + i, y);
           
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

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            string[] words = textBox5.Text.Split(':');
            TimeSpan saat1 = new TimeSpan(Convert.ToInt16(words[0]), Convert.ToInt16(words[1]), 0);
            TimeSpan saat2 = new TimeSpan(17, 0, 0);
            TimeSpan sonuc = new TimeSpan();
           sonuc = saat2 - saat1;
           TimeSpan karsılastirma = new TimeSpan(0, 0, 0);
           if (sonuc < karsılastirma)
               MessageBox.Show("Mesaiti Saatini Geçntiz...", "Dikkat");

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            izinGüncelle gün = new izinGüncelle(Convert.ToInt16(dataGridView2.CurrentRow.Cells[0].Value), ServerName );
            gün.Show();
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


    }
}
