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
using System.IO;

namespace Personel_İzin
{
    public partial class Form2 : Form
    {
        String[] Aylar = { "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık" };
        string[] dataheadertext = new string[16];
        string izin_Name = "isim";
        string Rapor_Türü = "rapor_türü";
        string Soyad = "soyad";
        string Tc_No = "tc_no";
        string Sil_id = "Sil_id";
        string ServerName;
        public Form2(string ServerName)
        {
            this.ServerName = ServerName;
            InitializeComponent();
        }
        public void AddCbxAyYıl()
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            CbxYıl.Items.Clear();
            CbxYıl.Items.Add(23);
            myCon.Open();
            SqlCommand cmd = new SqlCommand("select * from dbo.izinBilgileri where Sil_id=0", myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {

                string[] str = dr["bas_tar"].ToString().Split('.');
                int bayrak = 0;
                for (int i = 0; i < CbxYıl.Items.Count; i++)
                {
                    if (CbxYıl.Items[i].ToString().CompareTo(str[2]) == 0)
                    {
                        bayrak = 1;
                    }
                }
                if (bayrak == 0)
                    CbxYıl.Items.Add(str[2]);
            }
            CbxYıl.Items.RemoveAt(0);

            myCon.Close();
        }
        string Personel_Tc = "Personel_Tc";
        string Personel_isim = "Personel_Name";
        string Personel_soyad = "Personel_Surname";
        string PersonelSil_id = "Sil_id";
        private void Form2_Load(object sender, EventArgs e)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlCommand cmd = new SqlCommand("select PersonelOtomasyonName,PersonelOtomasyonSurname from Ayarlar", myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                txtotomüdürname.Text=dr["PersonelOtomasyonName"].ToString();
                txtotomüdürsurname.Text = dr["PersonelOtomasyonSurname"].ToString();
            }
            myCon.Close();
            cbxRaporlama.SelectedIndex = 0;
            comboBox3.SelectedIndex = 1;
            comboBox2.SelectedIndex = 1;
            tabControl1.SelectedIndex = 1;
            comboBox6.SelectedIndex = 59;
            dataGridView1.Width = this.Size.Width - 50;
            datagridRaporlama.Height = this.Size.Height - 300;
            datagridRaporlama.Width = this.Size.Width - 50;
            dataGridView1.Height = this.Size.Height - 300;
            dataGridView3.Width = this.Size.Width - 60;
            dataGridView3.Height = this.Size.Height - 300;
            dataGridView2.Width = this.Size.Width - 60;
            dataGridView4.Width = this.Size.Width - 50;
            dataGridView4.Height = this.Size.Height - 400;
            dataGridView5.Width = this.Size.Width - 50;
            tabControl1.Width = this.Width;
            tabControl1.Size = this.Size;
            string command = "select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi,Sil_id from PersonelBilgileri where Sil_id=0";
            DatagridLoad(command);
            İzin_Türü.SelectedIndex = 0;
            izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by rapor_türü ");
            try
            {
                dataGridView4.Visible = true;
                izinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where bas_tar like '%" + Tarih + "%' and Sil_id like 0 and rapor_id!=4 order by  rapor_türü");
                dataGridView4.Rows[0].Cells[0].Value.ToString();
            }
            catch
            {
                dataGridView4.Visible = false;
            }
            try
            {
                dataGridView4.Location = new Point(8, 235);
                dataGridView5.Visible = true;
                GünlükizinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id  from izinBilgileri where  bas_tar like '%" + Tarih + "%' and Sil_id=0 and rapor_id=4");
                dataGridView5.Rows[0].Cells[0].Value.ToString();
            }
            catch
            {
                dataGridView4.Location = new Point(9, 85);
                dataGridView5.Visible = false;
            }
            GünlükizinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id from izinBilgileri  where Sil_id=0 and rapor_id=4");
            AddCbxAyYıl();

        }
        public  string BaglantiOlustur()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = ServerName;
            builder.InitialCatalog = "personel_izin";
            builder.IntegratedSecurity = true;
            return builder.ConnectionString;
        }
        public void izinDataLoad(string command)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlDataAdapter cmd = new SqlDataAdapter(command, myCon);
            DataSet ds = new DataSet();
            ds.Clear();
            cmd.Fill(ds, "izinBilgileri");
            dataGridView3.DataSource = ds.Tables["izinBilgileri"];
            dataGridView3.Columns[0].Visible = false;
            dataGridView3.Columns[0].HeaderText = "izin_id";
            dataGridView3.Columns[4].HeaderText = "İZİN TÜRÜ";
            dataGridView3.Columns[2].HeaderText = "İSİM";
            dataGridView3.Columns[3].HeaderText = "SOYİSİM";
            dataGridView3.Columns[1].HeaderText = "TC NUMARASİ";
            dataGridView3.Columns[5].HeaderText = "BASLAMA TARİHİ";
            dataGridView3.Columns[6].HeaderText = "RAPOR SÜRESİ";
            dataGridView3.Columns[7].HeaderText = "BİTİS TARİHİ";
            dataGridView3.Columns[8].HeaderText = "ADRES";
            dataGridView3.Columns[9].Visible = false;
            for (int i = 0; i < dataGridView3.Columns.Count; i++)
            {

                dataGridView3.Columns[i].Width = 119;
            }
            for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
            {
                if (dataGridView3.Rows[i].Cells[9].Value.ToString() == "1")
                    dataGridView3.Rows[i].DefaultCellStyle.BackColor = Color.Orange;
            }
            myCon.Close();
        }
        public void Günübirlikizindataload(string command)
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
            dataGridView2.Columns[1].HeaderText = "T.C. NUMARASİ";
            dataGridView2.Columns[5].HeaderText = "BASLAMA TARİHİ";
            dataGridView2.Columns[6].HeaderText = "BASLAMA SAATİ";
            dataGridView2.Columns[7].HeaderText = "KAÇ SAAT";
            dataGridView2.Columns[8].HeaderText = "BİTİS SAATİ";
            dataGridView2.Columns[9].HeaderText = "MAZERET";
            dataGridView2.Columns[10].Visible = false;
            dataGridView2.Columns[4].Width = 150;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                if (dataGridView2.Rows[i].Cells[10].Value.ToString() == "1")
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.Orange;
            }
            myCon.Close();
        }
        public void izinExcelDataLoad(string command)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlDataAdapter cmd = new SqlDataAdapter(command, myCon);
            DataSet ds = new DataSet();
            ds.Clear();
            cmd.Fill(ds, "izinBilgileri");
            dataGridView4.DataSource = ds.Tables["izinBilgileri"];
            dataGridView4.Columns[0].Visible = false;
            dataGridView4.DataSource = ds.Tables["izinBilgileri"];
            dataGridView4.Columns[0].Visible = false;
            dataGridView4.Columns[0].HeaderText = "izin_id";
            dataGridView4.Columns[4].HeaderText = "İZİN TÜRÜ";
            dataGridView4.Columns[2].HeaderText = "İSİM";
            dataGridView4.Columns[3].HeaderText = "SOYİSİM";
            dataGridView4.Columns[1].HeaderText = "TC NUMARASİ";
            dataGridView4.Columns[5].HeaderText = "BASLAMA TARİHİ";
            dataGridView4.Columns[6].HeaderText = "RAPOR SÜRESİ";
            dataGridView4.Columns[7].HeaderText = "BİTİS TARİHİ";
            dataGridView4.Columns[8].HeaderText = "ADRES";
            dataGridView4.Columns[9].Visible = false;
            for (int i = 0; i < dataGridView4.Columns.Count; i++)
            {
                dataGridView4.Columns[i].Width = 119;
            }
            
            myCon.Close();
        }
        public void GünlükizinExcelDataLoad(string command)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlDataAdapter cmd = new SqlDataAdapter(command, myCon);
            DataSet ds = new DataSet();
            ds.Clear();
            cmd.Fill(ds, "izinBilgileri");
            dataGridView5.DataSource = ds.Tables["izinBilgileri"];
            dataGridView5.Columns[0].Visible = false;
            dataGridView5.Columns[0].HeaderText = "izin_id";
            dataGridView5.Columns[4].HeaderText = "İZİN TÜRÜ";
            dataGridView5.Columns[2].HeaderText = "İSİM";
            dataGridView5.Columns[3].HeaderText = "SOYİSİM";
            dataGridView5.Columns[1].HeaderText = "T.C. NUMARASİ";
            dataGridView5.Columns[5].HeaderText = "BASLAMA TARİHİ";
            dataGridView5.Columns[6].HeaderText = "BASLAMA SAATİ";
            dataGridView5.Columns[7].HeaderText = "KAÇ SAAT";
            dataGridView5.Columns[8].HeaderText = "BİTİS SAATİ";
            dataGridView5.Columns[9].HeaderText = "MAZERET";
            dataGridView5.Columns[10].Visible = false;
            dataGridView5.Columns[4].Width = 150;
            myCon.Close();
        }
        public void PersonelRaporlama(string command, int headertext)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            int count = 0;
            SqlDataAdapter cmd = new SqlDataAdapter(command, myCon);
            DataSet ds = new DataSet();
            ds.Clear();
            cmd.Fill(ds, "PersonelBilgileri");
            datagridRaporlama.DataSource = ds.Tables["PersonelBilgileri"];
            if (headertext == 2)
                count = 9;
            else
                count = 5;
            datagridRaporlama.Columns[0].Visible = false;
            for (int i = 0; i < count; i++)
            {
                datagridRaporlama.Columns[i].Width = 160;
            }
            if (headertext == 0)
            {
                datagridRaporlama.Columns[4].HeaderText = "GÖREV YERİ";
            }
            else if (headertext == 1)
            {
                datagridRaporlama.Columns[4].HeaderText = "HESAP NUMARASİ";
            }
            else if (headertext == 2)
            {
                datagridRaporlama.Columns[4].HeaderText = "MEDENİ DURUMU";
                datagridRaporlama.Columns[5].HeaderText = "ÇOCUK SAYİSİ";
                datagridRaporlama.Columns[6].HeaderText = "ESİNİN DURUMU";
                datagridRaporlama.Columns[7].HeaderText = "ÖZÜRLÜLÜK DURUMU";
                datagridRaporlama.Columns[8].HeaderText = "ÖZÜRLÜLÜK YÜZDESİ";
            }
            else if (headertext == 3)
            {
                datagridRaporlama.Columns[4].HeaderText = "TELEFON NUMARASİ";
            }
            else if (headertext == 4)
            {
                datagridRaporlama.Columns[4].HeaderText = "SGK NUMARASİ";
            }
            datagridRaporlama.Columns[1].HeaderText = "TC NUMARASİ";
            datagridRaporlama.Columns[2].HeaderText = "İSİM";
            datagridRaporlama.Columns[3].HeaderText = "SOYİSİM";
            cmd.Dispose();
            myCon.Close();

        }
        public void DatagridLoad(string command)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            SqlDataAdapter cmd = new SqlDataAdapter(command, myCon);
            DataSet ds = new DataSet();
            ds.Clear();
            cmd.Fill(ds, "PersonelBilgileri");
            dataGridView1.DataSource = ds.Tables["PersonelBilgileri"];
            dataheadertext = new string[12];
            for (int i = 0; i < 9; i++)
            {
                dataheadertext[i] = dataGridView1.Columns[i].Name;
            }
            dataGridView1.Columns[0].HeaderText = "Personel id";
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "T.C. NO ";
            dataGridView1.Columns[2].HeaderText = "İSİM";
            dataGridView1.Columns[3].HeaderText = "SOYAD";
            dataGridView1.Columns[4].HeaderText = "DEVREDEN";
            dataGridView1.Columns[5].HeaderText = "BU YIL İZİNLERİ";
            dataGridView1.Columns[6].HeaderText = "TOPLAM";
            dataGridView1.Columns[8].HeaderText = "KALAN";
            dataGridView1.Columns[7].HeaderText = "KULLANILAN";
            dataGridView1.Columns[9].Visible = false;
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {

                dataGridView1.Columns[i].Width = 160;
            }
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[9].Value.ToString() == "1")
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Orange;
            }
            cmd.Dispose();
            myCon.Close();
            SqlBaglanti sql = new SqlBaglanti(ServerName);
            string[][] PersonelBilgileri = sql.PersonelOku();
            comboBox1.Items.Clear();
            comboBox1.Text = "";
            textBox13.Text = "";
            textBox11.Text = "";
            for (int i = 0; i < Convert.ToInt16(PersonelBilgileri[15][0]); i++)
            {
                comboBox1.Items.Add(PersonelBilgileri[1][i].ToString());
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {

            if (hata == 0)
            {
                if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox9.Text != "" && textBox8.Text != "" && comboBox2.Text != "" && textBox7.Text != "" && textBox5.Text != "" && textBox6.Text != "" && txthesapno.Text != "" && txtbanka.Text != "" && txtmail.Text != "")
                {
                    if (textBox1.Text.Length == 11)
                    {
                        int flag = 0;
                        SqlBaglanti sql = new SqlBaglanti(ServerName);
                        string[][] PersonelTc = sql.PersonelOku();
                        for (int i = 0; i < Convert.ToInt16(PersonelTc[15][0]); i++)
                        {
                            if (textBox1.Text == PersonelTc[0][i].ToString())
                            {
                                flag = 1;
                            }
                        }
                        if (flag == 0)
                        {
                            sql.PersonelEkle(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker1.Value), comboBox6.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker2.Value), textBox9.Text, textBox6.Text, string.Format("{0:dd/MM/yyyy}", dateTimePicker1.Value), Form1.KullaniciAdi, textBox8.Text, Form1.Personel_id, txtbanka.Text, txtmail.Text, txthesapno.Text, comboBox2.Items[comboBox2.SelectedIndex].ToString(), textBox7.Text, comboBox3.Items[comboBox3.SelectedIndex].ToString(), textBox12.Text, cbxEsdurumu.Text);
                            MessageBox.Show("kaydedildi");
                            string command = "select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi,Sil_id from PersonelBilgileri where Sil_id=0";
                            PersonelRaporlama("select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,Personel_Göreviyeri from PersonelBilgileri where Sil_id=0 ", 0);
                            DatagridLoad(command);
                            txtmail.Text = txtbanka.Text = txthesapno.Text = textBox1.Text = textBox2.Text = textBox3.Text = textBox4.Text = textBox5.Text = textBox6.Text = textBox8.Text = textBox9.Text = "";
                        }
                        else
                            MessageBox.Show("Bu T.C. Numarasi Zaten Kayitlarda Bulunuyo...!!!");
                    }
                    else
                        MessageBox.Show("Tc numarasinin kontrol ediniz...!!!");
                }
                else
                    MessageBox.Show("Gerekli Bosluklari doldurunuz....!!!");
            }
            else
            {
                MessageBox.Show("Lütfen Bilgilerinizi Kontrol Ediniz...");
            }
        }
        int selectedindex = 0;
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            SqlBaglanti sql = new SqlBaglanti(ServerName);
            if (dataGridView1.CurrentRow.Cells[0].Value.ToString() != "")
            {
                if (e.RowIndex != -1)
                {
                    if (izin.Checked == true)
                    {
                        Form6 form6 = new Form6(Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value), ServerName );
                        form6.Show();
                    }
                    else if (Silradio.Checked == true)
                    {

                        DialogResult sonuc = MessageBox.Show("Silmek istedginizden eminmisin?", "Dikkat", MessageBoxButtons.YesNo);
                        if (sonuc == DialogResult.Yes)
                        {
                            
                            sql.SilPersonel(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                            sql.SilKullanici(Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value));
                            sql.Silizin(dataGridView1.CurrentRow.Cells[0].Value.ToString(), "personel_id");
                            izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by rapor_türü");
                            try
                            {
                                dataGridView4.Visible = true;
                                izinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where bas_tar like '%" + Tarih + "%' and Sil_id like 0 and rapor_id!=4 order by  rapor_türü");
                                dataGridView4.Rows[0].Cells[0].Value.ToString();
                            }
                            catch
                            {
                                dataGridView4.Visible = false;
                            }
                            try
                            {
                                dataGridView4.Location = new Point(8, 235);
                                dataGridView5.Visible = true;
                                GünlükizinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id  from izinBilgileri where  bas_tar like '%" + Tarih + "%' and Sil_id=0 and rapor_id=4");
                                dataGridView5.Rows[0].Cells[0].Value.ToString();
                            }
                            catch
                            {
                                dataGridView4.Location = new Point(9, 85);
                                dataGridView5.Visible = false;
                            }
                            string command = "select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi,Sil_id from PersonelBilgileri where Sil_id=0";
                            DatagridLoad(command);
                            MessageBox.Show("Basariyla Silindi");
                        }
                    }
                    else if (GüncelleRadio.Checked == true)
                    {
                        PersonelGüncelle per = new PersonelGüncelle(Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value), Form1.KullaniciAdi, Form1.Personel_id, ServerName );
                        per.Show();
                    }
                    string[][] PersonelBilgileri = sql.PersonelOku();
                    comboBox1.Items.Clear();
                    for (int i = 0; i < Convert.ToInt16(PersonelBilgileri[15][0]); i++)
                    {
                        comboBox1.Items.Add(PersonelBilgileri[1][i].ToString());
                        selectedindex = Convert.ToInt16(e.ColumnIndex);

                    }
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            SqlBaglanti sql = new SqlBaglanti(ServerName);
            sql.KullaniciEkle(textBox13.Text, textBox11.Text, comboBox1.Text, sql.PersonelOku()[16][comboBox1.SelectedIndex]);
            MessageBox.Show("Kullanici Eklendi");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlBaglanti sql = new SqlBaglanti(ServerName);
            string[][] KullaniciBilgileri = sql.KullaniciOku();
            int bayrak = 1;
            for (int i = 0; i < Convert.ToInt16(KullaniciBilgileri[3][0]); i++)
            {
                if (sql.PersonelOku()[16][comboBox1.SelectedIndex].CompareTo(KullaniciBilgileri[4][i]) == 0)
                {
                    textBox13.Text = KullaniciBilgileri[0][i];
                    textBox11.Text = KullaniciBilgileri[1][i];
                    textBox11.UseSystemPasswordChar = false;
                    button1.Visible = false;
                    Sil.Visible = true;
                    Güncelle.Visible = true;
                    bayrak = 0;
                }
            }
            if (bayrak == 1)
            {
                textBox11.UseSystemPasswordChar = true;
                textBox13.Text = "";
                textBox11.Text = "";
                button1.Visible = true;
                Sil.Visible = false;
                Güncelle.Visible = false;
            }

        }
        private void Güncelle_Click(object sender, EventArgs e)
        {
            SqlBaglanti sql = new SqlBaglanti(ServerName);
            sql.GüncelleKullanici(comboBox1.Text, textBox11.Text, textBox13.Text, sql.PersonelOku()[16][comboBox1.SelectedIndex]);
            MessageBox.Show("Basariyla Güncelendi");
        }
        private void Sil_Click(object sender, EventArgs e)
        {
            SqlBaglanti sql = new SqlBaglanti(ServerName );
            sql.SilKullanici(Convert.ToInt16(sql.PersonelOku()[16][comboBox1.SelectedIndex]));
            comboBox1.Text = "";
            textBox13.Text = "";
            textBox11.Text = "";
            MessageBox.Show("Kayıt Silindi");
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
        }
        private void İzin_Türü_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
                Sil_id = " IN(0,1)";
            if (checkBox2.Checked == false)
                Sil_id = "=0";
            if (İsimTxT.Text == "")
                izin_Name = "isim";
            if (SoyadTxT.Text == "")
                Soyad = "soyad";
            if (TcTxT.Text == "")
                Tc_No = "tc_no";
            if (İzin_Türü.Text != "Hepsi")
            {
                Rapor_Türü = "'%" + İzin_Türü.Text + "%'";
            }
            if (İzin_Türü.SelectedIndex == 0)
            {
                dataGridView3.Visible = true;
                dataGridView3.Location = new Point(8, 215);
                dataGridView2.Visible = true;
            }
            else
            {
                dataGridView3.Location = new Point(8, 76);
                dataGridView2.Visible = false;
            }
            if (İzin_Türü.SelectedIndex == 4)
            {
                dataGridView2.Visible = true;
                dataGridView3.Visible = false;
            }
            else if(İzin_Türü.SelectedIndex!=0)
            {
                dataGridView2.Visible = false;
                dataGridView3.Visible = true;
            }
            if (İzin_Türü.SelectedIndex == 0)
                Rapor_Türü = "rapor_türü";
            izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where rapor_türü like " + Rapor_Türü + " and Sil_id " + Sil_id + "  and isim like " + izin_Name + " and soyad like " + Soyad + " and tc_no like " + Tc_No + " and rapor_id!=4");
            Günübirlikizindataload("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id from izinBilgileri where rapor_türü like " + Rapor_Türü + " and Sil_id " + Sil_id + "  and isim like " + izin_Name + " and soyad like " + Soyad + " and tc_no like " + Tc_No + " and rapor_id=4");

        }
        private void İsimTxT_TextChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
                Sil_id = " IN(0,1)";
            if (checkBox2.Checked == false)
                Sil_id = "=0";
            if (İzin_Türü.Text == "" || İzin_Türü.Text == "Hepsi")
                Rapor_Türü = "rapor_türü";
            if (İsimTxT.Text == "")
                izin_Name = "isim";
            if (SoyadTxT.Text == "")
                Soyad = "soyad";
            if (TcTxT.Text == "")
                Tc_No = "tc_no";
            if (İsimTxT.Text != "")
                izin_Name = "'%" + İsimTxT.Text + "%'";
            if (İzin_Türü.SelectedIndex == 0)
                Rapor_Türü = "rapor_türü";
            izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where rapor_türü like " + Rapor_Türü + " and Sil_id " + Sil_id + "  and isim like " + izin_Name + " and soyad like " + Soyad + " and tc_no like " + Tc_No + " and rapor_id!=4");
            Günübirlikizindataload("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id from izinBilgileri where rapor_türü like " + Rapor_Türü + " and Sil_id " + Sil_id + "  and isim like " + izin_Name + " and soyad like " + Soyad + " and tc_no like " + Tc_No + " and rapor_id=4");

        }
        private void SoyadTxT_TextChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
                Sil_id = " IN(0,1)";
            if (checkBox2.Checked == false)
                Sil_id = "=0";
            if (İsimTxT.Text == "")
                izin_Name = "isim";
            if (SoyadTxT.Text == "")
                Soyad = "soyad";
            if (TcTxT.Text == "")
                Tc_No = "tc_no";
            if (SoyadTxT.Text != "")
                Soyad = "'%" + SoyadTxT.Text + "%'";
            if (İzin_Türü.SelectedIndex == 0)
                Rapor_Türü = "rapor_türü";
            izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where rapor_türü like " + Rapor_Türü + " and Sil_id " + Sil_id + "  and isim like " + izin_Name + " and soyad like " + Soyad + " and tc_no like " + Tc_No + " and rapor_id!=4");
            Günübirlikizindataload("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id from izinBilgileri where rapor_türü like " + Rapor_Türü + " and Sil_id " + Sil_id + "  and isim like " + izin_Name + " and soyad like " + Soyad + " and tc_no like " + Tc_No + " and rapor_id=4");

        }
        private void TcTxT_TextChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
                Sil_id = " IN(0,1)";
            if (checkBox2.Checked == false)
                Sil_id = "=0";
            if (İsimTxT.Text == "")
                izin_Name = "isim";
            if (SoyadTxT.Text == "")
                Soyad = "soyad";
            if (TcTxT.Text == "")
                Tc_No = "tc_no";
            if (TcTxT.Text != "")
                Tc_No = "'%" + TcTxT.Text + "%'";
            if (İzin_Türü.SelectedIndex == 0)
                Rapor_Türü = "rapor_türü";
            izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where rapor_türü like " + Rapor_Türü + " and Sil_id " + Sil_id + "  and isim like " + izin_Name + " and soyad like " + Soyad + " and tc_no like " + Tc_No + " and rapor_id!=4");
            Günübirlikizindataload("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id from izinBilgileri where rapor_türü like " + Rapor_Türü + " and Sil_id " + Sil_id + "  and isim like " + izin_Name + " and soyad like " + Soyad + " and tc_no like " + Tc_No + " and rapor_id=4");

        }
        private void İzin_Türü_TextChanged(object sender, EventArgs e)
        {
            if (İzin_Türü.Text == "")
                İzin_Türü.SelectedIndex = 0;
        }

        private void CbxYıl_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            CbxAy.Text = "";
            if (CbxAy.Text == "")
            {
                BtnExcel.Enabled = false;
            }
            CbxAy.Enabled = true;
            CbxAy.Items.Clear();

            myCon.Open();
            CbxAy.Items.Add(23);
            SqlCommand cmd = new SqlCommand("select * from dbo.izinBilgileri where bas_tar like'%" + CbxYıl.Items[CbxYıl.SelectedIndex] + "%' and Sil_id=0", myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string[] str = dr["bas_tar"].ToString().Split('.');
                int bayrak = 0;
                for (int i = 0; i < CbxAy.Items.Count; i++)
                {
                    if (CbxAy.Items[i].ToString().CompareTo(Aylar[int.Parse(str[1]) - 1]) == 0)
                    {
                        bayrak = 1;
                    }
                }
                if (bayrak == 0)
                    CbxAy.Items.Add(Aylar[int.Parse(str[1]) - 1]);
            }
            CbxAy.Items.RemoveAt(0);
            myCon.Close();
        }
        string Tarih = "";
        private void CbxAy_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnExcel.Enabled = true;
            for (int i = 0; i < Aylar.Length; i++)
            {
                if (Aylar[i] == CbxAy.Items[CbxAy.SelectedIndex].ToString())
                {
                    Tarih = (i + 1) + "." + CbxYıl.Items[CbxYıl.SelectedIndex].ToString();
                    if (i + 1 < 10)
                        Tarih = ".0" + Tarih;
                    else
                        Tarih = "." + Tarih;
                    try
                    {
                        dataGridView4.Visible = true;
                        izinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where bas_tar like '%" + Tarih + "%' and Sil_id like 0 and rapor_id!=4 order by  rapor_türü");
                        dataGridView4.Rows[0].Cells[0].Value.ToString();
                    }
                    catch
                    {
                        dataGridView4.Visible = false;
                    }
                    try
                    {
                        dataGridView4.Location = new Point(8, 235);
                        dataGridView5.Visible = true;
                        GünlükizinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id  from izinBilgileri where  bas_tar like '%" + Tarih + "%' and Sil_id=0 and rapor_id=4");
                        dataGridView5.Rows[0].Cells[0].Value.ToString();
                    }
                    catch
                    {
                        dataGridView4.Location = new Point(9, 85);
                        dataGridView5.Visible = false;
                    }
                }
            }

        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            CbxAy.Text = "";
            CbxYıl.Text = "";
            BtnExcel.Enabled = false;
            CbxAy.Enabled = false;
            int temp = 0;
            Directory.CreateDirectory(CbxYıl.Items[CbxYıl.SelectedIndex].ToString());
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            worksheet = workbook.ActiveSheet;
            try
            {
                string a=dataGridView4.Rows[0].Cells[0].Value.ToString();
                for (int i = 2; i < dataGridView4.Columns.Count; i++)
                {
                    worksheet.get_Range("A1", "J1").Font.Bold = true;
                    worksheet.Cells[1, i - 1] = dataGridView4.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView4.Rows.Count - 1; i++)
                {
                    for (int j = 1; j < dataGridView4.Columns.Count - 1; j++)
                    {
                        worksheet.Cells[i + 2, j] = dataGridView4.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            catch
            {
                temp = 1;
            }

            try
            {
                dataGridView5.Rows[0].Cells[0].Value.ToString();
                for (int i = 2; i < dataGridView5.Columns.Count; i++)
                {
                    int temp2;
                    if (temp == 0)
                        temp2 = dataGridView4.Rows.Count + 1;
                    else
                        temp2 = dataGridView4.Rows.Count;

                    worksheet.get_Range("A" + temp2 + "", "J" + temp2 + "").Font.Bold = true;
                    worksheet.Cells[temp2, i - 1] = dataGridView5.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView5.Rows.Count - 1; i++)
                {
                    int temp2;
                    if (temp == 0)
                        temp2 = dataGridView4.Rows.Count + 2;
                    else
                        temp2 = dataGridView4.Rows.Count+1;
                    for (int j = 1; j < dataGridView5.Columns.Count - 1; j++)
                    {
                        worksheet.Cells[temp2 + i, j] = dataGridView5.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            catch
            {

            }
            for (char i = 'A'; i <= 'I'; i++)
            {
                Excel.Range er = worksheet.get_Range("" + i + ":" + i + "", System.Type.Missing);
                er.EntireColumn.ColumnWidth = 20;
            }
            worksheet.PageSetup.PrintGridlines = true;
            workbook.SaveAs(Application.StartupPath + @"\" + CbxYıl.Items[CbxYıl.SelectedIndex] + @"\" + CbxAy.Items[CbxAy.SelectedIndex] + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            app.Quit();
            MessageBox.Show(CbxAy.Items[CbxAy.SelectedIndex] + ".xlsx basariyla olusturulmustur");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
                PersonelSil_id = " IN(0,1)";
            if (checkBox1.Checked == false)
                PersonelSil_id = "=0";
            if (txtPersonelisim.Text == "")
                Personel_isim = "Personel_Name";
            if (txtPersonelSoyad.Text == "")
                Personel_soyad = "Personel_Surname";
            if (txtPersonelTc.Text == "")
                Personel_Tc = "Personel_Tc";
            string command = "select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi,Sil_id from PersonelBilgileri where Sil_id" + PersonelSil_id + " and Personel_Name like " + Personel_isim + " and Personel_Surname like " + Personel_soyad + " and Personel_Tc like " + Personel_Tc;
            DatagridLoad(command);
            if (checkBox1.Checked == true)
            {
                Silradio.Checked = false;
                GüncelleRadio.Checked = false;
                izin.Checked = false;
                Silradio.Enabled = false;
                GüncelleRadio.Enabled = false;
                izin.Enabled = false;

            }
            else
            {

                Silradio.Enabled = true;
                GüncelleRadio.Enabled = true;
                izin.Enabled = true;

            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                İzinSil.Enabled = false;
                İzinGüncelle.Enabled = false;
                İzinSil.Checked = false;
                İzinGüncelle.Checked = false;
            }
            if (checkBox2.Checked == false)
            {
                İzinSil.Enabled = true;
                İzinGüncelle.Enabled = true;
            }
            if (checkBox2.Checked == true)
                Sil_id = " IN(0,1)";
            if (checkBox2.Checked == false)
                Sil_id = "=0";
            if (İsimTxT.Text == "")
                izin_Name = "isim";
            if (SoyadTxT.Text == "")
                Soyad = "soyad";
            if (TcTxT.Text == "")
                Tc_No = "tc_no";
            if (SoyadTxT.Text != "")
                Soyad = "'%" + SoyadTxT.Text + "%'";
            if (İzin_Türü.SelectedIndex == 0)
                Rapor_Türü = "rapor_türü";
            izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where rapor_türü like " + Rapor_Türü + " and Sil_id " + Sil_id + "  and isim like " + izin_Name + " and soyad like " + Soyad + " and tc_no like " + Tc_No + " and rapor_id!=4");
            Günübirlikizindataload("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id from izinBilgileri where rapor_türü like " + Rapor_Türü + " and Sil_id " + Sil_id + "  and isim like " + izin_Name + " and soyad like " + Soyad + " and tc_no like " + Tc_No + " and rapor_id=4");

        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            SqlBaglanti sql = new SqlBaglanti(ServerName);
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            string Personel_id = "";
            string raptürü = "";
            if (e.RowIndex != -1)
            {
                if (İzinSil.Checked == true)
                {
                    DialogResult sonuc = MessageBox.Show("Silmek istediginizden eminmisiniz?", "Dikkat", MessageBoxButtons.YesNo);
                    if (sonuc == DialogResult.Yes)
                    {
                        sql.Silizin(dataGridView3.CurrentRow.Cells[0].Value.ToString(), "izin_id");
                        myCon.Open();
                        SqlCommand cmd = new SqlCommand("select * from dbo.izinBilgileri where izin_id=" + dataGridView3.CurrentRow.Cells[0].Value.ToString(), myCon);
                        SqlDataReader dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            Personel_id = dr["personel_id"].ToString();
                            raptürü = dr["rapor_türü"].ToString();
                        }
                        myCon.Close();
                        if (raptürü == "YILLIKİZİN")
                        {
                            myCon.Open();
                            int KullanılanGün = 0;
                            int ToplamGün = 0;
                            SqlCommand cmd2 = new SqlCommand("select * from dbo.PersonelBilgileri where Personel_id=" + Personel_id, myCon);
                            SqlDataReader dr2 = cmd2.ExecuteReader(); ;
                            while (dr2.Read())
                            {
                                ToplamGün = Convert.ToInt16(dr2["Yıllıkizin_Süresi"]) + Convert.ToInt16(dataGridView3.CurrentRow.Cells[6].Value);
                                KullanılanGün = Convert.ToInt16(dr2["Kullanılan_Süre"]) - Convert.ToInt16(dataGridView3.CurrentRow.Cells[6].Value);
                            }
                            myCon.Close();
                            myCon.Open();
                            SqlCommand cmd3 = new SqlCommand("  Update PersonelBilgileri set Yıllıkizin_Süresi=" + ToplamGün + ",Kullanılan_Süre=" + KullanılanGün + " where personel_id=" + Personel_id, myCon);
                            cmd3.ExecuteNonQuery();
                            myCon.Close();
                        }
                        try
                        {
                            dataGridView4.Visible = true;
                            izinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where bas_tar like '%" + Tarih + "%' and Sil_id like 0 and rapor_id!=4 order by  rapor_türü");
                            dataGridView4.Rows[0].Cells[0].Value.ToString();
                        }
                        catch
                        {
                            dataGridView4.Visible = false;
                        }
                        try
                        {
                            dataGridView4.Location = new Point(8, 235);
                            dataGridView5.Visible = true;
                            GünlükizinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id  from izinBilgileri where  bas_tar like '%" + Tarih + "%' and Sil_id=0 and rapor_id=4");
                            dataGridView5.Rows[0].Cells[0].Value.ToString();
                        }
                        catch
                        {
                            dataGridView4.Location = new Point(9, 85);
                            dataGridView5.Visible = false;
                        }
                        izinDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,rap_sür,bitis_tar,adres,Sil_id from izinBilgileri where Sil_id=0 and rapor_id!=4 order by rapor_türü");
                        string command = "select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi, Sil_id from PersonelBilgileri where Sil_id=0";
                        DatagridLoad(command);
                        MessageBox.Show("Basariyla Silindi");
                    }
                }
                else if (İzinGüncelle.Checked == true)
                {
                    izinGüncelle per = new izinGüncelle(Convert.ToInt16(dataGridView3.CurrentRow.Cells[0].Value), ServerName );
                    per.Show();
                }
            }

        }

        private void dataGridView3_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == 0 || comboBox2.SelectedIndex == 2)
            {
                textBox7.Text = "";
                textBox7.Enabled = true;
            }
            else
            {

                textBox7.Enabled = false;
                textBox7.Text = "X";
            }
            if (comboBox2.SelectedIndex == 1 || comboBox2.SelectedIndex == 2)
            {
                cbxEsdurumu.Enabled = false;
                cbxEsdurumu.Text = "X";
            }
            else if (comboBox2.SelectedIndex == 0)
            {
                cbxEsdurumu.Text = "Çalışmıyor";
                cbxEsdurumu.Enabled = true;
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex == 0)
            {
                textBox12.Text = "";
                textBox12.Enabled = true;
            }
            else
            {
                textBox12.Enabled = false;
                textBox12.Text = "X";
            }
        }

        private void cbxEsdurumu_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnraporlamaexcel_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory("Raporlama");
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            worksheet = workbook.ActiveSheet;
            for (int i = 2; i < datagridRaporlama.Columns.Count + 1; i++)
            {
                worksheet.get_Range("A1", "J1").Font.Bold = true;
                worksheet.Cells[1, i - 1] = datagridRaporlama.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < datagridRaporlama.Rows.Count - 1; i++)
            {
                for (int j = 1; j < datagridRaporlama.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j] = datagridRaporlama.Rows[i].Cells[j].Value.ToString();
                }
            }
            for (char i = 'A'; i <= 'I'; i++)
            {
                Excel.Range er = worksheet.get_Range("" + i + ":" + i + "", System.Type.Missing);
                er.EntireColumn.ColumnWidth = 20;
            }
            workbook.SaveAs(Application.StartupPath + @"\" + "Raporlama" + @"\" + cbxRaporlama.Items[cbxRaporlama.SelectedIndex] + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            app.Visible = true;
            MessageBox.Show(cbxRaporlama.Items[cbxRaporlama.SelectedIndex] + ".xlsx basariyla olusturulmustur");
        }

        private void cbxRaporlama_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (cbxRaporlama.SelectedIndex == 0)
            {
                PersonelRaporlama("select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,Personel_Göreviyeri from PersonelBilgileri where Sil_id=0 ", 0);
            }
            else if (cbxRaporlama.SelectedIndex == 1)
            {
                PersonelRaporlama("select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,HesapNo from PersonelBilgileri where Sil_id=0 ", 1);
            }
            else if (cbxRaporlama.SelectedIndex == 2)
            {
                PersonelRaporlama("select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,EvlilikDurumu,ÇocukSayisi,EsDurumu,ÖzürlülükDurumu,ÖzürlülükYüzdesi from PersonelBilgileri where Sil_id=0 ", 2);
            }
            else if (cbxRaporlama.SelectedIndex == 3)
            {
                PersonelRaporlama("select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,Personel_TelNumber from PersonelBilgileri where Sil_id=0 ", 3);
            }
            else if (cbxRaporlama.SelectedIndex == 4)
            {
                PersonelRaporlama("select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,Personel_SGK from PersonelBilgileri where Sil_id=0 ", 4);
            }
        }

        private void txtPersonelTc_TextChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
                PersonelSil_id = " IN(0,1)";
            if (checkBox1.Checked == false)
                PersonelSil_id = "=0";
            if (txtPersonelisim.Text == "")
                Personel_isim = "Personel_Name";
            if (txtPersonelSoyad.Text == "")
                Personel_soyad = "Personel_Surname";
            if (txtPersonelTc.Text == "")
                Personel_Tc = "Personel_Tc";
            if (txtPersonelTc.Text != "")
                Personel_Tc = "'%" + txtPersonelTc.Text + "%'";
            string command = "select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi,Sil_id from PersonelBilgileri where Sil_id" + PersonelSil_id + " and Personel_Name like " + Personel_isim + " and Personel_Surname like " + Personel_soyad + " and Personel_Tc like " + Personel_Tc;
            DatagridLoad(command);
        }

        private void txtPersonelisim_TextChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
                PersonelSil_id = " IN(0,1)";
            if (checkBox1.Checked == false)
                PersonelSil_id = "=0";
            if (txtPersonelisim.Text == "")
                Personel_isim = "Personel_Name";
            if (txtPersonelSoyad.Text == "")
                Personel_soyad = "Personel_Surname";
            if (txtPersonelTc.Text == "")
                Personel_Tc = "Personel_Tc";
            if (txtPersonelisim.Text != "")
                Personel_isim = "'%" + txtPersonelisim.Text + "%'";
            string command = "select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi,Sil_id from PersonelBilgileri where Sil_id" + PersonelSil_id + " and Personel_Name like " + Personel_isim + " and Personel_Surname like " + Personel_soyad + " and Personel_Tc like " + Personel_Tc;
            DatagridLoad(command);
        }

        private void txtPersonelSoyad_TextChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
                PersonelSil_id = " IN(0,1)";
            if (checkBox1.Checked == false)
                PersonelSil_id = "=0";
            if (txtPersonelisim.Text == "")
                Personel_isim = "Personel_Name";
            if (txtPersonelSoyad.Text == "")
                Personel_soyad = "Personel_Surname";
            if (txtPersonelTc.Text == "")
                Personel_Tc = "Personel_Tc";
            if (txtPersonelSoyad.Text != "")
                Personel_soyad = "'%" + txtPersonelSoyad.Text + "%'";
            string command = "select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi,Sil_id from PersonelBilgileri where Sil_id" + PersonelSil_id + " and Personel_Name like " + Personel_isim + " and Personel_Surname like " + Personel_soyad + " and Personel_Tc like " + Personel_Tc;
            DatagridLoad(command);
        }
        int hata = 0;
        int bosluk = 0;
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                if (textBox1.Text.Length > 11)
                {
                    MessageBox.Show("Lütfen Tc Numaranizi Kontrol Ediniz");
                }
                try
                {
                    hata = 0;
                    Convert.ToDouble(textBox1.Text);
                }
                catch
                {
                    hata = 1;
                    MessageBox.Show("Lütfen Tc Numaranizi Kontrol Ediniz");
                }

                bosluk = 0;
            }
            else
                bosluk = 1;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (textBox4.Text != "")
            {
                try
                {
                    hata = 0;
                    Convert.ToDouble(textBox4.Text);
                }
                catch
                {
                    hata = 1;
                    MessageBox.Show("Lütfen SGK Numaranizi Kontrol Ediniz");
                }
                bosluk = 0;
            }
            else
                bosluk = 1;
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if (textBox9.Text.Length <= 2)
            {
                textBox9.Text = "05";
                textBox9.SelectionStart = 3;
            }
            if (textBox9.Text[0] != '0' || textBox9.Text[1] != '5')
            {
                textBox9.Text = "05";
                textBox9.SelectionStart = 3;
            }
            string txt = "";
            try
            {
                if (textBox9.Text[4] != ' ')
                {
                    for (int i = 0; i < textBox9.Text.Length; i++)
                    {
                        if (i == 4)
                            txt += " " + textBox9.Text[i].ToString();
                        else
                            txt += textBox9.Text[i].ToString();
                    }
                    textBox9.Text = txt;
                    textBox9.SelectionStart = 6;
                }

            }
            catch
            {
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (textBox7.Enabled != false)
            {
                if (textBox7.Text != "")
                {
                    try
                    {
                        hata = 0;
                        Convert.ToDouble(textBox7.Text);
                    }
                    catch
                    {
                        hata = 1;
                        MessageBox.Show("Lütfen Cocuk Sayisini Kontrol Ediniz");
                    }
                    bosluk = 0;
                }
                else
                    bosluk = 1;

            }

        }

        private void txthesapno_TextChanged(object sender, EventArgs e)
        {
            if (txthesapno.Text != "")
            {
                try
                {
                    hata = 0;
                    Convert.ToDouble(txthesapno.Text);
                }
                catch
                {
                    hata = 1;
                    MessageBox.Show("Lütfen Hesap Numaranizi Kontrol Ediniz");
                }
                bosluk = 0;
            }
            else
                bosluk = 1;
        }

        private void txtmail_TextChanged(object sender, EventArgs e)
        {
            int bayrak = 0;
            for (int i = 0; i < txtmail.Text.Length; i++)
            {
                if (txtmail.Text[i] == '@')
                    bayrak = 1;
            }
            if (bayrak == 1)
                hata = 0;
            else
                hata = 1;
            bayrak = 0;
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            if (textBox12.Enabled != false)
            {
                if (textBox12.Text != "")
                {
                    try
                    {

                        if (Convert.ToDouble(textBox12.Text) > 100)
                        {
                            MessageBox.Show("Özürlülük Dereceniz 100 ü Geçemez");
                            hata = 1;
                        }
                        hata = 0;
                    }
                    catch
                    {

                        MessageBox.Show("Lütfen Özürlülü Derecenizi Kontrol Ediniz");
                        hata = 1;
                    }
                    bosluk = 0;
                }
                else
                    bosluk = 1;
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            SqlBaglanti sql = new SqlBaglanti(ServerName);
           string Personel_id = "";
            string raptürü = "";
            if (e.RowIndex != -1)
            {
                if (İzinSil.Checked == true)
                {
                    DialogResult sonuc = MessageBox.Show("Silmek istediginizden eminmisiniz?", "Dikkat", MessageBoxButtons.YesNo);
                    if (sonuc == DialogResult.Yes)
                    {
                        SqlConnection myCon = new SqlConnection(BaglantiOlustur());
                        sql.Silizin(dataGridView2.CurrentRow.Cells[0].Value.ToString(), "izin_id");
                        myCon.Open();
                        SqlCommand cmd = new SqlCommand("select * from dbo.izinBilgileri where izin_id=" + dataGridView2.CurrentRow.Cells[0].Value.ToString(), myCon);
                        SqlDataReader dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            Personel_id = dr["personel_id"].ToString();
                            raptürü = dr["rapor_türü"].ToString();
                        }
                        myCon.Close();
                        try
                        {
                            dataGridView4.Location = new Point(8, 235);
                            dataGridView5.Visible = true;
                            GünlükizinExcelDataLoad("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id  from izinBilgileri where  bas_tar like '%" + Tarih + "%' and Sil_id=0 and rapor_id=4");
                            dataGridView5.Rows[0].Cells[0].Value.ToString();
                        }
                        catch
                        {
                            dataGridView4.Location = new Point(9, 85);
                            dataGridView5.Visible = false;
                        }
                        Günübirlikizindataload("select izin_id,tc_no,isim,soyad,rapor_türü,bas_tar,Baslama_Saati,Kac_Saat,Bitis_Saati,Mazeret,Sil_id  from izinBilgileri where Sil_id=0 and rapor_id=4");
                        string command = "select Personel_id,Personel_Tc,Personel_Name,Personel_Surname,PreYear,ThisYear,total,Kullanılan_Süre,Yıllıkizin_Süresi, Sil_id from PersonelBilgileri where Sil_id=0";
                        DatagridLoad(command);
                        MessageBox.Show("Basariyla Silindi");
                    }
                }
                else if (İzinGüncelle.Checked == true)
                {

                    GünüBrilikGüncelleme gün = new GünüBrilikGüncelleme(Convert.ToInt16(dataGridView2.CurrentRow.Cells[0].Value), ServerName );
                    gün.Show();
                }
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked == true)
            {
                groupBox1.Visible = true;
                groupBox2.Visible = false;
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked == true)
            {
                groupBox1.Visible = false;
                groupBox2.Visible = true;
            }
        }

        private void radioButton1_CheckedChanged_1(object sender, EventArgs e)
        {
            BirinciLogo.Visible = true;
            İkinciLogo.Visible = false;
            ücüncüLogo.Visible = false;
        }

        private void radioButton2_CheckedChanged_1(object sender, EventArgs e)
        {
            BirinciLogo.Visible = true;
            İkinciLogo.Visible = true;
            ücüncüLogo.Visible = false;
        }

        private void radioButton3_CheckedChanged_1(object sender, EventArgs e)
        {
            BirinciLogo.Visible = true;
            İkinciLogo.Visible = true;
            ücüncüLogo.Visible = true;
        }
        int hata1 = 0;
        int hata2 = 0;
        int hata3 = 0;
        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Resim Dosylari|*.png;*.jpg*";
            openFileDialog1.ShowDialog();
            if (openFileDialog1.FileName != "openFileDialog1")
            {
                pictureBox1.ImageLocation = openFileDialog1.FileName;
                Logo1 = openFileDialog1.FileName;
                imagedatalogo1 = ReadImageFile(Logo1);
                hata1 = 1;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            openFileDialog2.Filter = "Resim Dosylari|*.png;*.jpg*";
            openFileDialog2.ShowDialog();
            if (openFileDialog2.FileName != "openFileDialog2")
            {
                pictureBox2.ImageLocation = openFileDialog2.FileName;
                Logo2 = openFileDialog2.FileName;
                imagedatalogo2 = ReadImageFile(Logo2);
                hata2 = 1;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            openFileDialog3.Filter = "Resim Dosylari|*.png;*.jpg*";
            openFileDialog3.ShowDialog();
            if (openFileDialog3.FileName != "openFileDialog3")
            {
                pictureBox3.ImageLocation = openFileDialog3.FileName;
                Logo3 = openFileDialog3.FileName;
                imagedatalogo3 = ReadImageFile(Logo3);
                hata3 = 1;
            }
        }
        string Logo1="";
        string Logo2="";
        string Logo3="";
        byte[] imagedatalogo1;
        byte[] imagedatalogo2;
        byte[] imagedatalogo3;
        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlCommand cmd = new SqlCommand("  Update Ayarlar set PersonelOtomasyonName='" + txtotomüdürname.Text + "', PersonelOtomasyonSurname='"+txtotomüdürsurname.Text+"'", myCon);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Basariyla Güncelendi");
            myCon.Close();
        }
        private void button8_Click(object sender, EventArgs e)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
                if (radioButton1.Checked == true)
                {
                    if (hata1 == 1)
                    {
                        string vsql = string.Format("UPDATE  Ayarlar set LogoSayisi = @sayi, logo1 = @DATA");
                        myCon.Open();
                        SqlCommand azert = new SqlCommand(vsql, myCon);
                        azert.Parameters.Add("@sayi", '1');
                        azert.Parameters.Add("@DATA", imagedatalogo1);
                        azert.ExecuteNonQuery();
                        MessageBox.Show("Basariyla Güncelendi");
                        myCon.Close();
                    }
                    else
                    {
                        MessageBox.Show("Logo Seciniz...","Dikkat");
                    }

                }
                else if (radioButton2.Checked == true)
                {
                    if (hata1 == 1 && hata2 == 1)
                    {
                        string vsql = string.Format("UPDATE  Ayarlar set LogoSayisi = @sayi, logo1 = @DATA,logo2=@DATA2");
                        myCon.Open();
                        SqlCommand azert = new SqlCommand(vsql, myCon);
                        azert.Parameters.Add("@sayi", '2');
                        azert.Parameters.Add("@DATA", imagedatalogo1);
                        azert.Parameters.Add("@DATA2", imagedatalogo2);
                        azert.ExecuteNonQuery();
                        MessageBox.Show("Basariyla Güncelendi");
                        myCon.Close();
                    }
                    else
                    {
                        MessageBox.Show("Logo Seciniz...", "Dikkat");
                    }
                   
                }
                else if (radioButton3.Checked == true)
                {
                    if (hata1 == 1 && hata2 == 1 && hata3 == 1)
                    {
                        string vsql = string.Format("UPDATE  Ayarlar set LogoSayisi = @sayi, logo1 = @DATA,logo2=@DATA2,logo3=@DATA3");
                        myCon.Open();
                        SqlCommand azert = new SqlCommand(vsql, myCon);
                        azert.Parameters.Add("@sayi", '3');
                        azert.Parameters.Add("@DATA", imagedatalogo1);
                        azert.Parameters.Add("@DATA2", imagedatalogo2);
                        azert.Parameters.Add("@DATA3", imagedatalogo3);
                        azert.ExecuteNonQuery();
                        MessageBox.Show("Basariyla Güncelendi");
                        myCon.Close();
                    }
                    else
                    {
                        MessageBox.Show("Logo Seciniz...", "Dikkat");
                    }
                
                }

        }
        public byte[] ReadImageFile(string imageLocation)
        {
            byte[] imageData = null;
            FileInfo fileInfo = new FileInfo(imageLocation);
            long imageFileLength = fileInfo.Length;
            FileStream fs = new FileStream(imageLocation, FileMode.Open, FileAccess.Read);
            BinaryReader br = new BinaryReader(fs);
            imageData = br.ReadBytes((int)imageFileLength);
            return imageData;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            excel.Visible = true;

            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);

            Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];


            var boldformat = sheet1.get_Range("A1", "T1");
            var m_objfont = boldformat.Font;
            m_objfont.Bold = true;





            //

            int StartCol = 1;

            int StartRow = 1;

            for (int j = 0; j < datagridRaporlama.Columns.Count; j++)
            {

                Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];

                myRange.Value2 = datagridRaporlama.Columns[j].HeaderText;

            }

            StartRow++;

            for (int i = 0; i < datagridRaporlama.Rows.Count; i++)
            {

                for (int j = 0; j < datagridRaporlama.Columns.Count; j++)
                {

                    try
                    {

                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];

                        myRange.Value2 = datagridRaporlama[j, i].Value == null ? "" : datagridRaporlama[j, i].Value;

                    }

                    catch
                    {

                        ;

                    }

                }

            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }
    }
}
