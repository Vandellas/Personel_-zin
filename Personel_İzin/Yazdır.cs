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
    public partial class Yazdır : Form
    {
        string AdıSoyadi;
        string Görevi = "Bilgi işlem personeli";
        string isegirisTarihi;
        string izinistemesebebi;
        string izninaitolduguyil;
        string izninbaslayacagitarih;
        string iznindönüstarihi;
        string izinsüresi;
        string izningeçecegiadres;
        string telNo;
        int ResimNo;
        string ServerName;
        public  string BaglantiOlustur()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = ServerName;
            builder.InitialCatalog = "personel_izin";
            builder.IntegratedSecurity = true;
            return builder.ConnectionString;
        }
        public Yazdır(string AdıSoyadi,string isegirisTarihi, string izinistemesebebi,string izninaitolduguyil,string izninbaslayacagitarih, string iznindönüstarihi, string izinsüresi, string izningeçecegiadres, string telNo,int ResimNo,string ServerName )
        {
            this.ServerName = ServerName;
            this.AdıSoyadi = AdıSoyadi;
            this.isegirisTarihi = isegirisTarihi;
            this.izinistemesebebi = izinistemesebebi;
            this.izninaitolduguyil = izninaitolduguyil;
            this.izninbaslayacagitarih = izninbaslayacagitarih;
            this.iznindönüstarihi = iznindönüstarihi;
            this.izinsüresi = izinsüresi;
            this.izningeçecegiadres = izningeçecegiadres;
            this.telNo = telNo;
          
            this.ResimNo = ResimNo;
            InitializeComponent();
        }
        string logosayisi = "";
        public void LogoLoad()
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            pictureBox1.Visible = false;
            pictureBox2.Visible = false;
            pictureBox3.Visible = false;
            pictureBox4.Visible = false;
            pictureBox5.Visible = false;
            myCon.Open();
            SqlCommand cmd = new SqlCommand("select*from Ayarlar", myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
             //   label5.Text = dr["PersonelOtomasyonName"] + " " + dr["PersonelOtomasyonSurname"];
               logosayisi= dr["LogoSayisi"].ToString();
                if (logosayisi == "1")
                {
                    pictureBox4.BringToFront();
                    pictureBox4.Visible = true;
                    pictureBox4.Image = Image.FromStream(dr.GetSqlBytes(1).Stream);
                }
                else if (logosayisi == "2")
                {
                    pictureBox1.BringToFront();
                    pictureBox2.BringToFront();
                    pictureBox1.Visible = true;
                    pictureBox2.Visible = true;
                    pictureBox1.Image = Image.FromStream(dr.GetSqlBytes(1).Stream);
                    pictureBox2.Image = Image.FromStream(dr.GetSqlBytes(2).Stream);
                }
                else if (logosayisi == "3")
                {
                    pictureBox3.BringToFront();
                    pictureBox4.BringToFront();
                    pictureBox5.BringToFront();
                    pictureBox3.Visible = true;
                    pictureBox4.Visible = true;
                    pictureBox5.Visible = true;
                    pictureBox3.Image = Image.FromStream(dr.GetSqlBytes(1).Stream);
                    pictureBox4.Image = Image.FromStream(dr.GetSqlBytes(2).Stream);
                    pictureBox5.Image = Image.FromStream(dr.GetSqlBytes(3).Stream);
                }
            }
            myCon.Close();
        }
        private void Yazdır_Load(object sender, EventArgs e)
        {
            LogoLoad();
            AdSoyad.Text= AdıSoyadi;
            giristar.Text = isegirisTarihi;
            izinsebeb.Text = izinistemesebebi;
            yıl.Text = izninaitolduguyil;
            bastar.Text= izninbaslayacagitarih;
            döntar.Text = iznindönüstarihi;
            rapsür.Text = izinsüresi;
            adres.Text= izningeçecegiadres;
            görev.Text = Görevi;
            label1.Text = telNo;
            if (ResimNo == 1)
                izintürü.Text = "RAPOR";
            else if (ResimNo == 5)
               izintürü.Text = "İDARİ İZİN";
            else if (ResimNo == 3)
                izintürü.Text = "YILLIK İZİN";
            else if (ResimNo == 2)
                izintürü.Text = "SEVK";
            else if (ResimNo == 4)
                izintürü.Text = "MAZERET İZNİ";

            printDocument1.Print();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawImage(yazdırılacakresim.Image, new Point());
            if (logosayisi == "1")
            {
                e.Graphics.DrawImage(pictureBox4.Image, 296, 56,271, 173);
            }
            else if (logosayisi == "2")
            {
                e.Graphics.DrawImage(pictureBox1.Image, 127, 56, 271, 173);
                e.Graphics.DrawImage(pictureBox2.Image, 404, 56,271, 173);
            }
            else if (logosayisi == "3")
            {
                e.Graphics.DrawImage(pictureBox3.Image, 19, 56, 271, 173);
                e.Graphics.DrawImage(pictureBox4.Image, 296, 56, 271, 173);
                e.Graphics.DrawImage(pictureBox5.Image,573, 56, 271, 173);

            }
            e.Graphics.DrawString(AdıSoyadi, new Font("Arial", 12), Brushes.Black, 364, 468);
            e.Graphics.DrawString(izintürü.Text, new Font("Arial", 15,FontStyle.Bold), Brushes.Black, 364, 358);
            e.Graphics.DrawString(Görevi, new Font("Arial", 12), Brushes.Black, 364, 488);
            e.Graphics.DrawString(isegirisTarihi, new Font("Arial", 12), Brushes.Black, 364, 508);
            e.Graphics.DrawString(izinistemesebebi, new Font("Arial", 12), Brushes.Black, 364, 528);
            e.Graphics.DrawString(izninaitolduguyil, new Font("Arial", 12), Brushes.Black, 364, 548);
            e.Graphics.DrawString(izninbaslayacagitarih, new Font("Arial", 12), Brushes.Black, 364, 568);
            e.Graphics.DrawString(iznindönüstarihi, new Font("Arial", 12), Brushes.Black, 364, 588);
            e.Graphics.DrawString(izinsüresi, new Font("Arial", 12), Brushes.Black, 364, 608);
            e.Graphics.DrawString(izningeçecegiadres, new Font("Arial", 12), Brushes.Black, 364, 631);
            e.Graphics.DrawString(telNo, new Font("Arial", 12), Brushes.Black, 364, 653);
        }
    }
}
