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
    public partial class GünlükYazdır : Form
    {
        string Tarih, baslamasaati, bitissaati, isim, Mazaret, ServerName;
        public GünlükYazdır(string Tarih,string baslamasaati,string bitissaati,string isim,string Mazaret,string ServerName)
        {
            this.ServerName = ServerName;
            this.Tarih = Tarih;
            this.baslamasaati = baslamasaati;
            this.bitissaati = bitissaati;
            this.isim = isim;
            this.Mazaret = Mazaret;
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
        string logosayisi = "";
        private void GünlükYazdır_Load(object sender, EventArgs e)
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
                label5.Text = dr["PersonelOtomasyonName"] + " " + dr["PersonelOtomasyonSurname"];
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
            label1.Text = Tarih;
            label2.Text = baslamasaati;
            label3.Text = bitissaati;
            label4.Text = isim;
            textBox2.Text = Mazaret;
            printDocument1.Print();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawImage(yazdırılacakresim.Image, new Point());
            if (logosayisi == "1")
            {
                e.Graphics.DrawImage(pictureBox4.Image, 286, 36, 273, 106);
            }
            else if (logosayisi == "2")
            {
                e.Graphics.DrawImage(pictureBox1.Image, 138, 36, 273, 106);
                e.Graphics.DrawImage(pictureBox2.Image, 411, 36, 273, 106);
            }
            else if (logosayisi == "3")
            {
                e.Graphics.DrawImage(pictureBox3.Image, 15, 36, 273, 106);
                e.Graphics.DrawImage(pictureBox4.Image, 286, 36, 273, 106);
                e.Graphics.DrawImage(pictureBox5.Image, 558, 36, 273, 106);

            }
            e.Graphics.DrawString(label1.Text, new Font("Arial", 10), Brushes.Black, 92, 288);
            e.Graphics.DrawString(label2.Text, new Font("Arial", 10), Brushes.Black, 238, 288);
            e.Graphics.DrawString(label3.Text, new Font("Arial", 10), Brushes.Black, 322, 288);
            e.Graphics.DrawString(label4.Text, new Font("Arial", 10), Brushes.Black, 152, 468);
            e.Graphics.DrawString(textBox2.Text, new Font("Arial", 12), Brushes.Black, 214, 335);
            e.Graphics.DrawString(label5.Text, new Font("Arial", 12), Brushes.Black, 500, 463);

        }



    }
}
