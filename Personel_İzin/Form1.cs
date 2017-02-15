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
using System.IO;
using System.Data.SqlClient;
using System.IO;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Smo;
using System.Text.RegularExpressions;
namespace Personel_İzin
{
    public partial class Form1 : Form
    {
       
        string[][] KullaniciBilgileri;
        string ServerName;
        public Form1(string ServerName)
        {
            this.ServerName = ServerName;
            InitializeComponent();
        }
        public static string KullaniciAdi = "";
        public static int Personel_id = 0;

        private void button1_Click(object sender, EventArgs e)
        {
            SqlBaglanti sql = new SqlBaglanti(ServerName);
            KullaniciBilgileri = sql.KullaniciOku();
            for (int i = 0; i < Convert.ToInt16(KullaniciBilgileri[3][0]); i++)
            {
                if (textBox1.Text.CompareTo(KullaniciBilgileri[0][i]) == 0 && textBox2.Text.CompareTo(KullaniciBilgileri[1][i]) == 0)
                {
                    KullaniciAdi = KullaniciBilgileri[2][i];
                    Personel_id = Convert.ToInt16(KullaniciBilgileri[4][i]);
                    Form2 form2 = new Form2(ServerName);
                    this.Hide();
                    form2.ShowDialog();
                    this.Close();
                    
                }
            }
        }
        string BaglantiOlustur()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = ServerName;
            builder.InitialCatalog = "master";
            builder.IntegratedSecurity = true;
            return builder.ConnectionString;
        }
         string PersonelBaglantiOlustur()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = ServerName;
            builder.InitialCatalog = "personel_izin";
            builder.IntegratedSecurity = true;
            return builder.ConnectionString;
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            Taployubastanolusturma();
            SqlBaglanti sql = new SqlBaglanti(ServerName);
            try
            {
                sql.KullaniciEkle("Vandellas", "24519551", "Mehmet", "1");
            }
            catch
            {
            }
            Yıllıkİzin();
            KullaniciBilgileri = sql.KullaniciOku();
            textBox1.Text = "Vandellas"; ;
            textBox2.Text = "24519551";
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            string personelotomasyonname = "";
            myCon.Open();
            SqlCommand cmd = new SqlCommand("select * from personel_izin.dbo.Ayarlar", myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                personelotomasyonname = dr["PersonelOtomasyonName"].ToString();
            }
            myCon.Close();
            if (personelotomasyonname == "")
            {
                SqlCommand sorgu = new SqlCommand();
                myCon.Open();
                sorgu.Connection = myCon;
                sorgu.CommandText = "insert into personel_izin.dbo.Ayarlar(PersonelOtomasyonName)values(@Adi)";
                sorgu.Parameters.AddWithValue("@Adi", "Ahmet");
                sorgu.ExecuteNonQuery();
                myCon.Close();
            }
        }

        public void Yıllıkİzin()
        {
            SqlConnection myCon = new SqlConnection(PersonelBaglantiOlustur());
            int ToplamGün = 0;
            int personelid = 0;
            int PreYear = 0;
            myCon.Open();
            SqlCommand cmd = new SqlCommand("select * from dbo.PersonelBilgileri", myCon);
            SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    personelid = Convert.ToInt16(dr["Personel_id"]);
                    string[] str = Convert.ToString(dr["YıllıkizinBaslama_Tar"]).Split('.');
                    ToplamGün = Convert.ToInt16(dr["Yıllıkizin_Süresi"]) + 14;
                    PreYear = Convert.ToInt16(dr["Yıllıkizin_Süresi"]);
                    DateTime buyukTarih = new DateTime(Convert.ToInt16(str[2]), Convert.ToInt16(str[1]), Convert.ToInt16(str[0]));
                    TimeSpan fark = DateTime.Now.Date - buyukTarih;
                    if (fark.TotalDays >= 365)
                    {
                        SqlBaglanti sql = new SqlBaglanti(ServerName);
                        sql.Personelgün(ToplamGün, str[0], str[1], str[2],PreYear, personelid);
                    }
                  
                }
                myCon.Close();
               
        }
        public void Taployubastanolusturma()
        {
            SqlConnection baglanti = new SqlConnection(BaglantiOlustur());
            string dosya = @"USE [master]
GO
/****** Object:  Database [personel_izin]    Script Date: 07/09/2013 17:54:36 ******/
CREATE DATABASE [personel_izin] ON  PRIMARY 
( NAME = N'personelizin', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL11.SQLEXPRESS\MSSQL\DATA\personelizin.mdf' , SIZE = 5120KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )
 LOG ON 
( NAME = N'personelizin_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL11.SQLEXPRESS\MSSQL\DATA\personelizin_log.ldf' , SIZE = 1024KB , MAXSIZE = 2048GB , FILEGROWTH = 10%)
GO
ALTER DATABASE [personel_izin] SET COMPATIBILITY_LEVEL = 100
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [personel_izin].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [personel_izin] SET ANSI_NULL_DEFAULT OFF
GO
ALTER DATABASE [personel_izin] SET ANSI_NULLS OFF
GO
ALTER DATABASE [personel_izin] SET ANSI_PADDING OFF
GO
ALTER DATABASE [personel_izin] SET ANSI_WARNINGS OFF
GO
ALTER DATABASE [personel_izin] SET ARITHABORT OFF
GO
ALTER DATABASE [personel_izin] SET AUTO_CLOSE OFF
GO
ALTER DATABASE [personel_izin] SET AUTO_CREATE_STATISTICS ON
GO
ALTER DATABASE [personel_izin] SET AUTO_SHRINK OFF
GO
ALTER DATABASE [personel_izin] SET AUTO_UPDATE_STATISTICS ON
GO
ALTER DATABASE [personel_izin] SET CURSOR_CLOSE_ON_COMMIT OFF
GO
ALTER DATABASE [personel_izin] SET CURSOR_DEFAULT  GLOBAL
GO
ALTER DATABASE [personel_izin] SET CONCAT_NULL_YIELDS_NULL OFF
GO
ALTER DATABASE [personel_izin] SET NUMERIC_ROUNDABORT OFF
GO
ALTER DATABASE [personel_izin] SET QUOTED_IDENTIFIER OFF
GO
ALTER DATABASE [personel_izin] SET RECURSIVE_TRIGGERS OFF
GO
ALTER DATABASE [personel_izin] SET  DISABLE_BROKER
GO
ALTER DATABASE [personel_izin] SET AUTO_UPDATE_STATISTICS_ASYNC OFF
GO
ALTER DATABASE [personel_izin] SET DATE_CORRELATION_OPTIMIZATION OFF
GO
ALTER DATABASE [personel_izin] SET TRUSTWORTHY OFF
GO
ALTER DATABASE [personel_izin] SET ALLOW_SNAPSHOT_ISOLATION OFF
GO
ALTER DATABASE [personel_izin] SET PARAMETERIZATION SIMPLE
GO
ALTER DATABASE [personel_izin] SET READ_COMMITTED_SNAPSHOT OFF
GO
ALTER DATABASE [personel_izin] SET HONOR_BROKER_PRIORITY OFF
GO
ALTER DATABASE [personel_izin] SET  READ_WRITE
GO
ALTER DATABASE [personel_izin] SET RECOVERY FULL
GO
ALTER DATABASE [personel_izin] SET  MULTI_USER
GO
ALTER DATABASE [personel_izin] SET PAGE_VERIFY CHECKSUM
GO
ALTER DATABASE [personel_izin] SET DB_CHAINING OFF
GO
EXEC sys.sp_db_vardecimal_storage_format N'personel_izin', N'ON'
GO
USE [personel_izin]
GO
/****** Object:  Table [dbo].[PersonelBilgileri]    Script Date: 07/09/2013 17:54:38 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PersonelBilgileri](
	[Personel_id] [int] IDENTITY(1,1) NOT NULL,
	[Personel_Tc] [nvarchar](50) NULL,
	[Personel_Name] [nvarchar](50) NULL,
	[Personel_Surname] [nvarchar](50) NULL,
	[Personel_SGK] [nvarchar](50) NULL,
	[Personel_FatherName] [nvarchar](50) NULL,
	[Personel_Birthday] [nvarchar](50) NULL,
	[Personel_Hometown] [nvarchar](50) NULL,
	[Personel_BasTarihi] [nvarchar](50) NULL,
	[Personel_TelNumber] [nvarchar](50) NULL,
	[Personel_Adres] [nvarchar](50) NULL,
	[Personel_Göreviyeri] [nvarchar](50) NULL,
	[Recort_Date] [nvarchar](50) NULL,
	[Update_Date] [nvarchar](50) NULL,
	[Recort_User] [nvarchar](50) NULL,
	[Update_User] [nvarchar](50) NULL,
	[Recort_id] [int] NULL,
	[Update_id] [int] NULL,
	[Sil_id] [int] NULL,
	[YıllıkizinBaslama_Tar] [nvarchar](50) NULL,
	[HesapNo] [nvarchar](50) NULL,
	[Banka] [nvarchar](50) NULL,
	[Mail] [nvarchar](50) NULL,
	[Kullanılan_Süre] [nvarchar](50) NULL,
	[Yıllıkizin_Süresi] [nvarchar](50) NULL,
	[PreYear] [nvarchar](50) NULL,
	[ThisYear] [nvarchar](50) NULL,
	[total] [nvarchar](50) NULL,
	[EvlilikDurumu] [nvarchar](50) NULL,
	[ÇocukSayisi] [nvarchar](50) NULL,
	[ÖzürlülükDurumu] [nvarchar](50) NULL,
	[ÖzürlülükYüzdesi] [nvarchar](50) NULL,
	[EsDurumu] [nvarchar](50) NULL,
 CONSTRAINT [PK_PersonelBilgileri] PRIMARY KEY CLUSTERED 
(
	[Personel_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[KullaniciBilgileri]    Script Date: 07/09/2013 17:54:38 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[KullaniciBilgileri](
	[KullaniciAdi] [nvarchar](50) NOT NULL,
	[Parola] [nvarchar](50) NOT NULL,
	[PersonelName] [nvarchar](50) NULL,
	[Personelid] [nvarchar](50) NULL,
 CONSTRAINT [PK_KullaniciBilgileri] PRIMARY KEY CLUSTERED 
(
	[KullaniciAdi] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[izinBilgileri]    Script Date: 07/09/2013 17:54:38 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[izinBilgileri](
	[izin_id] [int] IDENTITY(1,1) NOT NULL,
	[personel_id] [int] NULL,
	[rapor_id] [int] NULL,
	[rapor_türü] [nvarchar](50) NULL,
	[isim] [nvarchar](50) NULL,
	[soyad] [nvarchar](50) NULL,
	[tc_no] [nvarchar](50) NULL,
	[tel_no] [nvarchar](50) NULL,
	[isegiris_tar] [nvarchar](50) NULL,
	[bas_tar] [nvarchar](50) NULL,
	[rap_sür] [nvarchar](50) NULL,
	[bitis_tar] [nvarchar](50) NULL,
	[adres] [nvarchar](50) NULL,
	[Sil_id] [int] NULL,
	[Baslama_Saati] [nvarchar](50) NULL,
	[Bitis_Saati] [nvarchar](50) NULL,
	[Mazeret] [nvarchar](max) NULL,
	[Kac_Saat] [nvarchar](50) NULL,
 CONSTRAINT [PK_izinBilgileri] PRIMARY KEY CLUSTERED 
(
	[izin_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Ayarlar]    Script Date: 07/09/2013 17:54:38 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Ayarlar](
	[Ayarlar_id] [int] NULL,
	[Logo1] [image] NULL,
	[Logo2] [image] NULL,
	[Logo3] [image] NULL,
	[PersonelOtomasyonName] [nvarchar](50) NULL,
	[PersonelOtomasyonSurname] [nvarchar](50) NULL,
	[LogoSayisi] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO";
            string[] komutlar = Regex.Split(dosya, @"^\s*GO\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase);
            baglanti.Open();
            bool sonuc = true;
            foreach (string komut in komutlar)
            {
                if (komut.Trim() != "")
                {
                    try
                    {
                        new SqlCommand(komut, baglanti).ExecuteNonQuery();
                    }
                    catch
                    {
                        sonuc = false;
                        break;
                    }
                }
            }
            if (sonuc)
                
            baglanti.Close();
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
