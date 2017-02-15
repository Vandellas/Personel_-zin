using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Personel_İzin
{
    class SqlBaglanti
    {
        string ServerName;
        public SqlBaglanti(string ServerName)
        {
            this.ServerName = ServerName;
        }
       public  string BaglantiOlustur()
        {
            string kod = ServerName;
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = ServerName;
            builder.InitialCatalog = "personel_izin";
            builder.IntegratedSecurity = true;
            return builder.ConnectionString;
        }
       
        SqlCommand sorgu = new SqlCommand();
        public void KullaniciEkle(string Kullaniciisim,string password,string PersonelName,string Personelid)
        {
            
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            SqlCommand sorgu = new SqlCommand();
            myCon.Open();
            sorgu.Connection = myCon;
            sorgu.CommandText = "insert into dbo.KullaniciBilgileri(KullaniciAdi,Parola,PersonelName,Personelid)values(@Adi,@Parola,@PersonelName,@Personelid)";
            sorgu.Parameters.AddWithValue("@Adi",Kullaniciisim);
            sorgu.Parameters.AddWithValue("@Parola", password);
            sorgu.Parameters.AddWithValue("@PersonelName", PersonelName);
            sorgu.Parameters.AddWithValue("@Personelid", Personelid);
            sorgu.ExecuteNonQuery();
            myCon.Close();
        }
        public void PersonelEkle(string Personel_Tc, string Personel_Name, string Personel_Surname, string Personel_SGK, string Personel_FatherName, string Personel_Birthday, string Personel_HomeTown, string Personel_BasTarihi, string Personel_TelNumber,string GörevYeri,string Record_Date,String Record_User,string Adres,int Recort_id,string Banka,string Mail,string hesapNo,string EvlilikDurumu,string ÇocukSayisi,string ÖzürlülükDurumu,string ÖzürlülükYüzdesi,string EsDurumu)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            sorgu.Connection = myCon;
            sorgu.CommandText = "insert into PersonelBilgileri(Personel_Tc, Personel_Name, Personel_Surname,Personel_SGK,Personel_FatherName,Personel_Birthday,Personel_Hometown,Personel_BasTarihi,Personel_TelNumber,Personel_Göreviyeri,Recort_Date,Recort_User,Personel_Adres,Recort_id,Sil_id,Yıllıkizin_Süresi,Kullanılan_Süre,YıllıkizinBaslama_Tar,HesapNo,Banka,Mail,total,PreYear,ThisYear,EvlilikDurumu,ÇocukSayisi,ÖzürlülükDurumu,ÖzürlülükYüzdesi,EsDurumu)values( @Tc, @Name, @Surname,@SGK,@FatherName,@Birthday,@Hometown,@BasTarihi,@TelNumber,@GörevYeri,@RecordDate,@RecordUser,@Adres,@Recort_id,@Sil_id,@Yıllıkizin_Süresi,@Kullanılan_Süre,@YıllıkizinBaslama_Tar,@HesapNo,@Banka,@Mail,@total,@PreYear,@ThisYear,@evlilik,@cocuk,@özürlü,@yüzdesi,@esdurumu)";
            sorgu.Parameters.AddWithValue("@Tc", Personel_Tc);
            sorgu.Parameters.AddWithValue("@Name",Personel_Name);
            sorgu.Parameters.AddWithValue("@Surname", Personel_Surname);
            sorgu.Parameters.AddWithValue("@SGK", Personel_SGK);
            sorgu.Parameters.AddWithValue("@FatherName", Personel_FatherName);
            sorgu.Parameters.AddWithValue("@Birthday",Personel_Birthday);
            sorgu.Parameters.AddWithValue("@Hometown", Personel_HomeTown);
            sorgu.Parameters.AddWithValue("@BasTarihi", Personel_BasTarihi);
            sorgu.Parameters.AddWithValue("@TelNumber", Personel_TelNumber);
            sorgu.Parameters.AddWithValue("@GörevYeri",GörevYeri);
            sorgu.Parameters.AddWithValue("@RecordDate", Record_Date);
            sorgu.Parameters.AddWithValue("@RecordUser", Record_User);
            sorgu.Parameters.AddWithValue("@Adres", Adres);
            sorgu.Parameters.AddWithValue("@Recort_id", Recort_id);
            sorgu.Parameters.AddWithValue("@Sil_id", '0');
            sorgu.Parameters.AddWithValue("@Yıllıkizin_Süresi",0);
            sorgu.Parameters.AddWithValue("@Kullanılan_Süre", 0);
            sorgu.Parameters.AddWithValue("@ThisYear", 0);
            sorgu.Parameters.AddWithValue("@PreYear", 0);
            sorgu.Parameters.AddWithValue("@total", 0);
            sorgu.Parameters.AddWithValue("@YıllıkizinBaslama_Tar", Personel_BasTarihi);
            sorgu.Parameters.AddWithValue("@HesapNo", hesapNo);
            sorgu.Parameters.AddWithValue("@Banka", Banka);
            sorgu.Parameters.AddWithValue("@Mail", Mail);
            sorgu.Parameters.AddWithValue("@evlilik", EvlilikDurumu);
            sorgu.Parameters.AddWithValue("@cocuk",ÇocukSayisi);
            sorgu.Parameters.AddWithValue("@özürlü", ÖzürlülükDurumu);
            sorgu.Parameters.AddWithValue("@yüzdesi", ÖzürlülükYüzdesi);
            sorgu.Parameters.AddWithValue("@esdurumu", EsDurumu);
            sorgu.ExecuteNonQuery();
            myCon.Close();
            
        }
        public string[][] KullaniciOku()
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            int i = 0;
            string[][] str=new string[5][];
            
            myCon.Open();
            SqlCommand cmd = new SqlCommand("select * from dbo.KullaniciBilgileri", myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            
            while (dr.Read())
            {
                i++;
            }
            for (int j = 0; j < str.GetLength(0); j++)
            {
                str[j] = new string[i];
            }
            str[3] = new string[1];
           str[3][0] = i.ToString();
            i = 0;
            dr.Close();
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                str[0][i] = dr["KullaniciAdi"].ToString(); ;
                str[1][i] = dr["Parola"].ToString();
                str[2][i] = dr["PersonelName"].ToString();
                str[4][i] = dr["Personelid"].ToString();
                i++;
            }
            myCon.Close();
            return str;
        }
        public void Personelgün(int toplamgün,string gün,string ay,string yıl,int PreYear,int personelid)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlCommand cmd2 = new SqlCommand("  Update PersonelBilgileri set Yıllıkizin_Süresi=" + toplamgün + ",Kullanılan_Süre=0,PreYear=" + PreYear + ",ThisYear=14,total="+toplamgün+",YıllıkizinBaslama_Tar='" + gün + "." + ay + "." + (int.Parse(yıl) + 1) + "' where Personel_id=" + personelid + " and Sil_id=0", myCon);
            cmd2.ExecuteNonQuery();
            myCon.Close();
        }
        public string[][] PersonelOku()
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            int i=0;
            string[][] str=new string[17][];
            myCon.Open();
            SqlCommand cmd = new SqlCommand("select * from dbo.PersonelBilgileri", myCon);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                i++;
            }
            for (int j = 0; j < str.GetLength(0)-2; j++)
            {
                str[j] = new string[i];
            }
            str[15]=new string[1];
            str[15][0] = i.ToString();
            str[16] = new string[i];
            i = 0;
            dr.Close();
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {

                str[0][i] = dr["Personel_Tc"].ToString();
                str[1][i] = dr["Personel_Name"].ToString();
                str[2][i] = dr["Personel_Surname"].ToString();
                str[3][i] = dr["Personel_SGK"].ToString();
                str[4][i] = dr["Personel_FatherName"].ToString();
                str[5][i] = dr["Personel_Birthday"].ToString();
                str[6][i] = dr["Personel_Hometown"].ToString();
                str[7][i] = dr["Personel_BasTarihi"].ToString();
                str[8][i] = dr["Personel_TelNumber"].ToString();
                str[9][i] = dr["Personel_Adres"].ToString();
                str[10][i] = dr["Personel_Göreviyeri"].ToString();
                str[11][i] = dr["Recort_Date"].ToString();
                str[12][i] = dr["Update_Date"].ToString();
                str[13][i] = dr["Recort_User"].ToString();
                str[14][i] = dr["Update_User"].ToString();
                str[16][i] = dr["Personel_id"].ToString();
                i++;
            }
            myCon.Close();
            return str;
        }
        public void SilKullanici(int Personelid)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlCommand cmd = new SqlCommand(" Delete from KullaniciBilgileri where Personelid=" + "'" + Personelid + "'", myCon);
            cmd.ExecuteNonQuery();
            myCon.Close();
        }
        public void Silizin(string id,string idTürü)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlCommand cmd = new SqlCommand("  Update izinBilgileri set Sil_id=1 where "+idTürü+"=" + id, myCon);
            cmd.ExecuteNonQuery();
            myCon.Close();
        }
        public void SilPersonel(string Personelid)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlCommand cmd = new SqlCommand("  Update PersonelBilgileri set Sil_id=1 where Personel_id="+Personelid , myCon);
            cmd.ExecuteNonQuery();
            myCon.Close();
        }
        public void GüncelleKullanici(string PersonelAdi,string Password,string KullaniciAdi,string Personelid)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            SqlCommand sorgu = new SqlCommand(" Update KullaniciBilgileri set Parola=" + Password + ",KullaniciAdi=" + "'" + KullaniciAdi + "'" + " where Personelid=" + "'" + Personelid + "'", myCon);
            sorgu.ExecuteNonQuery();
            myCon.Close();
        }
        public void İzinEkle(int raporid,int personelid,string raportürü,string isim,string soyad,string tc_no,string tel_no,string isegiris_tar,string bas_tar,string rap_sür,string bitis_tar,string adres)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            sorgu.Connection = myCon;
            sorgu.CommandText = "insert into izinBilgileri(rapor_id,personel_id,rapor_türü,isim, soyad,tc_no,tel_no,isegiris_tar,bas_tar,rap_sür,bitis_tar,adres,Sil_id)values(@raporid,@personelid,@raportür,@isim,@soyad,@tc,@tel,@isegirtar,@bastar,@rapsür,@bitistar,@adres,@Sil_id)";
            sorgu.Parameters.AddWithValue("@personelid", personelid);
            sorgu.Parameters.AddWithValue("@raporid",raporid);
            sorgu.Parameters.AddWithValue("@raportür", raportürü);
            sorgu.Parameters.AddWithValue("@isim", isim);
            sorgu.Parameters.AddWithValue("@soyad", soyad);
            sorgu.Parameters.AddWithValue("@tc", tc_no);
            sorgu.Parameters.AddWithValue("@tel",tel_no);
            sorgu.Parameters.AddWithValue("@isegirtar", isegiris_tar);
            sorgu.Parameters.AddWithValue("@bastar", bas_tar);
            sorgu.Parameters.AddWithValue("@rapsür", rap_sür);
            sorgu.Parameters.AddWithValue("@bitistar",bitis_tar);
            sorgu.Parameters.AddWithValue("@adres", adres);
            sorgu.Parameters.AddWithValue("@Sil_id", 0);
            sorgu.ExecuteNonQuery();
            myCon.Close();
        }
        public void GünlükizinEkle(int raporid, int personelid, string raportürü,string isim, string soyad, string tc_no, string bas_tar, string bas_saati, string bit_saati, string kac_saati, string mazeret)
        {
            SqlConnection myCon = new SqlConnection(BaglantiOlustur());
            myCon.Open();
            sorgu.Connection = myCon;
            sorgu.CommandText = "insert into izinBilgileri(rapor_id,personel_id,rapor_türü,isim,soyad,tc_no,bas_tar,Baslama_Saati,Bitis_Saati,Kac_Saat,Mazeret,Sil_id)values(@raporid,@personelid,@raportür,@isim,@soyad,@tc,@bastar,@bassaati,@bitsaati,@kaçsaat,@mazeret,@Sil_id)";
            sorgu.Parameters.AddWithValue("@personelid", personelid);
            sorgu.Parameters.AddWithValue("@raporid", raporid);
            sorgu.Parameters.AddWithValue("@raportür", raportürü);
            sorgu.Parameters.AddWithValue("@isim", isim);
            sorgu.Parameters.AddWithValue("@soyad", soyad);
            sorgu.Parameters.AddWithValue("@tc", tc_no);
            sorgu.Parameters.AddWithValue("@bastar", bas_tar);
            sorgu.Parameters.AddWithValue("@Sil_id", 0);
            sorgu.Parameters.AddWithValue("@bassaati", bas_saati);  
            sorgu.Parameters.AddWithValue("@bitsaati", bit_saati); 
            sorgu.Parameters.AddWithValue("@kaçsaat", kac_saati);
            sorgu.Parameters.AddWithValue("@mazeret", mazeret);
            sorgu.ExecuteNonQuery();
            myCon.Close();
        }
    }
}
