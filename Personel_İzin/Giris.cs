using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Personel_İzin
{
    public partial class Giris : Form
    {
        public Giris()
        {
            InitializeComponent();
        }
        private void Kaydet_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory("C:\\Personelizin\\ServerName");
            StreamWriter yaz = new StreamWriter("C:\\Personelizin\\ServerName\\ServerName.txt");
            yaz.WriteLine(textBox1.Text);
            yaz.Close();
            Form1 frm1 = new Form1(textBox1.Text);
            this.Hide();
            frm1.ShowDialog();
            this.Close();
        }

        private void Giris_Load(object sender, EventArgs e)
        {
            string dosyayolu = "C:\\Personelizin\\ServerName\\ServerName.txt";
            if (File.Exists(dosyayolu))
            {
                StreamReader oku = File.OpenText("C:\\Personelizin\\ServerName\\ServerName.txt");
                string metin = oku.ReadLine();
                if (metin!= "")
                {
                    Form1 frm1 = new Form1(metin);
                    this.Hide();
                    frm1.ShowDialog();
                    this.Close();
                }
                
            }
            else
                textBox1.Text = Environment.MachineName + @"\SQLEXPRESS";

        }
    }
}
