# Kargo-Takip
Kargo Takip üzerine yaptığım projem
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.VisualBasic;

namespace KargoTakip
{
    public partial class musteri : Form
    {
        public musteri()
        {
            InitializeComponent();
        }
       
         Form1 FRM1 = new Form1();
        
 
        private void button1_Click(object sender, EventArgs e)
        {
            Random rdm = new Random();
         int takipno=   rdm.Next(100001,999999);
            string mailkayit = Form1.girismail;
        

            string baglantiyolu = "provider=microsoft.ace.oledb.12.0;data source=" + Application.StartupPath + "\\kargotakip.accdb";
            OleDbConnection baglanti = new OleDbConnection(baglantiyolu);
            baglanti.Open();
            string eklemekomutu = "insert into gonderi_takip (takip_no,teslim_tipi,gonderilen_tarih,odeme_tipi,alici_adres,gonderen_mail) values ('" + takipno.ToString() + "','" + comboBox1.Text + "','" + textBox1.Text + "','" + comboBox3.Text + "','" + richTextBox1.Text + "','" + mailkayit + "')";
            OleDbCommand komut = new OleDbCommand(eklemekomutu, baglanti);
            komut.ExecuteNonQuery();
            baglanti.Close();
            MessageBox.Show("Gönderiniz Kaydedildi Kargonuz Ekiplerimizce Kapınızdan Alınıcak");
            label16.Text = "Son Gönderi Takip No: "+takipno.ToString();

        }
        public void musteribilgileri()
        {
            

       
            OleDbConnection con;
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=kargotakip.accdb");
            con.Open();
            OleDbCommand okuma = new OleDbCommand();
            okuma.Connection = con;
            okuma.CommandText = "SELECT * FROM gonderi_takip WHERE gonderen_mail= '" + Form1.girismail + "'";
            OleDbDataReader reader = okuma.ExecuteReader();
            while (reader.Read())
            {
                comboBox2.Items.Add(reader["takip_no"]);

            }

           con.Close();

           
        }
        public void doldur()
        {
            textBox1.Text = DateTime.Today.ToShortDateString();
            OleDbConnection con1;
            con1 = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=kargotakip.accdb");
            con1.Open();
            OleDbCommand okuma1 = new OleDbCommand();
            okuma1.Connection = con1;
            okuma1.CommandText = "SELECT * FROM uyeler WHERE mail= '" + Form1.girismail + "'";
            OleDbDataReader reader1 = okuma1.ExecuteReader();
            while (reader1.Read())
            {
                textBox2.Text = reader1["adsoyad"].ToString();
                textBox3.Text = reader1["tcno"].ToString();
                textBox4.Text = reader1["dogum_yili"].ToString();
            }
            con1.Close();
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            OleDbConnection con;
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=kargotakip.accdb");
            con.Open();
            OleDbCommand okuma = new OleDbCommand();
            okuma.Connection = con;
            okuma.CommandText = "SELECT * FROM gonderi_takip WHERE takip_no= '" + comboBox2.Text + "'";
            OleDbDataReader reader = okuma.ExecuteReader();
            while (reader.Read())
            {
               label7.Text =reader["durum"].ToString();

            }

            con.Close();



                      string durum = "";
                        durum = label7.Text;
                        if (durum=="Kargo Ulastı")
                        {
                
                pictureBox1.Image = Image.FromFile(Application.StartupPath + "\\no-icon.png");
            }
                        else if (durum=="Kargo Bayide")
            {
                pictureBox1.Image = Image.FromFile(Application.StartupPath + "\\bayide.png");



            }
                        else if (durum=="Kargo Taşınıyor")
                        {

                pictureBox1.Image = Image.FromFile(Application.StartupPath + "\\tasiniyor.png");
            }
                        else
                        {
                pictureBox1.Image = Image.FromFile(Application.StartupPath + "\\no-icon.png");


            } 
                        
        }

        private void musteri_Load(object sender, EventArgs e)
        {
            musteribilgileri();
            doldur();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void çıkışToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void şifreDeğiştirToolStripMenuItem_Click(object sender, EventArgs e)
        {
           string YeniSifre = Interaction.InputBox("Lütfen Yeni Şifre oluşturunuz", "Yeni Şifre Girişi", "123");
           string KontrolSifre = Interaction.InputBox("Lütfen Güvenlik Sebebiyle Yeni Şifrenizi Tekrar Giriniz", "Yeni Şifre Girişi", "123");
            if (YeniSifre==KontrolSifre&& YeniSifre!="123" &&YeniSifre!="")
            {

                string baglantiyolu = "provider=microsoft.ace.oledb.12.0;data source=" + Application.StartupPath + "\\kargotakip.accdb";
                OleDbConnection baglanti = new OleDbConnection(baglantiyolu);
                baglanti.Open();

                OleDbCommand kaydet = new OleDbCommand("update uyeler set sifre='" + YeniSifre  + "' where mail='" + Form1.girismail + "'", baglanti);
                
                kaydet.ExecuteNonQuery();
                baglanti.Close();
                MessageBox.Show("Şifre Değiştirme Başarılı !");
                

            }
            else if (YeniSifre=="123")
            {
                MessageBox.Show("123 dışında bir şifre giriniz !");

            }
            else
            {
                MessageBox.Show("Şifre değiştirme başarısız !");
            }
        }

        private void bilgilerimiGüncelleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;
         Güncelle.Visible=true;

           






        }

        private void Güncelle_Click(object sender, EventArgs e)
        {
            string baglantiyolu = "provider=microsoft.ace.oledb.12.0;data source=" + Application.StartupPath + "\\kargotakip.accdb";
            OleDbConnection baglanti = new OleDbConnection(baglantiyolu);
            baglanti.Open();

            OleDbCommand kaydet = new OleDbCommand("update uyeler set adsoyad='" + textBox2.Text + "', tcno='" + textBox3.Text + "' , dogum_yili='" + textBox4.Text + "' where mail='" + Form1.girismail + "'", baglanti);

            kaydet.ExecuteNonQuery();
            baglanti.Close();
            MessageBox.Show("Bilgileriniz Güncellendi !");
            doldur();
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
            Güncelle.Visible = false;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string baglantiyolu = "provider=microsoft.ace.oledb.12.0;data source=" + Application.StartupPath + "\\kargotakip.accdb";
            OleDbConnection baglanti = new OleDbConnection(baglantiyolu);
            baglanti.Open();
            string eklemekomutu = "insert into destek (konu,mesaj,gonderen_mail) values ('" + textBox5.Text + "','" + richTextBox2.Text + "','" + Form1.girismail + "')";
            OleDbCommand komut = new OleDbCommand(eklemekomutu, baglanti);
            komut.ExecuteNonQuery();
            baglanti.Close();
            MessageBox.Show("Destek talebiniz Gönderildi ! 24 Saat içerisinde mailinize Cevap gelicektir.");
            groupBox3.Visible = false;
        }

        private void çıkışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox3.Visible = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult tus;
            tus = colorDialog1.ShowDialog();
            if (tus == DialogResult.OK)
            {
                this.BackColor = colorDialog1.Color;
            }

        }

    
    }
}
