using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;


namespace KargoTakip
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        int kontrol = 0;
       public static string girismail;
        string girisifre;
        private static void button3_ClickExtracted()
        {
          
        }
        private void button5_Click(object sender, EventArgs e)
        {
            panel1.Show();
            panel2.Hide();
            panel3.Hide();
            panel4.Hide();


        }
        public void kaydet()
        {
            string baglantiyolu = "provider=microsoft.ace.oledb.12.0;data source=" + Application.StartupPath + "\\kargotakip.accdb";
            OleDbConnection baglanti = new OleDbConnection(baglantiyolu);
            baglanti.Open();
            string eklemekomutu = "insert into uyeler (mail,adsoyad,tcno,dogum_yili,sifre) values ('" + textBox6.Text + "','" + textBox8.Text + "','" + textBox9.Text + "','" + textBox10.Text + "','" + textBox11.Text + "')";
            OleDbCommand komut = new OleDbCommand(eklemekomutu, baglanti);
            komut.ExecuteNonQuery();
            baglanti.Close();
            MessageBox.Show("Üyeliğiniz Başarılı !");

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (textBox6.Text.Contains("@") && textBox6.Text.Contains(".com"))
            {

                if (textBox6.Text == textBox7.Text && textBox11.Text == textBox12.Text)
                {
                    kaydet();
                }
                else
                {
                    MessageBox.Show("lütfen şifre ve mail tekrarı alanlarını birbiriyle eşit yapınız");
                }
            }
            else {
                MessageBox.Show("lütfen mail adresini doğru giriniz");
            }



        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                textBox11.PasswordChar = '\0'; textBox12.PasswordChar = '\0';
            }
            else
            {
                textBox11.PasswordChar = '*'; textBox12.PasswordChar = '*';
            }

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            panel4.Hide();
            panel3.Hide();
            panel2.Visible = true;
            panel1.Visible = false;

            try
            {

                OleDbConnection con;
                con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=kargotakip.accdb");
                con.Open();
                DataTable dt = new DataTable();
                OleDbDataAdapter ad = new OleDbDataAdapter("SELECT * FROM gonderi_takip WHERE takip_no ='" + textBox1.Text + "'", con);
                ad.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Visible = false;
                con.Close();


            }
            catch (Exception)
            {

                kontrol += 1;
                if (kontrol == 4)
                {
                    Application.Exit();
                }
                MessageBox.Show(" Kod Yanlış ! /n 3 kere yanlış kod girerseniz güvenlik sebebiyle program sonlandırılacaktır ");
            }


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Düzce")
            {
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Merkez");
                comboBox2.Items.Add("Beyköy");
                comboBox2.Items.Add("cumayeri");
            }
            if (comboBox1.Text == "İstanbul")
            {

                comboBox2.Items.Clear();
                comboBox2.Items.Add("Fatih");
                comboBox2.Items.Add("Taksim");
                comboBox2.Items.Add("Bebek");
            }
            if (comboBox1.Text == "Ankara")
            {

                comboBox2.Items.Clear();
                comboBox2.Items.Add("Yeni Mahalle");
                comboBox2.Items.Add("Mamak");
                comboBox2.Items.Add("Etimesgut");
            }

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text == "Düzce")
            {
                comboBox4.Items.Clear();
                comboBox4.Items.Add("Merkez");
                comboBox4.Items.Add("Beyköy");
                comboBox4.Items.Add("cumayeri");
            }
            if (comboBox3.Text == "İstanbul")
            {

                comboBox4.Items.Clear();
                comboBox4.Items.Add("Fatih");
                comboBox4.Items.Add("Taksim");
                comboBox4.Items.Add("Bebek");
            }
            if (comboBox3.Text == "Ankara")
            {

                comboBox4.Items.Clear();
                comboBox4.Items.Add("Yeni Mahalle");
                comboBox4.Items.Add("Mamak");
                comboBox4.Items.Add("Etimesgut");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Düzce" && comboBox3.Text == "Ankara" || comboBox3.Text == "Düzce" && comboBox1.Text == "Ankara")
            {
                MessageBox.Show("Ücret 5 TL Teslimat Süreniz 48 Saattir");
                comboBox2.Items.Clear();
                comboBox4.Items.Clear();
            }

            if (comboBox1.Text == "Düzce" && comboBox3.Text == "İstanbul" || comboBox3.Text == "İstanbul" && comboBox1.Text == "Düzce")
            {
                MessageBox.Show("Ücret 10 TL Teslimat Süreniz 62 Saattir");
                comboBox2.Items.Clear();
                comboBox4.Items.Clear();
            }
            if (comboBox1.Text == "Düzce" && comboBox3.Text == "Düzce")
            {
                if (comboBox2.Text != comboBox4.Text)
                {
                    MessageBox.Show("Ücret 3 TL Teslimat Süreniz 6 Saattir");
                    comboBox2.Items.Clear();
                    comboBox4.Items.Clear();
                }
                else
                {
                    MessageBox.Show("Aynı ildeki ilçeler arası kargo hizmeti yokdur");

                }

            }
            if (comboBox1.Text == "Ankara" && comboBox3.Text == "İstanbul" || comboBox3.Text == "İstanbul" && comboBox1.Text == "Ankara")
            {
                MessageBox.Show("Ücret 5 TL Teslimat Süreniz 48 Saattir");
                comboBox2.Items.Clear();
                comboBox4.Items.Clear();
            }
            if (comboBox1.Text == "Ankara" && comboBox3.Text == "Ankara")
            {
                if (comboBox2.Text != comboBox4.Text)
                {
                    MessageBox.Show("Ücret 3 TL Teslimat Süreniz 6 Saattir");
                    comboBox2.Items.Clear();
                    comboBox4.Items.Clear();
                }
                else
                {
                    MessageBox.Show("Aynı ildeki ilçeler arası kargo hizmeti yokdur");

                }

            }
            if (comboBox1.Text == "istanbul" && comboBox3.Text == "İstanbul")
            {
                if (comboBox2.Text != comboBox4.Text)
                {
                    MessageBox.Show("Ücret 3 TL Teslimat Süreniz 6 Saattir");
                    comboBox2.Items.Clear();
                    comboBox4.Items.Clear();
                }
                else
                {
                    MessageBox.Show("Aynı ildeki ilçeler arası kargo hizmeti yokdur");

                }

            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            panel3.Show();
            panel1.Hide();
            panel2.Hide();
            panel4.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            panel3.Show();
            panel1.Hide();
            panel2.Hide();
            panel4.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form1 frm1 = new Form1();
          
           
       string kontrolmail = textBox2.Text;
            string kontrolsifre = textBox3.Text;
            OleDbConnection con;
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=kargotakip.accdb");
            con.Open();
            OleDbCommand okuma = new OleDbCommand();
            okuma.Connection = con;
            okuma.CommandText = "SELECT * FROM uyeler WHERE mail= '" + textBox2.Text + "'";
            OleDbDataReader reader = okuma.ExecuteReader();
            while (reader.Read())
            {
                girismail = reader["mail"].ToString();
                girisifre = reader["sifre"].ToString();
                
            }
            con.Close();
         //   MessageBox.Show(girismail,girisifre);
            if (girismail ==kontrolmail&& girisifre==kontrolsifre)
            {
                this.Hide();
                musteri mstr = new musteri();
                mstr.Show();
            }
            else
            {
                label18.Text ="lütfen üye olunuz";
            }
            
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

            Form1 frm1 = new Form1();


            string kontrolname = textBox5.Text;
            string kontrolsifre = textBox4.Text;
         
            OleDbConnection con;
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=kargotakip.accdb");
            con.Open();
            OleDbCommand okuma = new OleDbCommand();
            okuma.Connection = con;
            okuma.CommandText = "SELECT * FROM admin WHERE username= '" + textBox5.Text + "'";
            OleDbDataReader reader = okuma.ExecuteReader();
            while (reader.Read())
            {
                girismail = reader["username"].ToString();
                girisifre = reader["password"].ToString();

            }
            con.Close();
           
            if (girismail == kontrolname && girisifre == kontrolsifre)
            {
                this.Hide();
                admin adm = new admin();
                adm.Show();
            }
            else
            {
                label18.Text = "lütfen üye olunuz";
            }
        }
    }
}
