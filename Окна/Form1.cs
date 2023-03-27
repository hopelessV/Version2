using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Окна
{
    public partial class Form1 : Form
    {
        double sale;
        string product;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
        }
        /// <summary>
        /// Метод рассчёта
        /// </summary>
        private void Calculation()
        {
            try
            {
                double width = Convert.ToInt32(textBox1.Text);
                double height = Convert.ToInt32(textBox2.Text);
                double m2 = (width * height) / 100;
                if (comboBox1.SelectedIndex == 0 || textBox1.Text != "" && textBox2.Text != "")
                {
                    if (radioButton1.Checked)
                    {
                        sale = Price.deaf * m2;
                        pictureBox1.Image = Image.FromFile(@"C:\Users\hopless\Desktop\Окна\Окна\bin\Debug\image\1.jpg");
                        label5.Text = sale.ToString();
                        product = radioButton1.Text;
                    }
                    else if (radioButton2.Checked)
                    {
                        sale = Price.rotary * m2;
                        pictureBox1.Image = Image.FromFile(@"C:\Users\hopless\Desktop\Окна\Окна\bin\Debug\image\2.jpg");
                        label5.Text = sale.ToString();
                        product = radioButton2.Text;
                    }
                    else if (radioButton3.Checked)
                    {
                        sale = Price.folding * m2;
                        pictureBox1.Image = Image.FromFile(@"C:\Users\hopless\Desktop\Окна\Окна\bin\Debug\image\3.jpg");
                        label5.Text = sale.ToString();
                        product = radioButton3.Text;
                    }
                    else if (radioButton4.Checked)
                    {
                        sale = Price.transom * m2;
                        pictureBox1.Image = Image.FromFile(@"C:\Users\hopless\Desktop\Окна\Окна\bin\Debug\image\4.jpg");
                        label5.Text = sale.ToString();
                        product = radioButton4.Text;
                    }
                    else if (radioButton5.Checked)
                    {
                        sale = Price.sliding * m2;
                        pictureBox1.Image = Image.FromFile(@"C:\Users\hopless\Desktop\Окна\Окна\bin\Debug\image\5.jpg");
                        label5.Text = sale.ToString();
                        product = radioButton5.Text;
                    }
                    else
                    {
                        MessageBox.Show("Тип окна не был выбран", "Сообщение");
                    }
                }
                else
                {
                    MessageBox.Show("Укажите необходимые размеры окна", "Сообщение");
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        /// <summary>
        /// Метод вывода изображений
        /// </summary>
        private void SelectImage()
        {
            try
            {
                if (radioButton1.Checked)
                {
                    pictureBox1.Image = Image.FromFile(@"C:\Users\hopless\Desktop\Окна\Окна\bin\Debug\image\1.jpg");
                }
                else if (radioButton2.Checked)
                {
                    pictureBox1.Image = Image.FromFile(@"C:\Users\hopless\Desktop\Окна\Окна\bin\Debug\image\2.jpg");
                }
                else if (radioButton3.Checked)
                {
                    pictureBox1.Image = Image.FromFile(@"C:\Users\hopless\Desktop\Окна\Окна\bin\Debug\image\3.jpg");
                }
                else if (radioButton4.Checked)
                {
                    pictureBox1.Image = Image.FromFile(@"C:\Users\hopless\Desktop\Окна\Окна\bin\Debug\image\4.jpg");
                }
                else if (radioButton5.Checked)
                {
                    pictureBox1.Image = Image.FromFile(@"C:\Users\hopless\Desktop\Окна\Окна\bin\Debug\image\5.jpg");
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Calculation();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SelectImage();
        }
        //По нажатию на кнопку "Сформировать отчёт"
        private void button3_Click(object sender, EventArgs e)
        {
          Calculation();

          Random rand = new Random();
          string randId = "";

          for (int i = 0; i < 5; i++)
              randId += rand.Next(10).ToString();

          var helper = new WordHelper("report.docx");
          var items = new Dictionary<string, string>
          {
              { "{Уникальный_номер}", randId },
              { "{Дата}", DateTime.Now.ToString("dd.MM.yy") },
              { "{Товар}", product },
              { "{итог}", sale.ToString() }
          };

              helper.Process(items);
          }
        
    }
}
