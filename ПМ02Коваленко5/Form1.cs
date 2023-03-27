using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace ПМ02Коваленко5
{
    public partial class Form1 : Form
    {
        //создаем переменные для удобной работы
        double dopSum = 0;
        double sumKvSm = 0;
        double itogSum = 0;
        public Form1()
        {
            InitializeComponent();
            pictureBox1.Image = Properties.Resources.logo;
            comboBox1.SelectedIndex = 0;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedIndex)
            {
                case 0:
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton3.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                    radioButton1.Checked = true;
                    sumKvSm = 50;
                    break;

                case 1:
                    pictureBox2.Image = Properties.Resources.балкон;
                    radioButton1.Enabled = false;
                    radioButton2.Enabled = false;
                    radioButton3.Enabled = false;
                    radioButton4.Enabled = false;
                    radioButton5.Enabled = false;
                    sumKvSm = 180.50;
                    dopSum = 0;
                    break;

                case 2:
                    pictureBox2.Image = Properties.Resources.дверь;
                    radioButton1.Enabled = false;
                    radioButton2.Enabled = false;
                    radioButton3.Enabled = false;
                    radioButton4.Enabled = false;
                    radioButton5.Enabled = false;
                    sumKvSm = 200;
                    dopSum = 0;
                    break;

            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox2.Image = Properties.Resources._1;
            if (radioButton1.Checked)
            {
                radioButton2.Checked = false;
                radioButton3.Checked = false;
                radioButton4.Checked = false;
                radioButton5.Checked = false;
            }
            dopSum = 1000.00;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox2.Image = Properties.Resources._2;
            if (radioButton2.Checked)
            {
                radioButton1.Checked = false;
                radioButton3.Checked = false;
                radioButton4.Checked = false;
                radioButton5.Checked = false;
            }
            dopSum = 3400.50;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox2.Image = Properties.Resources._3;
            if (radioButton3.Checked)
            {
                radioButton2.Checked = false;
                radioButton1.Checked = false;
                radioButton4.Checked = false;
                radioButton5.Checked = false;
            }
            dopSum = 2560;
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox2.Image = Properties.Resources._4;
            if (radioButton4.Checked)
            {
                radioButton2.Checked = false;
                radioButton3.Checked = false;
                radioButton1.Checked = false;
                radioButton5.Checked = false;
            }
            dopSum = 7900.90;
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox2.Image = Properties.Resources._5;
            if (radioButton5.Checked)
            {
                radioButton2.Checked = false;
                radioButton3.Checked = false;
                radioButton4.Checked = false;
                radioButton1.Checked = false;
            }
            dopSum = 6210.50;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //рассчёт итоговой суммы
            itogSum = dopSum + sumKvSm/ 10000 * (Convert.ToDouble(numericUpDown1.Value) * Convert.ToDouble(numericUpDown2.Value));
            label4.Text = "Итоговая сумма: "+itogSum.ToString();
            button2.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                //в файле хранится номер для заказа
                int code = Convert.ToInt32(File.ReadAllText(@"код.txt"));
                string date = DateTime.Now.ToShortDateString();
                File.Copy(@"чек.docx", @"чеки\" + code + "_" + date + "_" + itogSum + ".docx");
                //копируем образец чека и работаем с ним
                Word.Document doc = null;
                Word.Application app = new Word.Application();
                string sourse = @"C:\Users\Василиса\Documents\ПМ02Коваленко5\ПМ02Коваленко5\bin\Debug\чеки\" + code + "_" + date + "_" + itogSum + ".docx";
                doc = app.Documents.Open(sourse);
                doc.Activate();
                Word.Bookmarks book = doc.Bookmarks;
                Word.Range range;
                int i = 0;
                string[] data = new string[] { code.ToString(), date.ToString(), comboBox1.Text, itogSum.ToString() };
                foreach (Word.Bookmark b in book)
                {
                    range = b.Range;
                    range.Text = data[i];
                    i++;
                }
                doc.Close();
                doc = null;
                code++;
                //перезаписываем код в файл
                File.WriteAllText(@"код.txt", code.ToString());
                MessageBox.Show("Успешно!");
            }
            catch
            {
                MessageBox.Show("Ошибка, попробуйте ещё раз");
            }
           
        }
    }
}
