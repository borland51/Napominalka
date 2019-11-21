using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace hakaton
{
    public partial class Settings : Form
    {
        string oldDat = null, oldPath;
        Main form;
        public Settings(Main frm)
        {
            InitializeComponent();
            form = frm;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FBD = new FolderBrowserDialog();

            if (FBD.ShowDialog() == DialogResult.OK)
                textBox1.Text = FBD.SelectedPath;  

            FBD.Dispose();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                textBox1.Text = openFileDialog1.FileName;
        }

        private void Settings_Load(object sender, EventArgs e)
        {
            if (TrshConfig.SettFile != null)
                oldPath = textBox1.Text = TrshConfig.SettFile;

            for (int i = 0; i < TrshConfig.SettDays.Count; i++)
                textBox2.Text += (i > 0 ? "," : "") + TrshConfig.SettDays[i].ToString();

            oldDat = textBox2.Text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            bool save = false;

            try 
            {
                if (String.Compare(oldPath, textBox1.Text) != 0)
                {
                    oldPath = TrshConfig.SettFile = textBox1.Text;
                    save = true;
                }

                if (String.Compare(oldDat, textBox2.Text) != 0)
                {
                    if (String.IsNullOrWhiteSpace(textBox2.Text) || textBox2.Text[0] == ',')
                    {
                        MessageBox.Show("Вы не указали дни!");
                        return;
                    }

                    List<int> tempdays = new List<int>();
                    string[] days = textBox2.Text.Split(',');
                    int cnt = days.Count();

                    for (int i = 0; i < cnt; i++)
                        tempdays.Add(Convert.ToInt32(days[i]));

                    oldDat = textBox2.Text;
                    TrshConfig.SettDays = tempdays;
                    save = true;
                }
            } 
            catch
            {
                MessageBox.Show("Вы неправильно указали дни!\nДни нужно указвать через запятую без пробелов\nНапример: 15,45", "Ошибка!");
            }

            if (save)
            {
                if (TrshConfig.CreateConfig(true))
                {
                    MessageBox.Show("Настройки сохранены!");
                    if (Main.isLoaded)
                        form.LoadForm();

                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if(TrshConfig.SettFile == null)
            {
                MessageBox.Show("Укажите путь к файлу или директории!");
                return;
            }

            if (String.Compare(oldPath, textBox1.Text) != 0 || String.Compare(oldDat, textBox2.Text) != 0)
            {
                if(MessageBox.Show("Вы не сохранили изменения\nВыйти без соханения?", "", MessageBoxButtons.YesNo) == DialogResult.No)
                    return;
            }

            Close();
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != 8 && e.KeyChar != ',')
                e.Handled = true;
        }
    }
}
