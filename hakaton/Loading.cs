using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace hakaton
{
    public partial class Loading : Form
    {
        int count = 0;
        public Loading()
        {
            InitializeComponent();
        }

        public void SetProgress(double percent)
        {
            label1.Text = "Загрузка";
            count = ++count > 3 ? 0 : count;
            for (int i = 0; i < count; i++)
                label1.Text += ".";
            
            label1.Text += " " + percent.ToString() + "%";
        }
    }
}
