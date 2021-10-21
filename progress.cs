using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FP_Gen_1._0
{
    public partial class progress : Form
    {
        public progress()
        {
            InitializeComponent();
        }

        private void timer2_Tick_1(object sender, EventArgs e)
        {
            if (Opacity == 1)
            {
                timer2.Stop();
            }
            Opacity += .15;
        }
        private void timer1_Tick_1(object sender, EventArgs e)
        {
            panel2.Width += 3;
            if (panel2.Width >= 494)
            {
                timer1.Stop();
                main m = new main();
                m.Show();
                this.Hide();
            }
            
        }

        private void progress_Load(object sender, EventArgs e)
        {
            label1.Parent = pictureBox1;
            label2.Parent = pictureBox1;
            label1.BackColor = Color.Transparent; label2.BackColor = Color.Transparent; 
        }

    }
}
