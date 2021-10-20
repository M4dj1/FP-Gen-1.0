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
    public partial class main : Form
    {
        public main()
        {
            InitializeComponent();
            printBtnPnl.Visible = false;
            addBtnPnl.Visible = false;
            listBtnPnl.Visible = false;
            abtBtnPnl.Visible = false;
            addCusBtnPnl.Visible = false;
            hisBtnPnl.Visible = true;
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void printShBtn_Click(object sender, EventArgs e)
        {
            hisBtnPnl.Visible = false;
            printBtnPnl.Visible = true;
            addCusBtnPnl.Visible = false;
            addBtnPnl.Visible = false;
            listBtnPnl.Visible = false;    
            abtBtnPnl.Visible = false;  
            printPnl.Visible=true; 
            addPnl.Visible=false;   
            listPnl.Visible=false;  
            abtPnl.Visible=false;
            addCusPnl.Visible = false;
            hisPnl.Visible = false;
        }

        private void addBtn_Click(object sender, EventArgs e)
        {
            hisBtnPnl.Visible = false;
            printBtnPnl.Visible = false;
            addCusBtnPnl.Visible = false;
            addBtnPnl.Visible = true;
            listBtnPnl.Visible = false;
            abtBtnPnl.Visible = false;
            printPnl.Visible = false;
            addPnl.Visible = true;
            listPnl.Visible = false;
            abtPnl.Visible = false;
            addCusPnl.Visible = false;
            hisPnl.Visible = false;
        }

        private void listBtn_Click(object sender, EventArgs e)
        {
            hisBtnPnl.Visible = false;
            printBtnPnl.Visible = false;
            addCusBtnPnl.Visible = false;
            addBtnPnl.Visible = false;
            listBtnPnl.Visible = true;
            abtBtnPnl.Visible = false;
            printPnl.Visible = false;
            addPnl.Visible = false;
            listPnl.Visible = true;
            abtPnl.Visible = false;
            addCusPnl.Visible = false;
            hisPnl.Visible = false;
        }

        private void abtBtn_Click(object sender, EventArgs e)
        {
            hisBtnPnl.Visible = false;
            printBtnPnl.Visible = false;
            addCusBtnPnl.Visible = false;
            addBtnPnl.Visible = false;
            listBtnPnl.Visible = false;
            abtBtnPnl.Visible = true;
            printPnl.Visible = false;
            addPnl.Visible = false;
            listPnl.Visible = false;
            abtPnl.Visible = true;
            addCusPnl.Visible = false;
            hisPnl.Visible = false;
        }

        private void extBtn_Click(object sender, EventArgs e)
        {
            this.Close();   
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            hisBtnPnl.Visible = false;
            printBtnPnl.Visible = false;
            addCusBtnPnl.Visible = false;
            addBtnPnl.Visible = true;
            listBtnPnl.Visible = false;
            abtBtnPnl.Visible = false;
            printPnl.Visible = false;
            addPnl.Visible = true;
            listPnl.Visible = false;
            abtPnl.Visible = false;
            addCusPnl.Visible = false;
            hisPnl.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            hisBtnPnl.Visible = false;
            printBtnPnl.Visible = true;
            addCusBtnPnl.Visible = false;
            addBtnPnl.Visible = false;
            listBtnPnl.Visible = false;
            abtBtnPnl.Visible = false;
            printPnl.Visible = true;
            addPnl.Visible = false;
            listPnl.Visible = false;
            abtPnl.Visible = false;
            addCusPnl.Visible = false;
            hisPnl.Visible = false;
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            printBtnPnl.Visible = false;
            hisBtnPnl.Visible = false;
            addCusBtnPnl.Visible = true;
            addBtnPnl.Visible = false;
            listBtnPnl.Visible = false;
            abtBtnPnl.Visible = false;
            printPnl.Visible = false;
            addPnl.Visible = false;
            listPnl.Visible = false;
            abtPnl.Visible = false;
            addCusPnl.Visible = true;
            hisPnl.Visible = false;
        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void hisBtn_Click(object sender, EventArgs e)
        {
            printBtnPnl.Visible = false;
            hisBtnPnl.Visible = true;
            addCusBtnPnl.Visible = false;
            addBtnPnl.Visible = false;
            listBtnPnl.Visible = false;
            abtBtnPnl.Visible = false;
            printPnl.Visible = false;
            addPnl.Visible = false;
            listPnl.Visible = false;
            abtPnl.Visible = false;
            addCusPnl.Visible = false;
            hisPnl.Visible = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }
    }
}
