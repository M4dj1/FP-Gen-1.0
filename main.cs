using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Reflection;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Spire.Doc;
using System.Drawing.Printing;
using System.Data.SqlClient;

namespace FP_Gen_1._0
{
    public partial class main : Form
    {
        SqlConnection connection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\FP-Gen\Database1.mdf;Integrated Security=True");
        public main()
        {
            InitializeComponent();
            displayListView();
            displayListView1();
            printBtnPnl.Visible = false;
            addBtnPnl.Visible = false;
            listBtnPnl.Visible = false;
            abtBtnPnl.Visible = false;
            addCusBtnPnl.Visible = false;
            hisBtnPnl.Visible = true;
        }

        private void printShBtn_Click(object sender, EventArgs e)
        {
            displayPrintCusCombo();
            hisBtnPnl.Visible = false;
            printBtnPnl.Visible = true;
            addCusBtnPnl.Visible = false;
            addBtnPnl.Visible = false;
            listBtnPnl.Visible = false;    
            abtBtnPnl.Visible = false;  
            printPnl.Visible=true;
            addPnl.Visible = false;   
            listPnl.Visible=false;  
            abtPnl.Visible=false;
            addCusPnl.Visible = false;
            hisPnl.Visible = false;
            addSF.Visible = false;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = true;
            SFPnlBtn.Visible = false;
            PFPnlBtn.Visible = true;
        }

        private void prntSFBtn_Click(object sender, EventArgs e)
        {
            SF.Visible = true;
            PF.Visible = false;
            PFPnlBtn.Visible = false;
            SFPnlBtn.Visible = true;
        }

        private void prntPFBtn_Click(object sender, EventArgs e)
        {
            SF.Visible = false;
            PF.Visible = true;
            SFPnlBtn.Visible = false;
            PFPnlBtn.Visible = true;
        }

        private void addBtn_Click(object sender, EventArgs e)
        {
            displayAddCusCombo();
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
            addSF.Visible = false;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
        }

        private void sfBtn_Click(object sender, EventArgs e)
        {
            hisBtnPnl.Visible = false;
            printBtnPnl.Visible = false;
            addCusBtnPnl.Visible = false;
            addBtnPnl.Visible = true;
            listBtnPnl.Visible = false;
            abtBtnPnl.Visible = false;
            printPnl.Visible = false;
            addPnl.Visible = false;
            listPnl.Visible = false;
            abtPnl.Visible = false;
            addCusPnl.Visible = false;
            hisPnl.Visible = false;
            addSF.Visible = true;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
        }

        private void fpBtn_Click(object sender, EventArgs e)
        {
            hisBtnPnl.Visible = false;
            printBtnPnl.Visible = false;
            addCusBtnPnl.Visible = false;
            addBtnPnl.Visible = true;
            listBtnPnl.Visible = false;
            abtBtnPnl.Visible = false;
            printPnl.Visible = false;
            addPnl.Visible = false;
            listPnl.Visible = false;
            abtPnl.Visible = false;
            addCusPnl.Visible = false;
            hisPnl.Visible = false;
            addPF.Visible = true;
            addSF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
        }

        private void listBtn_Click(object sender, EventArgs e)
        {
            displayListView();
            displayListView1();
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
            addSF.Visible = false;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
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
            addSF.Visible = false;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
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
            addSF.Visible = false;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
        }

        private void his()
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
            addSF.Visible = false;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
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
            addSF.Visible = false;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
        }

        private void addItCancelBtn_Click(object sender, EventArgs e)
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
            addSF.Visible = false;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
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
            addSF.Visible = false;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
        }

        private void sfAddCusBtn_Click_1(object sender, EventArgs e)
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
            addSF.Visible = false;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
        }

        private void sfAddItBtn_Click_1(object sender, EventArgs e)
        {
            displayAddCusCombo();
            hisBtnPnl.Visible = false;
            printBtnPnl.Visible = false;
            addCusBtnPnl.Visible = false;
            addBtnPnl.Visible = true;
            listBtnPnl.Visible = false;
            abtBtnPnl.Visible = false;
            printPnl.Visible = false;
            addPnl.Visible = false;
            listPnl.Visible = false;
            abtPnl.Visible = false;
            addCusPnl.Visible = false;
            hisPnl.Visible = false;
            addSF.Visible = true;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
        }

        private void printAddBtn1_Click(object sender, EventArgs e)
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
            addSF.Visible = false;
            addPF.Visible = false;
            SF.Visible = false;
            PF.Visible = false;
        }

        private void printAddBtn2_Click(object sender, EventArgs e)
        {
            displayAddCusCombo();
            hisBtnPnl.Visible = false;
            printBtnPnl.Visible = false;
            addCusBtnPnl.Visible = false;
            addBtnPnl.Visible = true;
            listBtnPnl.Visible = false;
            abtBtnPnl.Visible = false;
            printPnl.Visible = false;
            addPnl.Visible = false;
            listPnl.Visible = false;
            abtPnl.Visible = false;
            addCusPnl.Visible = false;
            hisPnl.Visible = false;
            addSF.Visible = false;
            addPF.Visible = true;
            SF.Visible = false;
            PF.Visible = false;
        }



        private void addCusSaveBtn_Click(object sender, EventArgs e)
        {
            connection.Open();
            SqlCommand cmd = connection.CreateCommand ();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "insert into [customer] (name, address) values ('" + textBox5.Text + "','" + textBox6.Text + "')";
            cmd.ExecuteNonQuery ();
            connection.Close ();

            textBox5.Text = "";
            textBox6.Text = "";
            MessageBox.Show("Data Inserted Successfully !");
            displayPrintCusCombo();
        }



        public void displayPrintCusCombo()
        {
            connection.Open();
            SqlCommand cmd = connection.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select id,name,address from [customer]";
            cmd.ExecuteNonQuery();
            connection.Close();

            DataTable table1 = new DataTable();
            SqlDataAdapter ada = new SqlDataAdapter(cmd);
            ada.Fill(table1);
            DataRow itemrow = table1.NewRow();
            itemrow[1] = "- Select Customer...";
            table1.Rows.InsertAt(itemrow, 0);
            pfCusBox.DisplayMember = "name";
            pfCusBox.ValueMember = "id";
            pfCusBox.DataSource = table1;
            comboBox5.DisplayMember = "name";
            comboBox5.ValueMember = "id";
            comboBox5.DataSource = table1;
            pfItemBox.Enabled = false;
            comboBox4.Enabled = false;
        }

        public void displayAddCusCombo()
        {
            connection.Open();
            SqlCommand cmd = connection.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select id,name,address from [customer]";
            cmd.ExecuteNonQuery();
            connection.Close();

            DataTable table1 = new DataTable();
            SqlDataAdapter ada = new SqlDataAdapter(cmd);
            ada.Fill(table1);
            DataRow itemrow = table1.NewRow();
            itemrow[1] = "- Select Customer...";
            table1.Rows.InsertAt(itemrow, 0);
            comboBox2.DisplayMember = "name";
            comboBox2.ValueMember = "id";
            comboBox2.DataSource = table1;
            comboBox8.DisplayMember = "name";
            comboBox8.ValueMember = "id";
            comboBox8.DataSource = table1;
        }


        private void saveBtn_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != null && textBox2.Text != null
                && comboBox2.SelectedItem != null)
            { connection.Open();
                SqlCommand cmd = connection.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "insert into [item] (itname, dimensions, customerid) values ('" + textBox1.Text + "','" + textBox2.Text + "','" + int.Parse(comboBox2.SelectedIndex.ToString()) + "')";
                cmd.ExecuteNonQuery();
                connection.Close();

                textBox1.Text = "";
                textBox2.Text = "";
                comboBox2.Text = "- Select Customer...";
                MessageBox.Show("Data Inserted Successfully !");
                displayAddCusCombo();
            } else
            {
                MessageBox.Show("Please fill all the required fields !");
            }
        }

        private void SaveBtn2_Click(object sender, EventArgs e)
        {
            if (textBox7.Text != null && textBox4.Text != null &&
                comboBox6.SelectedItem != null && comboBox7.SelectedItem != null
                && comboBox8.SelectedItem != null)
            {
                connection.Open();
                SqlCommand cmd = connection.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "insert into [cardboard] (cname, type, form, dimensions, customerid) values ('" + textBox7.Text + "','" + comboBox6.SelectedItem.ToString() + "','" + comboBox7.SelectedItem.ToString() + "','" + textBox4.Text + "','" + int.Parse(comboBox8.SelectedIndex.ToString()) + "')";
                cmd.ExecuteNonQuery();
                connection.Close();

                textBox1.Text = "";
                textBox2.Text = "";
                comboBox2.Text = "- Select Customer...";
                MessageBox.Show("Data Inserted Successfully !");
                displayAddCusCombo();
            }
            else
            {
                MessageBox.Show("Please fill all the required fields !");
            }
        }


        private void extBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void pfCusBox_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (pfCusBox.SelectedValue.ToString() != null)
            {
                connection.Open();
                SqlCommand cmd = connection.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "select id,itname,dimensions,customerid from [item] where customerid = @cusid";
                cmd.Parameters.AddWithValue("@cusid", pfCusBox.SelectedValue.ToString());
                cmd.ExecuteNonQuery();
                connection.Close();
                DataTable table1 = new DataTable();
                SqlDataAdapter ada = new SqlDataAdapter(cmd);
                ada.Fill(table1);
                DataRow itemrow = table1.NewRow();
                itemrow[1] = "- Select Item...";
                table1.Rows.InsertAt(itemrow, 0);
                pfItemBox.Enabled = true;
                pfItemBox.DisplayMember = "itname";
                pfItemBox.ValueMember = "id";
                pfItemBox.DataSource = table1;

                connection.Open();
                SqlCommand cmd1 = connection.CreateCommand();
                cmd1.CommandType = CommandType.Text;
                cmd1.CommandText = "select address from [customer] where id = @cusid";
                cmd1.Parameters.AddWithValue("@cusid", pfCusBox.SelectedValue.ToString());
                SqlDataReader dr = cmd1.ExecuteReader();
                while (dr.Read())
                {
                    pfAdrTxtBx.Text = dr.GetValue(0).ToString();
                }
                dr.Close();
                connection.Close();
            }
        }

        private void sfCusBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.SelectedValue.ToString() != null)
            {
                connection.Open();
                SqlCommand cmd = connection.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "select id,cname,form,type,dimensions,customerid from [cardboard] where customerid = @cusid";
                cmd.Parameters.AddWithValue("@cusid", comboBox5.SelectedValue.ToString());
                cmd.ExecuteNonQuery();
                connection.Close();
                DataTable table1 = new DataTable();
                SqlDataAdapter ada = new SqlDataAdapter(cmd);
                ada.Fill(table1);
                DataRow itemrow = table1.NewRow();
                itemrow[1] = "- Select Item...";
                table1.Rows.InsertAt(itemrow, 0);
                comboBox4.Enabled = true;
                comboBox4.DisplayMember = "cname";
                comboBox4.ValueMember = "id";
                comboBox4.DataSource = table1;

                connection.Open();
                SqlCommand cmd1 = connection.CreateCommand();
                cmd1.CommandType = CommandType.Text;
                cmd1.CommandText = "select address from [customer] where id = @cusid";
                cmd1.Parameters.AddWithValue("@cusid", comboBox5.SelectedValue.ToString());
                SqlDataReader dr = cmd1.ExecuteReader();
                while (dr.Read())
                {
                    adrTxtBx2.Text = dr.GetValue(0).ToString();
                }
                dr.Close();
                connection.Close();
            }
        }

        public void displayListView()
        {
            connection.Open();
            SqlCommand cmd = connection.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select customer.name, customer.address, item.itname," +
                "item.dimensions from [customer] " +
                "inner join item ON customer.id = item.customerid ORDER BY customer.id";
            SqlDataReader reader = cmd.ExecuteReader();
            listView1.Items.Clear();

            while (reader.Read())
            {
                ListViewItem lv = new ListViewItem(reader.GetString(0));
                lv.SubItems.Add(reader.GetString(1));
                lv.SubItems.Add(reader.GetString(2));
                lv.SubItems.Add(reader.GetString(3));
                listView1.Items.Add(lv);
            }
            reader.Close();
            connection.Close();
        }

        public void displayListView1()
        {
            connection.Open();
            SqlCommand cmd = connection.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select customer.name, customer.address, cardboard.cname," +
                "cardboard.type, cardboard.form, cardboard.dimensions from [customer] " +
                "inner join cardboard ON customer.id = cardboard.customerid ORDER BY customer.id";
            SqlDataReader reader = cmd.ExecuteReader();
            listView3.Items.Clear();

            while (reader.Read())
            {
                ListViewItem lv = new ListViewItem(reader.GetString(0));
                lv.SubItems.Add(reader.GetString(1));
                lv.SubItems.Add(reader.GetString(2));
                lv.SubItems.Add(reader.GetString(3));
                lv.SubItems.Add(reader.GetString(4));
                listView3.Items.Add(lv);
            }
            reader.Close();
            connection.Close();
        }



        private void timer3_tick(object sender, EventArgs e)
        {
            if (Opacity == 1)
            {
                timer3.Stop();
            }
            Opacity += .2;
        }

        private void pfItemBox_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            connection.Open();
            SqlCommand cmd = connection.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select dimensions from [item] where id = @itid";
            cmd.Parameters.AddWithValue("@itid", pfItemBox.SelectedValue.ToString());
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                pfDimTxtBx.Text = dr.GetValue(0).ToString();
            }
            dr.Close();
            connection.Close();
        }

        private void sfItemBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            connection.Open();
            SqlCommand cmd = connection.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select type,form,dimensions from [cardboard] where id = @itid";
            cmd.Parameters.AddWithValue("@itid", comboBox4.SelectedValue.ToString());
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                sfTypeTxtBox.Text = dr.GetValue(0).ToString();
                sfFrmTxtBox.Text = dr.GetValue(1).ToString();
                sfDimTxtBox.Text = dr.GetValue(2).ToString();
            }
            dr.Close();
            connection.Close();
        }
        private void button4_Click(object sender, EventArgs e)
        {

        }
        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        //Creeate the Doc Method
        private void CreateWordDocument(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                //find and replace
                this.FindAndReplace(wordApp, "<cus>", this.pfCusBox.GetItemText(this.pfCusBox.SelectedItem));
                this.FindAndReplace(wordApp, "<add>", pfAdrTxtBx.Text);
                this.FindAndReplace(wordApp, "<ite>", this.pfItemBox.GetItemText(this.pfItemBox.SelectedItem));
                this.FindAndReplace(wordApp, "<dim>", pfDimTxtBx.Text);
                this.FindAndReplace(wordApp, "<qua>", pfQuaTxtBox.Text);
                this.FindAndReplace(wordApp, "<date>", DateTime.Now.ToString("dd/MM/yyyy"));
            }
            else
            {
                MessageBox.Show("File not Found!");
            }

            //Save as
            myWordDoc.SaveAs(ref SaveAs, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing);

            myWordDoc.Close();
            wordApp.Quit();
        }

        private void CreateWordDocument2(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                //find and replace
                this.FindAndReplace(wordApp, "<cus>", this.comboBox5.GetItemText(this.pfCusBox.SelectedItem));
                this.FindAndReplace(wordApp, "<add>", adrTxtBx2.Text);
                this.FindAndReplace(wordApp, "<ite>", this.comboBox4.GetItemText(this.pfItemBox.SelectedItem));
                this.FindAndReplace(wordApp, "<typ>", sfTypeTxtBox.Text);
                this.FindAndReplace(wordApp, "<for>", sfFrmTxtBox.Text);
                this.FindAndReplace(wordApp, "<dim>", sfDimTxtBox.Text);
                this.FindAndReplace(wordApp, "<qua>", sfQuaTxtBox.Text);
                this.FindAndReplace(wordApp, "<date>", DateTime.Now.ToString("dd/MM/yyyy"));
            }
            else
            {
                MessageBox.Show("File not Found!");
            }

            //Save as
            myWordDoc.SaveAs(ref SaveAs, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing);

            myWordDoc.Close();
            wordApp.Quit();
        }

        private void printBtn_Click_1(object sender, EventArgs e)
        {
            printing printing = new printing();
            printing.Show(this);
            this.Enabled = false;
            CreateWordDocument(Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FP-Gen", "temp1.docx")), Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FP-Gen", "gen1.docx")));
            Document doc = new Document();
            doc.LoadFromFile(Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FP-Gen", "gen1.docx")));
            PrintDocument printDoc = doc.PrintDocument;
            printDoc.PrintController = new StandardPrintController();
            printDoc.Print();
            this.Enabled = true;
            printing.Close();
            his();
        }

        private void sfPrintBtn_Click(object sender, EventArgs e)
        {
            printing printing = new printing();
            printing.Show(this);
            this.Enabled = false;
            CreateWordDocument2(Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FP-Gen", "temp2.docx")), Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FP-Gen", "gen2.docx")));
            Document doc = new Document();
            doc.LoadFromFile(Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FP-Gen", "gen2.docx")));
            PrintDocument printDoc = doc.PrintDocument;
            printDoc.PrintController = new StandardPrintController();
            printDoc.Print();
            this.Enabled = true;
            printing.Close();
            his();
        }

        private void adrTxtBx2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label38_Click(object sender, EventArgs e)
        {

        }
    }
}