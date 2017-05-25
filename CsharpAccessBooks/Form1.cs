using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CsharpAccessBooks
{
    public partial class Form1 : Form
    {
        ADODB.Connection Con;
        ADODB.Recordset Rs;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Con = new ADODB.Connection();
            Rs = new ADODB.Recordset();
            Con.Provider = "Microsoft.jet.oledb.4.0";
            Con.ConnectionString = "G:\\Assignment 3  Access CSharp\\CsharpAccessBooks\\Books.mdb";
            Con.Open();
            Rs.Open("Select * from BooksTable", Con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic);
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            if (Rs.EOF == true && Rs.BOF == true)
            {
                MessageBox.Show("Table is empty!");
                return;
            }
            Rs.MoveFirst();
            ShowDataOnForm();
        }

        private void btnLast_Click(object sender, EventArgs e)
        {
            Rs.MoveLast();
            if (Rs.EOF == true)
            {
                MessageBox.Show("End of the table.");
                return;
            }
            ShowDataOnForm();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (Rs.EOF == true && Rs.BOF == true)
            {
                MessageBox.Show("Table is empty!");
                return;
            }
            Rs.MoveNext();
            if(Rs.EOF == true )
            {
                Rs.MovePrevious(); 
                MessageBox.Show("Passed end of file.");
            }
            ShowDataOnForm(); 
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            if (Rs.EOF == true && Rs.BOF == true)
            {
                MessageBox.Show("Table is empty!");
                return;
            }
            Rs.MovePrevious();
            if(Rs.BOF == true )
            {
                Rs.MoveNext(); 
                MessageBox.Show("Passed beginning of file.");
            }
            ShowDataOnForm(); 
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            ClearBoxes();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string FindRecord;

            if (txtISBN.Text == "")
            {
                MessageBox.Show("Please provide ID to Search");
                return;
            }
            else
            {
                FindRecord = "ISBN = " + txtISBN.Text;
                Rs.MoveFirst();
                Rs.Find(FindRecord);
                if (Rs.EOF == true)
                {
                    // Record has not been found 
                    MessageBox.Show("Record not found!");
                    return;
                }
                else
                {
                    ShowDataOnForm();
                    return;
                }

            }

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            string FindRecord;

            if (txtISBN.Text == "" && txtTitle.Text == "" && txtAuthor.Text == "" && txtCategory.Text == "" )
            {
                MessageBox.Show("Please fill up all the fields.");
                return;
            }
            else if (rdbYes.Checked == false && rdbNo.Checked == false)
            {
                MessageBox.Show("Please select Yes or No in Ebook.");
                return;
            }
            else
            {
                FindRecord = "ISBN = " + txtISBN.Text;
                Rs.MoveFirst();
                Rs.Find(FindRecord);
                if (Rs.EOF == true)
                {
                    // Record has not been found, so we can add it.
                    Rs.AddNew();
                    SaveInDataBase();
                    Rs.Update();
                    MessageBox.Show("Record saved succesfully!");
                    return;
                }
                else
                {
                    MessageBox.Show("Duplicated ID, use another one.");
                    ShowDataOnForm();
                    return;
                }

            }
        }

        private void btnModify_Click(object sender, EventArgs e)
        {
            string FindRecord;

            if (txtISBN.Text == "")
            {
                MessageBox.Show("Please provide ID to search and modify.");
                return;
            }
            else
            {
                FindRecord = "ISBN = " + txtISBN.Text;
                Rs.MoveFirst();
                Rs.Find(FindRecord);
                if (Rs.EOF == true)
                {
                    // Record has not been found. 
                    MessageBox.Show("Record not found, nothing can be modified.");
                    return;
                }
                else
                {
                    SaveInDataBase();
                    Rs.Update();
                    MessageBox.Show("Record modified succesfully!");
                    return;
                }

            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            string FindRecord;
            DialogResult Selection;

            if (txtISBN.Text == "")
            {
                MessageBox.Show("Please provide ID to search and delete.");
                return;
            }
            else
            {
                FindRecord = "ISBN = " + txtISBN.Text;
                Rs.MoveFirst();
                Rs.Find(FindRecord);
                if (Rs.EOF)
                {
                    // Record has not been found 
                    MessageBox.Show("Record not found, nothing can be deleted.");
                    return;
                }
                else
                {
                    Selection = MessageBox.Show("Are you sure?", "Confirm", MessageBoxButtons.YesNo);
                    if (Selection == DialogResult.Yes)
                    {
                        Rs.Delete();
                        ClearBoxes();
                        MessageBox.Show("Record deleted succesfully!");
                        Rs.Update();

                    }

                    return;
                }

            }
        }

#region "Functions"

        public void SaveInDataBase()
        {
            Rs.Fields["ISBN"].Value = Convert.ToInt32(txtISBN.Text);
            Rs.Fields["Title"].Value = txtTitle.Text;
            Rs.Fields["Author"].Value = txtAuthor.Text;
            Rs.Fields["Category"].Value = txtCategory.Text;
            if (rdbYes.Checked)
            {
                Rs.Fields["ebook"].Value = true;
            }
            else
            {
                Rs.Fields["ebook"].Value = false;

            }
        }

        public void ShowDataOnForm()
        {
            txtISBN.Text = Convert.ToString(Rs.Fields["ISBN"].Value);
            txtTitle.Text = Convert.ToString(Rs.Fields["Title"].Value);
            txtAuthor.Text = Convert.ToString(Rs.Fields["Author"].Value);
            txtCategory.Text = Convert.ToString(Rs.Fields["Category"].Value);
            if (Rs.Fields["ebook"].Value == true)
            {
                rdbYes.Checked = true;
                rdbNo.Checked = false;
            }
            else
            {
                rdbYes.Checked = false;
                rdbNo.Checked = true;
            }
           
        }

        public void ClearBoxes()
        {
            txtISBN.Clear();
            txtTitle.Clear();
            txtAuthor.Clear();
            txtCategory.Clear();
            rdbYes.Checked = false;
            rdbNo.Checked = false;
        }

    } 
}
#endregion "Functions"
