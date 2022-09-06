using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace XL2MySQL
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StatusText.Clear();
            if (txtInputFile.Text.Trim() == "")
            {
                
                StatusText.AppendText("Please select the input file \n");
                return;
            }
            xlImport objImport = new xlImport();
            objImport.server = ConfigurationManager.AppSettings["server"];
            objImport.uid = ConfigurationManager.AppSettings["uid"];
            objImport.pwd =ConfigurationManager.AppSettings["pwd"];
            objImport.database = ConfigurationManager.AppSettings["database"];
            objImport.xlFilePath = txtInputFile.Text;
            objImport.txtStatus = StatusText;
            if (checkRemove.Checked)
                objImport.clean = true;
            else
                objImport.clean = false;
            objImport.import();
            StatusText.AppendText("Done\n");
            StatusText.Refresh();

            MessageBox.Show("Import of records is complete", "Import", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog.FileName = "";
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtInputFile.Text = openFileDialog.FileName;
            }
        }
    }
}
