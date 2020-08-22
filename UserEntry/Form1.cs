using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UserEntry
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            OleDbConnection connection = new OleDbConnection();

            string path = System.IO.Path.GetFullPath(@"..\..\") + @"\UserEntry.accdb;";
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
           
            try
            {
                connection.Open();

                string query = $@"INSERT INTO ContactTable (FirstName,LastName,Email,PhoneNumber,Age )
                            VALUES ({tbFirstName.Text}, {tbLastName.Text}, {tbEmail.Text}, {tbPhoneNumber.Text},{tbPhoneNumber.Text}";

                OleDbCommand cmd = new OleDbCommand(query, connection);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Data saved successfuly...!");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to" + ex.Message);
            }
            finally
            {
                connection.Close();
            }

        }
    }
}
