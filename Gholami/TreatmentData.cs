using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;

namespace Gholami
{
    public partial class TreatmentData : Form
    {
        private string ConnectionString = string.Format(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = Treatment.accdb",Application.StartupPath.ToString());
        //private string ConnectionString = string.Format(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = Treatment.mdb", Application.StartupPath.ToString());
        public void SelectPatientAllRecords()
        {
            OleDbConnection OleDbConnection1 = new OleDbConnection(ConnectionString);
            OleDbCommand OleDbCommand1 = new OleDbCommand();
            OleDbCommand1.CommandType = CommandType.Text;
            OleDbCommand1.CommandText = "SELECT PatientId, FirstName, LastName  FROM Patient";
            //OleDbCommand1.CommandText = "SELECT * FROM Patient where id=1";
            OleDbCommand1.Connection = OleDbConnection1;
            OleDbDataAdapter adapter = new OleDbDataAdapter(OleDbCommand1);
            DataSet DataSet1 = new DataSet();
            adapter.Fill(DataSet1, "SelectPatientAllRecords");
            dataGridView1.DataSource = DataSet1.Tables["SelectPatientAllRecords"];
        }
        public void ShowPatientById(string pId)
        {
            //string query = "SELECT * FROM Patient WHERE PatientId LIKE '%" + this.txtPI.Text + "%'";
            string query = "SELECT PatientId, FirstName, LastName FROM Patient WHERE PatientId = " + pId;
            OleDbConnection OleDbConnection1 = new OleDbConnection(ConnectionString);
            OleDbCommand OleDbCommand1 = new OleDbCommand();
            OleDbCommand1.CommandType = CommandType.Text;
            OleDbCommand1.CommandText = query;
            OleDbCommand1.Connection = OleDbConnection1;
            OleDbDataAdapter adapter = new OleDbDataAdapter(OleDbCommand1);
            DataSet DataSet1 = new DataSet();
            adapter.Fill(DataSet1, "ShowPatientById");
            dataGridView1.DataSource = DataSet1.Tables["ShowPatientById"];
        }

        public int SearchPatientById(string pId)
        {
            int result;
            string query = "SELECT PatientId FROM Patient WHERE PatientId = " + pId;
            OleDbConnection OleDbConnection1 = new OleDbConnection(ConnectionString);
            OleDbCommand OleDbCommand1 = new OleDbCommand();
            OleDbCommand1.CommandType = CommandType.Text;
            OleDbCommand1.CommandText = query;
            OleDbCommand1.Connection = OleDbConnection1;
            OleDbDataAdapter adapter = new OleDbDataAdapter(OleDbCommand1);
            DataSet DataSet1 = new DataSet();
            adapter.Fill(DataSet1, "SearchPatientById");
            dataGridView5.DataSource = DataSet1.Tables["SearchPatientById"];       
            if (this.dataGridView5.Rows.Count == 1)
                result = 1;
            else result = 0;
            return result;
        }

        public TreatmentData()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.SelectPatientAllRecords();
        }
        private void Add_Click(object sender, EventArgs e)
        {
            int pIdisExist = this.SearchPatientById(this.txtPI.Text);
            if (pIdisExist == 1)
            {
                OleDbConnection OleDbConnection1 = new OleDbConnection(ConnectionString);
                OleDbCommand OleDbCommand1 = new OleDbCommand();
                OleDbCommand1.Connection = OleDbConnection1;
                OleDbCommand1.CommandType = CommandType.Text;
                OleDbCommand1.CommandText = "INSERT INTO Patient(FirstName, LastName, PatientId) VALUES (@FName,@LName,@PI)";
                OleDbCommand1.Parameters.AddWithValue("@FName", this.txtFName.Text);
                OleDbCommand1.Parameters.AddWithValue("@LName", this.txtLName.Text);
                OleDbCommand1.Parameters.AddWithValue("@PI", this.txtPI.Text);
                OleDbConnection1.Open();
                int result = OleDbCommand1.ExecuteNonQuery();
                OleDbConnection1.Close();
                if (result > 0)
                {
                    MessageBox.Show("Added Successfully");
                    this.ShowPatientById(this.txtPI.Text);
                    //this.SelectPatientAllRecords();
                    this.Log("Patient:" + "(" + this.txtPI.Text + ")" + this.txtFName.Text + " " + this.txtLName.Text + " " + "Added Successfully!");
                }
                else
                {
                    MessageBox.Show("Error");
                    this.SelectPatientAllRecords();
                }
            }
            else
            {
                MessageBox.Show("This PatientId is Added Before!");
                this.ShowPatientById(this.txtPI.Text);
            }
        }

        private void BtnSaveTreat_Click(object sender, EventArgs e)
        {
            int pIdisExist = this.SearchPatientById(this.txtPP.Text);
            if (pIdisExist != 1)
            {
                OleDbConnection OleDbConnection3 = new OleDbConnection(ConnectionString);
                OleDbCommand OleDbCommand3 = new OleDbCommand();
                OleDbCommand3.Connection = OleDbConnection3;
                OleDbCommand3.CommandType = CommandType.Text;
                OleDbCommand3.CommandText = "INSERT INTO Treatment(PatientId,TreatDate, X, Y, Z, TreatmentSite, TreatImage, TreatApply) VALUES (@PatientId, @TreatDate, @X, @Y, @Z, @TreatmentSite, @TreatImage, @CBApply)";
                OleDbCommand3.Parameters.AddWithValue("@PatientId", this.txtPP.Text);
                OleDbCommand3.Parameters.AddWithValue("@TreatDate", this.dateTimePicker1.Value.ToShortDateString());
                OleDbCommand3.Parameters.AddWithValue("@X", this.txtx.Text);
                OleDbCommand3.Parameters.AddWithValue("@Y", this.txty.Text);
                OleDbCommand3.Parameters.AddWithValue("@Z", this.txtz.Text);
                OleDbCommand3.Parameters.AddWithValue("@TreatmentSite", this.CBTSite.GetItemText(this.CBTSite.SelectedItem));
                OleDbCommand3.Parameters.AddWithValue("@TreatImage", this.CBImage.GetItemText(this.CBImage.SelectedItem));
                OleDbCommand3.Parameters.AddWithValue("@CBApply", this.CBApply.GetItemText(this.CBApply.SelectedItem));
                OleDbConnection3.Open();
                int result = OleDbCommand3.ExecuteNonQuery();
                OleDbConnection3.Close();
                if (result > 0)
                {
                    MessageBox.Show("Added Successfully");
                    OleDbConnection OleDbConnection2 = new OleDbConnection(ConnectionString);
                    OleDbCommand OleDbCommand2 = new OleDbCommand();
                    OleDbCommand2.CommandType = CommandType.Text;
                    OleDbCommand2.CommandText = "SELECT p.FirstName,p.LastName,t.PatientId, t.TreatDate, t.X, t.Y, t.Z, t.TreatmentSite,t.TreatImage,t.TreatApply FROM Treatment AS t INNER JOIN Patient AS p ON t.PatientId = p.PatientId Where t.PatientId =" + txtPP.Text + " ORDER by t.TreatDate DESC";
                    OleDbCommand2.Connection = OleDbConnection2;
                    OleDbDataAdapter adapter = new OleDbDataAdapter(OleDbCommand2);
                    DataSet DataSet1 = new DataSet();
                    adapter.Fill(DataSet1, "Treatment");
                    dataGridView2.DataSource = DataSet1.Tables["Treatment"];

                }
                else
                {
                    MessageBox.Show("Error");
                }
            }
            else
            {
                MessageBox.Show("First Add Patient please!");
            }
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            OleDbConnection OleDbConnection2 = new OleDbConnection(ConnectionString);
            OleDbCommand OleDbCommand2 = new OleDbCommand();
            OleDbCommand2.CommandType = CommandType.Text;
            OleDbCommand2.CommandText = "SELECT p.FirstName,p.LastName,t.PatientId, t.TreatDate, t.X, t.Y, t.Z, t.TreatmentSite,t.TreatImage,t.TreatApply FROM Treatment AS t INNER JOIN Patient AS p ON t.PatientId = p.PatientId Where t.PatientId =" + txtSearch.Text + " ORDER by t.TreatDate DESC";
            OleDbCommand2.Connection = OleDbConnection2;
            OleDbDataAdapter adapter = new OleDbDataAdapter(OleDbCommand2);
            DataSet DataSet1 = new DataSet();
            adapter.Fill(DataSet1, "Treatment");
            dataGridView3.DataSource = DataSet1.Tables["Treatment"];
        }

        private void BtnMsearch_Click(object sender, EventArgs e)
        {
            OleDbConnection OleDbConnection2 = new OleDbConnection(ConnectionString);
            OleDbCommand OleDbCommand2 = new OleDbCommand();
            OleDbCommand2.CommandType = CommandType.Text;
            OleDbCommand2.Connection = OleDbConnection2;
            OleDbCommand2.CommandText = "SELECT count(X),AVG(X), AVG(Y), AVG(Z),TreatmentSite FROM Treatment GROUP by TreatmentSite";
            OleDbDataAdapter adapter = new OleDbDataAdapter(OleDbCommand2);
            DataSet DataSet1 = new DataSet();
            adapter.Fill(DataSet1, "Treatment");
            dataGridView4.DataSource = DataSet1.Tables["Treatment"];
            dataGridView4.Columns[0].HeaderText = "Total Treatment Count";
            dataGridView4.Columns[1].HeaderText = "X-Mean";
            dataGridView4.Columns[2].HeaderText = "Y-Mean";
            dataGridView4.Columns[3].HeaderText = "Z-Mean";
        }

        private void BtnTre_Click(object sender, EventArgs e)
        {
            OleDbConnection OleDbConnection2 = new OleDbConnection(ConnectionString);
            OleDbCommand OleDbCommand2 = new OleDbCommand();
            OleDbCommand2.CommandType = CommandType.Text;
            OleDbCommand2.Connection = OleDbConnection2;
            OleDbCommand2.CommandText = "SELECT count(X),AVG(X), AVG(Y), AVG(Z),TreatmentSite FROM Treatment Where PatientId = " + txtSearch.Text + " GROUP by TreatmentSite";
            OleDbDataAdapter adapter = new OleDbDataAdapter(OleDbCommand2);
            DataSet DataSet1 = new DataSet();
            adapter.Fill(DataSet1, "Treatment");
            dataGridView3.DataSource = DataSet1.Tables["Treatment"];
            dataGridView3.Columns[0].HeaderText = "Total Treatment Count";
            dataGridView3.Columns[1].HeaderText = "X-Mean";
            dataGridView3.Columns[2].HeaderText = "Y-Mean";
            dataGridView3.Columns[3].HeaderText = "Z-Mean";
        }

        public void Log(string message)
        {
            File.AppendAllText("E:\\Me\\Project\\Gholami\\log.txt", DateTime.Now.ToString() + "  -  " + message + "\n");
        }

    }
}
