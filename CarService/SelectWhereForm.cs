using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace JobCentre
{
    public partial class SelectWhereForm : Form
    {
        DataSet ds;
        SqlDataAdapter adapter;
        string connectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=JobCentreBD2;Integrated Security=True"; //
        string sql = "SELECT * FROM Employment_Office"; //
        public SelectWhereForm()
        {
            InitializeComponent();
            ExecuteComp();
        }
        void ExecuteComp()
        {
            comboBox2.Items.Add("Employment_Office"); //
            comboBox2.Items.Add("Service_Hotline"); //
            comboBox2.Items.Add("Personnel_EO"); // 
            comboBox2.Items.Add("Posts");
            comboBox2.Items.Add("Vacancy");//
            comboBox2.Items.Add("Employers");
            comboBox2.Items.Add("Deals");
            comboBox2.Items.Add("Job_Seekers");
            comboBox2.Items.Add("Services_Employers");
            comboBox2.Items.Add("Services_JobSeekers");
            comboBox2.Text = "Employment_Office";
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            sql = $"SELECT * from dbo.[{comboBox2.Text}]";
            comboBox1.Items.Clear();
            SQLExecute();
        }
        private void SQLExecute()
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                adapter = new SqlDataAdapter(sql, connection);

                ds = new DataSet();
                adapter.Fill(ds);
                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataColumn column in dt.Columns)
                        comboBox1.Items.Add(column.ColumnName);
                }
                comboBox1.Items.Add("*");

            }
        }
    }
}
