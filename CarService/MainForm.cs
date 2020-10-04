using System;
using System.Data;
using System.Data.SqlClient;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace JobCentre
{
    public partial class MainForm : Form
    {
       int count = 0;
        DataSet ds;
        SqlDataAdapter adapter;
        DataView View;

        string connectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=JobCentreBD2;Integrated Security=True"; //+
        string sql = "SELECT * FROM Employment_Office; " +
            "SELECT * FROM Service_Hotline; " +
            "SELECT * FROM Personnel_EO; " +
            "SELECT * FROM Posts; " +
            "SELECT * FROM Vacancy; " +
            "SELECT * FROM Employers; " +
            "SELECT * FROM Deals; " +
            "SELECT * FROM Job_Seekers; " +
            "SELECT * FROM Services_Employers; " +
            "SELECT * FROM Services_JobSeekers;"; //
        string sql2 = "select * from Employment_Office"; //
        public MainForm()
        {
            InitializeComponent();
            InitComp();
        }
        public MainForm(string role)
        {
            InitializeComponent();
            InitComp();
            if (role == "admin")
            {
                button1.Enabled = true;
                //создатьЗаказToolStripMenuItem.Enabled = true;
            }
        }
        void InitComp()
        {
            comboBox1.Items.Add("Employment_Office"); // 
            comboBox1.Items.Add("Service_Hotline"); //
            comboBox1.Items.Add("Personnel_EO"); // 
            comboBox1.Items.Add("Posts");
            comboBox1.Items.Add("Vacancy");//
            comboBox1.Items.Add("Employers");
            comboBox1.Items.Add("Deals");
            comboBox1.Items.Add("Job_Seekers");
            comboBox1.Items.Add("Services_Employers");
            comboBox1.Items.Add("Services_JobSeekers");
            comboBox1.Text = "Employment_Office"; // 
            comboBox2.Items.Add("Выборка");
            comboBox2.Items.Add("Выборка по параметру");
            comboBox2.Items.Add("Создать запрос");
            comboBox2.Items.Add("Новая строка");
            comboBox2.Items.Add("Удалить строки");
            comboBox2.Items.Add("Обновить Базу Данных");
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList; //+
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AutoResizeColumns();
            button1.Enabled = false;
            //создатьЗаказToolStripMenuItem.Enabled = false;


            ExecuteSQL();
        }
        private void ExecuteSQL()
        {

            using (SqlConnection connection = new SqlConnection(connectionString))
            {

                connection.Open();
                adapter = new SqlDataAdapter(sql, connection);
                //adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;

                if (count == 0)
                {
                    ds = CreateRelationsDBDataSet();
                    count++;
                }
                ds = new DataSet("EmploymentOfficeDB");
                adapter.Fill(ds);

                View = new DataView(ds.Tables[0]);
                dataGridView1.DataSource = ds.Tables[0];
                dataGridView2.DataSource = View;
                dataGridView1.Columns[0].ReadOnly = true;
                dataGridView2.Columns[0].ReadOnly = true; //+
                count++;
            }
        }

        private static DataSet CreateRelationsDBDataSet()
        {
            string connectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=JobCentreBD2;Integrated Security=True"; //+
            string sql = "SELECT * FROM Employment_Office; " +
                "SELECT * FROM Service_Hotline; " +
                "SELECT * FROM Personnel_EO; " +
                "SELECT * FROM Posts; " +
                "SELECT * FROM Vacancy; " +
                "SELECT * FROM Employers; " +
                "SELECT * FROM Deals; " +
                "SELECT * FROM Job_Seekers; " +
                "SELECT * FROM Services_Employers; " +
                "SELECT * FROM Services_JobSeekers;"; //

            DataSet JobCentreDB = new DataSet("ShopDB");

            // создание адаптера данных для базы данных ShopDB
            SqlDataAdapter adapter = new SqlDataAdapter(sql, connectionString);
            adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;

            //Мапинг таблиц БД
            adapter.TableMappings.Add("Table", "Employment_Office");
            adapter.TableMappings.Add("Table1", "Service_Hotline");
            adapter.TableMappings.Add("Table2", "Personnel_EO");
            adapter.TableMappings.Add("Table3", "Posts");
            adapter.TableMappings.Add("Table4", "Vacancy");
            adapter.TableMappings.Add("Table5", "Employers");
            adapter.TableMappings.Add("Table6", "Deals");
            adapter.TableMappings.Add("Table7", "Job_Seekers");
            adapter.TableMappings.Add("Table8", "Services_Employers");
            adapter.TableMappings.Add("Table9", "Services_JobSeekers");

            adapter.Fill(JobCentreDB); //

            //Получение ссылок на таблицы
            var employment_Office = JobCentreDB.Tables["Employment_Office"];
            var service_Hotline = JobCentreDB.Tables["Service_Hotline"];
            var personnel_EO = JobCentreDB.Tables["Personnel_EO"];
            var posts = JobCentreDB.Tables["Posts"];
            var vacancy = JobCentreDB.Tables["Vacancy"];
            var employers = JobCentreDB.Tables["Employers"];
            var deals = JobCentreDB.Tables["Deals"];
            var job_Seekers = JobCentreDB.Tables["Job_Seekers"];
            var services_Employers = JobCentreDB.Tables["Services_Employers"];
            var services_JobSeekers = JobCentreDB.Tables["Services_JobSeekers"];


            //создание связей для таблиц
            JobCentreDB.Relations.Add("Employment_Office_Service_Hotline", employment_Office.Columns["ID_ЦентраЗанятости"], service_Hotline.Columns["ID_ЦентраЗанятости"], true);
            JobCentreDB.Relations.Add("Employment_Office_Personnel_EO", employment_Office.Columns["ID_ЦентраЗанятости"], personnel_EO.Columns["ID_ЦентраЗанятости"], true);
            JobCentreDB.Relations.Add("Employment_Office_Services_JobSeekers", employment_Office.Columns["ID_ЦентраЗанятости"], services_JobSeekers.Columns["ID_ЦентраЗанятости"], true);
            JobCentreDB.Relations.Add("Employment_Office_Services_Employers", employment_Office.Columns["ID_ЦентраЗанятости"], services_Employers.Columns["ID_ЦентраЗанятости"], true);
            JobCentreDB.Relations.Add("Employment_Office_Deals", employment_Office.Columns["ID_ЦентраЗанятости"], deals.Columns["ID_ЦентраЗанятости"], true);
            JobCentreDB.Relations.Add("job_Seekers_Deals", job_Seekers.Columns["ID"], deals.Columns["Код_Соискателя"], true);
            JobCentreDB.Relations.Add("Vacancy_Deals", vacancy.Columns["Код_Вакансии"], deals.Columns["Код_Вакансии"], true);
            JobCentreDB.Relations.Add("Posts_job_Seekers", posts.Columns["Должность"], job_Seekers.Columns["Квалификация"], true);
            JobCentreDB.Relations.Add("Posts_Vacancy", posts.Columns["Должность"], vacancy.Columns["Должность"], true);
            JobCentreDB.Relations.Add("Employers_Vacancy", employers.Columns["Код_Организации"], vacancy.Columns["Код_Организации"], true);

            //var FK_Employment_Office_Service_Hotline = employment_Office.Constraints["FK__Service_H__ID_Це__1367E606"] as ForeignKeyConstraint;
            //FK_Employment_Office_Service_Hotline.DeleteRule = Rule.Cascade;

            return JobCentreDB;
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Authification auth = new Authification();
            auth.Show();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            sql = $"SELECT * FROM [{comboBox1.Text}]";
            ExecuteSQL();
            labelLog(comboBox1.Text);
        }

        void labelLog(string text)
        {
            textBox1.Text = text;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            switch (comboBox2.Text)
            {
                case "Выборка":
                    select();
                    break;
                case "Выборка по параметру":
                    selectWhere();
                    break;
                case "Создать запрос":
                    createQuerry();
                    break;
                case "Новая строка":
                    addStrings();
                    break;
                case "Удалить строки":
                    deleteStrings();
                    break;
                case "Обновить Базу Данных":
                    SaveDB();
                    break;

            }
        }
        void select() 
        {
            SelectForm selForm = new SelectForm();

            DialogResult result = selForm.ShowDialog(this);

            if (result == DialogResult.Cancel)
                return;

            sql = $"SELECT {selForm.comboBox1.SelectedItem.ToString()} from [{selForm.comboBox2.SelectedItem.ToString()}]";
            labelLog(selForm.comboBox2.SelectedItem.ToString());
            ExecuteSQL();
        }

        void selectWhere() 
        {
            SelectWhereForm selForm = new SelectWhereForm();

            DialogResult result = selForm.ShowDialog(this);

            if (result == DialogResult.Cancel)
                return;

            sql = $"SELECT {selForm.comboBox1.SelectedItem.ToString()} from [{selForm.comboBox2.SelectedItem.ToString()}] where {selForm.textBox1.Text}";
            labelLog(selForm.comboBox2.SelectedItem.ToString());
            ExecuteSQL();
        }

        void createQuerry()
        {
            QuerryForm selForm = new QuerryForm();

            DialogResult result = selForm.ShowDialog(this);

            if (result == DialogResult.Cancel)
                return;

            sql = $"{selForm.textBox1.Text}";
            labelLog("Собственная выборка");
            ExecuteSQL();
        }

        void addStrings()
        {
            DataRow row = ds.Tables[0].NewRow();
            ds.Tables[0].Rows.Add(row);
        }

        void deleteStrings()
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.Remove(row);
            }
        }

        void SaveDB() //+
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlDataAdapter adapter = new SqlDataAdapter(sql, connection);

                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(adapter);
                adapter.Update(ds);
                ds.Clear();
                adapter.Fill(ds);
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e) //+
        {
            Close();
        }
        
        private void количествоРабочихToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sql = "SELECT Employment_Office.Название as 'Центр занятости', " +
                "count(Personnel_EO.ID_ЦентраЗанятости) AS 'Количество рабочих' " +
                "FROM Employment_Office " +
                "INNER JOIN Personnel_EO " +
                "ON Employment_Office.ID_ЦентраЗанятости = Personnel_EO.ID_ЦентраЗанятости " +
                "GROUP BY Employment_Office.Название";
            sql2 = "select * from Employment_Office";
            ExecuteSQL();
            labelLog("Количество рабочих ЦЗ");
        }
        private void количествоКонтрактовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sql = "SELECT Employment_Office.Название, " +
                "count(Deals.ID_ЦентраЗанятости) AS 'Количество контрактов' " +
                "FROM Deals " +
                "INNER JOIN Employment_Office " +
                "ON Employment_Office.ID_ЦентраЗанятости = Deals.ID_ЦентраЗанятости " +
                "GROUP BY Employment_Office.Название";
            sql2 = "select * from Employment_Office";
            ExecuteSQL();
            labelLog("Отчетность по контрактам ЦЗ");
        }
        
         private void спросНаСпециальностьToolStripMenuItem_Click(object sender, EventArgs e)
            {
            sql = "SELECT Posts.Должность, " +
                "count(Job_Seekers.Квалификация) AS 'Спрос' " +
                "FROM Job_Seekers " +
                "INNER JOIN Posts " +
                "ON Posts.Должность = Job_Seekers.Квалификация " +
                "GROUP BY Posts.Должность";
            sql2 = "select * from Employment_Office";
            ExecuteSQL();
            labelLog("Спрос на специальность");
        }
        private void спросНаСпециальностьпоГородуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sql = "SELECT Job_Seekers.Город, Posts.Должность, count(Job_Seekers.Квалификация) AS 'Спрос' " +
                "FROM Job_Seekers " +
                "JOIN Posts " +
                "ON Posts.Должность = Job_Seekers.Квалификация " +
                "GROUP BY Posts.Должность, Job_Seekers.Город";
            ExecuteSQL();
            sql2 = "select * from Employment_Office";
            labelLog("Спрос на специальность (по городу)");
        }
        private void предложениеПоКвалификациямToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sql = "SELECT Posts.Должность AS 'Квалификация', count(Vacancy.Должность) AS 'Предложения' " +
                "FROM Vacancy " +
                "JOIN Posts ON Posts.Должность = Vacancy.Должность " +
                "GROUP BY Posts.Должность";
            ExecuteSQL();
            sql2 = "select * from Employment_Office";
            labelLog("Предложения по квалификациям");
        }
        private void предложенияПоКвалификациямпоГородуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sql = "SELECT Vacancy.Город, Posts.Должность AS 'Квалификация', count(Vacancy.Должность) AS 'Предложения' " +
                "FROM Vacancy " +
                "FULL OUTER JOIN Posts " +
                "ON Posts.Должность = Vacancy.Должность " +
                "GROUP BY Posts.Должность, Vacancy.Город";
            ExecuteSQL();
            sql2 = "select * from Employment_Office";
            labelLog("Предложения по квалификациям");
        }
        private void количествоБезработныхпоСпециальностямToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sql = "SELECT Job_Seekers.ID, Job_Seekers.Квалификация, " +
                "COUNT(Deals.Код_Соискателя) AS 'Сделки', COUNT(Job_Seekers.ID) AS 'Предложения', " +
                "(COUNT(Job_Seekers.ID) - COUNT(Deals.Код_Соискателя)) AS 'Безработных' " +
                "FROM Deals " +
                "FULL OUTER JOIN Job_Seekers " +
                "ON Deals.Код_Соискателя = Job_Seekers.ID " +
                "GROUP BY Job_Seekers.Квалификация, Deals.Код_Соискателя, Job_Seekers.ID";
            ExecuteSQL();
            sql2 = "select * from Employment_Office";
            labelLog("Безработные (по специальностям)");
        }
       

        private void wordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int columns = dataGridView1.ColumnCount;
            int rows = dataGridView1.RowCount;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            Microsoft.Office.Interop.Word.Paragraph wordParag;
            Microsoft.Office.Interop.Word.Table wordTable;

            Object missing = Type.Missing;
            wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);

            wordDoc = wordApp.ActiveDocument;

            wordParag = wordDoc.Paragraphs.Add(Type.Missing);

            wordParag = wordDoc.Paragraphs.Add(Type.Missing);
            wordParag.Range.Font.Name = "Times New Roman";
            wordParag.Range.Font.Size = 16;
            wordParag.Range.Font.Bold = 1;
            wordParag.Range.Text = $"Биржа труда";
            wordParag.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

            wordDoc.Paragraphs.Add(Type.Missing);
            wordParag.Range.Font.Name = "Times New Roman";
            wordParag.Range.Font.Size = 16;
            wordParag.Range.Font.Bold = 1;
            wordParag.Range.Text = $"Отчет по выборке";
            wordParag.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

            wordDoc.Paragraphs.Add(Type.Missing);
            wordParag.Range.Font.Name = "Times New Roman";
            wordParag.Range.Font.Size = 10;
            wordParag.Range.Font.Bold = 1;
            wordParag.Range.Text = $"Вывод {textBox1.Text}";
            wordParag.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

            DateTime dt = DateTime.Now;
            string curDate = dt.ToShortDateString();


            wordDoc.Paragraphs.Add(Type.Missing);
            wordParag.Range.Tables.Add(wordParag.Range, rows, columns, Type.Missing, Type.Missing);
            wordTable = wordDoc.Tables[1];


            for (int j = 0; j < rows; j++)
                for (int i = 0; i < columns; i++)
                {
                    Word.Range wordcellrange = wordTable.Cell(j + 1, i + 1).Range;
                    wordcellrange.Text = dataGridView1[i, j].Value.ToString();
                    wordcellrange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle =
           wordcellrange.Borders[Word.WdBorderType.wdBorderRight].LineStyle =
           wordcellrange.Borders[Word.WdBorderType.wdBorderTop].LineStyle =
           wordcellrange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle =
           wordcellrange.Borders[Word.WdBorderType.wdBorderHorizontal].LineStyle =

           Word.WdLineStyle.wdLineStyleDouble;

                }
            wordApp.Visible = true;
            wordDoc.Paragraphs.Add(Type.Missing);
            wordParag.Range.Font.Name = "Times New Roman";
            wordParag.Range.Font.Size = 16;
            wordParag.Range.Font.Bold = 0;
            wordParag.Range.Text = $"Подпись ____________\tДата: {curDate}";
            wordParag.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            sql = sql2;
            ExecuteSQL();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            View.RowFilter = textBox3.Text;
            View.Sort = textBox2.Text;
            View.RowStateFilter = (DataViewRowState)Enum.Parse(typeof(DataViewRowState), comboBox3.Text, true);

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void xMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ds.ReadXmlSchema(@"D:\ADO.NET\JobCentreBDSchema.xml");
            ds.ReadXml(@"D:\ADO.NET\JobCentreBDData.xml");
        }

        private void adminToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ds.WriteXmlSchema(@"D:\ADO.NET\JobCentreBDSchema.xml"); // запсиь схемы ShopDB в XML файл
            ds.WriteXml(@"D:\ADO.NET\JobCentreBDData.xml"); // запись данных ShopDB в XML файл
        }

    }
}
