
using System;
using System.Windows.Forms;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
namespace JobCentre
{
    public partial class Authification : Form
    {
        public string login = string.Empty;
        public string password = string.Empty;
        private Users user = new Users(); 
        // Экземпляр класса пользователей.
        public Authification()
        {
            InitializeComponent();
            LoadUsers(); 
            // Метод десериализующий класс.
        }


        private void LoadUsers()
        {
            try
            {
                FileStream fs = new FileStream("Users.dat", FileMode.Open);

                BinaryFormatter formatter = new BinaryFormatter();

                user = (Users)formatter.Deserialize(fs);

                fs.Close();
            }
            catch { return; }
        }

        private void EnterToForm()
        {
            if ("admin" == loginTextBox.Text && "admin" == passwordTextBox.Text)
            {
                MainForm mainForm = new MainForm("admin");
                mainForm.Show();
                this.Hide();
                return;
            }

            for (int i = 0; i < user.Logins.Count; i++) // Ищем пользователя и проверяем правильность пароля.
            {
                if (user.Logins[i] == loginTextBox.Text && user.Passwords[i] == passwordTextBox.Text)
                {
                    login = user.Logins[i];
                    password = user.Passwords[i];

                    MainForm mainForm = new MainForm();
                    mainForm.Show();
                    this.Hide();


                }
                else if (user.Logins[i] == loginTextBox.Text && passwordTextBox.Text != user.Passwords[i])
                {
                    login = user.Logins[i];

                    MessageBox.Show("Неверный пароль!");
                }
            }

            if (login == "") { MessageBox.Show("Пользователь " + loginTextBox.Text + " не найден!"); }
        }

        private void AddUser() // Регистрируем нового пользователя.
        {
            if (loginTextBox.Text == "" || passwordTextBox.Text == "") { MessageBox.Show("Не введен логин или пароль!"); return; }

            user.Logins.Add(loginTextBox.Text);
            user.Passwords.Add(passwordTextBox.Text);

            FileStream fs = new FileStream("Users.dat", FileMode.OpenOrCreate);

            BinaryFormatter formatter = new BinaryFormatter();
            formatter.Serialize(fs, user); // Сериализуем класс.

            fs.Close();

            login = loginTextBox.Text;

            this.Close();
        }

        private void exitButton_Click(object sender, EventArgs e)
        {
            Application.Exit(); // Закрываем программу.
        }

        private void regButton_Click(object sender, EventArgs e)
        {
            AddUser();
        }

        private void enterButton_Click(object sender, EventArgs e)
        {
            EnterToForm();
        }

        private void RegistrationForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (login == "" | password == "") { Application.Exit(); }
        }
    }
}

