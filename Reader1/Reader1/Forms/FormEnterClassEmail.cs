using Reader1.Messages;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Reader1.Forms
{
    public partial class FormEnterClassEmail : Form
    {
        private string teacherEmail;
        private string validationCode;

        public string GetTeacherEmail
        {
            get { return teacherEmail; }
        }

        public void SetTeacherEmail(string teachEmail)
        {
            teacherEmail = teachEmail;
        }
        public FormEnterClassEmail(string teachEmail)
        {
            InitializeComponent();
            textBox1.Text = teachEmail;
            label3.Visible = false;
            textBox4.Visible = false;
            button1.Visible = false;
            button3.Visible = false;
            button4.Visible = false;
        }
        public void SetTextBoxTwoText(string teacherEmail)
        {
           textBox1.Text = teacherEmail;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void ButtonEnterVerifyCode(object sender, EventArgs e)
        {
            button3.Visible = false;
            button4.Visible = false;
            label3.Visible = true;
            textBox4.Visible = true;
            button1.Visible = true;
        }

        private void ButtonTryAgainEnterEmail(object sender, EventArgs e)
        {
            button3.Visible = false;
            button4.Visible = false;
            MessageBox.Show("Введите корректный почтовый адрес!", "Ошибка", MessageBoxButtons.OK);
        }

        private void ButtonClickConfirmCode(object sender, EventArgs e)
        {
            if (validationCode == textBox4.Text.Trim())
            {
                MessageBox.Show("Ваш почтовый адрес был подтвержден!", "Успех", MessageBoxButtons.OK);
                teacherEmail = textBox1.Text;
                this.Close();
            }
            else
            {
                MessageBox.Show("Код верификации некорректен!", "Ошибка", MessageBoxButtons.OK);
            }
        }

        private void ButtonClickConfirmEmail(object sender, EventArgs e)
        {
            teacherEmail = textBox1.Text;

            bool flag = false;

            validationCode = GenerateValidationCode();
            string fullValidationMessage = "Ваш код для подтверждения почты:\n" + validationCode;

            flag |= StartMessage.MailChecking("arserm8@gmail.com", teacherEmail, "", "PsychologistProblem", fullValidationMessage, "");

            if (flag == false)
            {
                MessageBox.Show("Введите корректный почтовый адрес!", "Ошибка", MessageBoxButtons.OK);
            }
            else
            {
                button3.Visible = true;
                button4.Visible = true;
            }
        }
        private string GenerateValidationCode()
        {
            Random random = new Random();
            return random.Next(100000, 1000000).ToString("D6"); // Генерация шестизначного кода
        }

        //private void FormEnterClassEmail_Load(object sender, EventArgs e)
        //{
        //    throw new System.NotImplementedException();
        //}
    }
}
