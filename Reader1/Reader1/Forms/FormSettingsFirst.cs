using System;
using System.Net;
using System.Net.Mail;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Config = Reader1.Models.Configuration.Configuration;
using Reader1;
using Reader1.Messages;
using Reader1.Forms;
using System.Runtime.Serialization.Formatters;

namespace Reader1.Forms
{
    public partial class FormSetting : Form
    {

        public FormSetting()
        {
            InitializeComponent();
            ChangedPages();
        }

        private void buttonFather_Click(object sender, EventArgs e)
        {
            Config config = new Config();
            config.OrganisationName = textBox1.Text;
            config.ClassNumber = textBox2.Text;
            config.ClassLetter = textBox3.Text;
            config.ClassroomTeacherName = textBox4.Text;
            // config.ClassroomTeacherEmail = "arsfaraway@gmail.com";
            config.ClassroomTeacherEmail = textBox5.Text;
            config.ClassroomTeacherPhone = textBox6.Text;
            config.PsychologistName = textBox7.Text;
            config.PsychologistEmail = textBox8.Text;
            config.PsychologistPhone = textBox9.Text;
            config.AcademicYearStartReporting = textBox10.Text;
            config.ReportingStartQuarterNumber = textBox11.Text;
            config.ReportingAcademicYear = textBox12.Text;
            config.ReportingQuarterNumber = textBox13.Text;

          
            FormEnterClassEmail formEnterClassEmail = new FormEnterClassEmail(config.ClassroomTeacherEmail);

            formEnterClassEmail.SetTeacherEmail(config.ClassroomTeacherEmail);
            formEnterClassEmail.ShowDialog();

            config.ClassroomTeacherEmail = formEnterClassEmail.GetTeacherEmail;

            // todo проверить психолога на правильность почты! Только как?
            // todo сохранить информацию в базу данных

            this.Close();
        }

        private void label2_Click(object sender, EventArgs e)
        {
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }


        private void FormSetting_Load(object sender, EventArgs e)
        {
            tabControl1.SelectedIndexChanged += TabControl1_SelectedIndexChanged;
        }

        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedIndex = tabControl1.SelectedIndex;

            if (selectedIndex == 0)
            {
                ChangedPages();
            }
            else if (selectedIndex == 1)
            {
                ChangedPages();
            }
        }

        private void buttonThen_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = tabControl1.SelectedIndex + 1;
        }

        private void buttonPrev_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = tabControl1.SelectedIndex - 1;
        }

        private void ChangedPages()
        {
            buttonNext.Visible = tabControl1.SelectedIndex != tabControl1.TabPages.Count - 1;
            buttonPrev.Visible = tabControl1.SelectedIndex != 0;
            buttonSave.Visible = tabControl1.SelectedIndex != 0;
        }
    }
}
