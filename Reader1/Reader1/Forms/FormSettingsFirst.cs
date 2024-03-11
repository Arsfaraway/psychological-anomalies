using System;
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


namespace Reader1.Forms
{
    public partial class FormSetting : Form
    {

        public FormSetting()
        {
            InitializeComponent();
            a();
        }

        private void buttonFather_Click(object sender, EventArgs e)
        {
            // MessageBox.Show("Hello, world.");

            //form2.Closed += (s, args) => this.Close();

            //this.Close();
            //var a = MessageBox.Show(text1, "Проверка данных", MessageBoxButtons.YesNoCancel);

            Config config = new Config();
            config.OrganisationName = textBox1.Text;
            config.ClassNumber = textBox2.Text;
            config.ClassLetter = textBox3.Text;
            config.ClassroomTeacherName = textBox4.Text;
            config.ClassroomTeacherEmail = textBox5.Text;
            config.ClassroomTeacherPhone = textBox6.Text;
            config.PsychologistName = textBox7.Text;
            config.PsychologistEmail = textBox8.Text;
            config.PsychologistPhone = textBox9.Text;
            config.AcademicYearStartReporting = textBox10.Text;
            config.ReportingStartQuarterNumber = textBox11.Text;
            config.ReportingAcademicYear = textBox12.Text;
            config.ReportingQuarterNumber = textBox13.Text;
            // Здесь надо пройти валидацию
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
                a();
            }
            else if (selectedIndex == 1)
            {
                a();
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

        private void a()
        {
            buttonNext.Visible = tabControl1.SelectedIndex != tabControl1.TabPages.Count - 1;
            buttonPrev.Visible = tabControl1.SelectedIndex != 0;
            buttonSave.Visible = tabControl1.SelectedIndex != 0;
        }
    }
}
