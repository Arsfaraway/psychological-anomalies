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
            var a = MessageBox.Show("Данные успешно получены!", "Проверка данных", MessageBoxButtons.YesNoCancel);

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
