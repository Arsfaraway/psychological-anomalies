using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Config = Reader1.Models.Configuration.Configuration;

namespace Reader1.Forms
{
    public partial class CheckingInformation : Form
    {
        public string field { get; set; }

        public CheckingInformation(string name, string field)
        {
            InitializeComponent();
            textBox1.Text = field;
            label1.Text = "Проверьте корретность следующего поля: " + name;
            this.field = field;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            field = textBox1.Text;
            this.Close();
        }
    }
}
