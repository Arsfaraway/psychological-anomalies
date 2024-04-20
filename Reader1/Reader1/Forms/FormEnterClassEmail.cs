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
        private string validationCode;
        private bool isCodeValid = false;

        public bool IsCodeValid
        {
            get { return isCodeValid; }
        }
        public FormEnterClassEmail(string teacherEmail, string validCode)
        {
            InitializeComponent();
            validationCode = validCode;
            SetTextBoxTwoText(teacherEmail);
        }
        public void SetTextBoxTwoText(string teacherEmail)
        {
            label2.Text = teacherEmail;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void ButtonClickConfirmCode(object sender, EventArgs e)
        {
            if(validationCode == textBox4.Text.Trim())
            {
                isCodeValid = true;
            }
            this.Close();
        }
    }
}
