using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MurbongMacro
{
    public partial class Form2 : Form
    {
        public delegate void FormSendDataHandler(string sendstring);

        public event FormSendDataHandler FormSendEvent;

        public void SetTextBox(string Key, string Interval, string Coordinate, string Action, string Code)
        {
            txtKey.Text = Key;
            txtInt.Text = Interval;
            txtCoord.Text = Coordinate;
            txtAct.Text = Action;
            txtCode.Text = Code;
            txtKey.ReadOnly = true;
            txtCode.ReadOnly = true;
            txtAct.ReadOnly = true;

        }

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            FormSendEvent(txtKey.Text + "/" + txtInt.Text + "/" + txtCoord.Text + "/" + txtAct.Text + "/" + txtCode.Text);
            this.Close();
        }
    }
}
