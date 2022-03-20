using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjetoC
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
                comboBox1.Items.Clear();
                comboBox1.Items.Add("Teste1");
                comboBox1.Items.Add("Teste2");
                comboBox1.Items.Add("Teste3");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = comboBox1.Text;

            listView1.Columns.Add("teste");
            listView1.Columns.Add("teste2");
       
        }
    }
}
