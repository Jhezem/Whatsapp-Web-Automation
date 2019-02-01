using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace WppAutoMat
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            string DirTxt = @"C:\Licencia";

            using (StreamWriter Licencia = File.AppendText(DirTxt))
            {
                Licencia.WriteLine(textBox1.Text);
            }

            MessageBox.Show("El programa se cerrará e iniciará de nuevo");
            this.Hide();
            Form fr1 = new Form1();
            fr1.Show();
        }
    }
}
