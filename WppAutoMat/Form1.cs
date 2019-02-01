using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OpenQA.Selenium.Support.UI;
using System.Net;
using HtmlAgilityPack;

namespace WppAutoMat
{

    public partial class Form1 : Form
    {
        OpenFileDialog OFD = new OpenFileDialog();
        string valor = "";
        string ruta = "";
        string UsN = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

       public static IWebDriver driver = new FirefoxDriver();


        public Form1()
        {
            InitializeComponent();
          
        }

        private void Licencia()
        {
            string html = string.Empty;
            string url = @"https://docs.google.com/spreadsheets/d/e/2PACX-1vTRg5XFd_QP1k8MFVc6LF7nAPDo3qQSvynPcLs-ojuuOV6eYDUQsecWRNPgzNbRMuWrAvKReKoOC6PY/pubhtml";
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();


            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.ContentType = ("text/xml");

            request.AutomaticDecompression = DecompressionMethods.GZip;

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream, Encoding.GetEncoding(response.CharacterSet)))
            {
                html = reader.ReadToEnd();
            }

        }

        private bool IsElementPresent(By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try { 

            string contacto = textBox2.Text;
            string mensaje = textBox1.Text;         

            if(textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "") { MessageBox.Show("Por favor ingresa toda la información necesaria"); return;}


            driver.FindElement(By.CssSelector("._2MSJr")).Click();
            this.Visible = false;
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/label/input")).SendKeys(contacto + OpenQA.Selenium.Keys.Enter);

            if (IsElementPresent(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[3]/div/div/span")))
            {
               
                this.Visible = true;
                return;
            }
             
                IWebElement nombre = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/header/div[2]/div/div/span"));
                if (!nombre.Text.Contains(contacto))
                {

                    MessageBox.Show("No se encontró el contacto: " + contacto);
                    this.Visible = true;
                    return;
                   

                }                     


            for (int i = 1; i <= int.Parse(textBox3.Text); i++) {
                driver.FindElement(By.CssSelector("._2S1VP.copyable-text.selectable-text")).SendKeys(mensaje + OpenQA.Selenium.Keys.Enter);
            }
            this.Visible = true;
            }
            catch(Exception)
            {
                MessageBox.Show("Ha ocurrido un problema, asegúrate que estés haciendo todos los procesos correctos.");
                this.Visible = true;
            }


        }

        private void Form1_Load(object sender, EventArgs e)
        {

            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
           

            string DirTxt = @"C:\Licencia";
            string lic = "VzQ0dTcwbTR0MTU0YzEwTg ==";
            string linea = "";
            Form Fr1 = new Form1();

            if (System.IO.File.Exists(DirTxt))
            {

                using (StreamReader file = new StreamReader(DirTxt))


                {

                    linea = file.ReadLine();
                    if (!linea.Contains(lic))
                    {
                       
                        MessageBox.Show("No se encontró licencia válida.", "Licencia no registrada");
                        Form FLicencia = new Form3();
                        FLicencia.Show();
                        Fr1.Hide();
                    }
                    else
                    {
                        button1.Enabled = true;
                        button2.Enabled = true;
                        button3.Enabled = true;
                        button4.Enabled = true;
                        button5.Enabled = true;
                    }

                }
            }
            else
            {

                using (StreamWriter Licencia = File.AppendText(DirTxt))
                {

                }

                        MessageBox.Show("No se encontró licencia válida.", "Licencia no registrada");
                        Form FLicencia = new Form3();
                        FLicencia.Show();
                        Fr1.Hide();              

            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {


            OFD.Filter = "Excel |*.xls;*.xlsx;*.xlsm";
            OFD.InitialDirectory = "Desktop";

            if (OFD.ShowDialog() == DialogResult.OK){
                ruta = OFD.FileName;
            }
            else { MessageBox.Show("No se seleccionó nada.", "Archivo"); return; }



            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            ruta +
                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * from [Hoja1$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            sda.Fill(data);
            dataGridView1.DataSource = data;
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string mensaje = textBox4.Text;

            if(mensaje == "") {
                MessageBox.Show("Por favor ingresa toda la información necesaria"); return;
            }else if(dataGridView1.Rows.Count == 0){
                MessageBox.Show("Por favor ingresa toda la información necesaria"); return;
            }
            try
            {
                driver.FindElement(By.CssSelector("._2MSJr")).Click();
                this.Visible = false;

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        valor = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        string contacto = valor;
                        driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/label/input")).SendKeys(valor + OpenQA.Selenium.Keys.Enter);
                        if (!IsElementPresent(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[3]/div/div/span")))
                        {

                            driver.FindElement(By.CssSelector("._2S1VP.copyable-text.selectable-text")).SendKeys(mensaje + OpenQA.Selenium.Keys.Enter);
                        }


                    }

                }
                this.Visible = true;
            }
            catch (Exception)
            {
                MessageBox.Show("Ocurrión un problema, inténtalo de nuevo", "Error");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form2 f2 = new Form2();
            f2.Show();
            
        }

        private void button5_Click(object sender, EventArgs e)
        {

            try
            {
                MessageBox.Show("Abriendo Whatsapp, inicia sesión", "WhatsappAutoMat");
                button5.Visible = false;
                driver.Url = "https://web.whatsapp.com";
                this.Visible = false;
            }
            catch (Exception xe)
            {
                MessageBox.Show(xe.ToString());
                this.Visible = true;
            }


           
        }
        private void Form1_Close(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void ConAgility()
        {

            string recibe;
            var clave = "VzQ0dTcwbTR0MTU0YzEwTg==";
            var html = @"https://docs.google.com/spreadsheets/d/e/2PACX-1vTRg5XFd_QP1k8MFVc6LF7nAPDo3qQSvynPcLs-ojuuOV6eYDUQsecWRNPgzNbRMuWrAvKReKoOC6PY/pubhtml";

            HtmlWeb web = new HtmlWeb();
            var htmlDoc = web.Load(html);
            var htmlNodes = htmlDoc.DocumentNode.SelectNodes("/html/body/div[2]/div/div/table/tbody/tr/td[1]");

            foreach (var node in htmlNodes)
            {

                recibe = node.InnerText.ToString();

                if (recibe == clave)
                {

                    Console.WriteLine("Correcto");
                    return;
                }

            }

            Console.WriteLine("No se encontró licencia");
        }
    }

  }

 

