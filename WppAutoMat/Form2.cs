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
using System.Runtime.InteropServices;
using OpenQA.Selenium.Support.UI;

namespace WppAutoMat
{
    

    public partial class Form2 : Form
    {       
     
        
        IWebDriver driver = Form1.driver;        
        OpenFileDialog OFD2 = new OpenFileDialog();
        OpenFileDialog OFD = new OpenFileDialog();

        string ruta = "";
        string rutaIma = "";
        string Nima = "";

        public Form2()
        {

            InitializeComponent();

        }        

    private void button2_Click(object sender, EventArgs e)
        {
            OFD.Filter = "Excel |*.xls;*.xlsx;*.xlsm";
            OFD.InitialDirectory = "Desktop";

            if (OFD.ShowDialog() == DialogResult.OK)
            {
                ruta = OFD.FileName;

            }
            else { MessageBox.Show("No se seleccionó nada.", "Archivo"); return; }


            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            ruta +
                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * from [enlaces$]", con);
            OleDbCommand second = new OleDbCommand("Select * from[Hoja1$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            sda.Fill(data);
            dataGridView1.DataSource = data;

            OleDbDataAdapter sde = new OleDbDataAdapter(second);
            DataTable info = new DataTable();
            sde.Fill(info);
            dataGridView2.DataSource = info;


            button2.Visible = false;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    textBox1.Text += dataGridView1.Rows[i].Cells[j].Value.ToString() + " ";
                }

            }

            button2.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try {

                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Por favor ingresa toda la información necesaria"); return;
                } else if (textBox2.Text == "")
                {
                    MessageBox.Show("Por favor ingresa toda la información necesaria"); return;
                }


                else if (rutaIma == null)
                {
                    SinImagen();

                } else if (rutaIma != null && checkBox1.Checked && textBox2.Text.Length > 1024) {

                    MessageBox.Show("El mensaje es demasiado largo, intenta sin el texto en la imagen."); return;
                }
                else if (rutaIma != null && checkBox1.Checked)
                {

                    ConTextImagen();
                }
                else if(rutaIma != null && !checkBox1.Checked)
                {
                    TextMasImagen();
                }

            }
            catch(Exception s)
            {
                MessageBox.Show(s.Message.ToString());
                this.Show();
            }

        }

        private void Form2_Close(object sender, EventArgs e)
        {
            Form1 f2 = new Form1();
            f2.Show();
            this.Hide();
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {
            OFD2.Filter = "Imágen |*.JPG;*.PNG;*.GIF";
            OFD2.InitialDirectory = "Desktop";

            if (OFD2.ShowDialog() == DialogResult.OK)
            {
                rutaIma = OFD2.FileName;
                Nima = System.IO.Path.GetFileName(OFD2.FileName);


            }
            else { MessageBox.Show("No se seleccionó nada.", "Archivo"); return; }
            label1.Text = Nima;
            pictureBox1.Image = System.Drawing.Image.FromFile(rutaIma);
            checkBox1.Visible = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {

            Form1 f2 = new Form1();
            f2.Show();
            this.Hide();

        }

        private void SinImagen()
        {
            button2.Enabled = true;
            driver.FindElement(By.CssSelector("._2MSJr")).Click();
            this.Visible = false;
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/label/input")).SendKeys("89270287" + OpenQA.Selenium.Keys.Enter);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            driver.FindElement(By.CssSelector("._2S1VP.copyable-text.selectable-text")).SendKeys(textBox1.Text + OpenQA.Selenium.Keys.Enter);

            IList<IWebElement> link = driver.FindElements(By.CssSelector("a.selectable-text.invisible-space.copyable-text"));
            int n = link.Count;

            for (int x = 0; x < n; x++)
            {
                IList<IWebElement> enviar = driver.FindElements(By.CssSelector("a.selectable-text.invisible-space.copyable-text"));
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                enviar[x].Click();
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                driver.FindElement(By.CssSelector("._2S1VP.copyable-text.selectable-text")).SendKeys(textBox2.Text + OpenQA.Selenium.Keys.Enter);
             //chat recipiente de los links                                                                         //cambiar por tu número o chat
                driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/label/input")).SendKeys("89270287" + OpenQA.Selenium.Keys.Enter);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            }

            //vacía el chat para evitar errores
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/header/div[3]/div/div[3]")).Click();
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/header/div[3]/div/div[3]/span/div/ul/li[4]/div")).Click();
            driver.FindElement(By.CssSelector("div._1WZqU.PNlAR")).Click();
            this.Visible = true;
        }


        private void ConTextImagen()
        {

            driver.FindElement(By.CssSelector("._2MSJr")).Click();
            this.Visible = false;
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/label/input")).SendKeys("89270287" + OpenQA.Selenium.Keys.Enter);          
            driver.FindElement(By.CssSelector("._2S1VP.copyable-text.selectable-text")).SendKeys(textBox1.Text + OpenQA.Selenium.Keys.Enter);

            IList<IWebElement> link = driver.FindElements(By.CssSelector("a.selectable-text.invisible-space.copyable-text"));
            int n = link.Count;

            for (int x = 0; x < n; x++)
            {               

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                IList<IWebElement> enviar = driver.FindElements(By.CssSelector("a.selectable-text.invisible-space.copyable-text"));                            
                enviar[x].Click();

                System.Threading.Thread.Sleep(2000);

                driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/header/div[3]/div/div[2]")).Click();
                //js para hacer visible el input file para poder enviar a la variable la ruta de la imagen                           
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                driver.FindElement(By.CssSelector("li._10anr.vidHz._3asN5"));

                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
               js.ExecuteScript(
                 @"document.querySelector('input[type=""file""]').style.display = 'block';"

                    );



                System.Threading.Thread.Sleep(2000);
                //envía la imagen
                driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/header/div[3]/div/div[2]/span/div/div/ul/li[1]/input")).SendKeys(rutaIma);
                //escribe el texto dentro del textbox en la imagen y luego envía           
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                driver.FindElement(By.CssSelector("img._1a4Ru"));

                driver.FindElement(By.CssSelector("._2S1VP.copyable-text.selectable-text")).SendKeys(textBox2.Text);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                driver.FindElement(By.CssSelector("div._3hV1n.yavlE")).Click();

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                driver.FindElement(By.CssSelector("._2MSJr")).Click();              
                driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/label/input")).SendKeys("Jairo" + OpenQA.Selenium.Keys.Enter);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            }
            //vacía el chat para evitar errores
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/header/div[3]/div/div[3]")).Click();
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/header/div[3]/div/div[3]/span/div/ul/li[4]/div")).Click();
            driver.FindElement(By.CssSelector("div._1WZqU.PNlAR")).Click();
            this.Visible = true;
        }

        private void TextMasImagen()
        {

            driver.FindElement(By.CssSelector("._2MSJr")).Click();
            this.Visible = false;
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/label/input")).SendKeys("89270287" + OpenQA.Selenium.Keys.Enter);
            driver.FindElement(By.CssSelector("._2S1VP.copyable-text.selectable-text")).SendKeys(textBox1.Text + OpenQA.Selenium.Keys.Enter);

            IList<IWebElement> link = driver.FindElements(By.CssSelector("a.selectable-text.invisible-space.copyable-text"));
            int n = link.Count;

            for (int x = 0; x < n; x++)
            {



                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                IList<IWebElement> enviar = driver.FindElements(By.CssSelector("a.selectable-text.invisible-space.copyable-text"));
                enviar[x].Click();

                System.Threading.Thread.Sleep(2000);

                driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/header/div[3]/div/div[2]")).Click();
                //js para hacer visible el input file para poder enviar a la variable la ruta de la imagen                           
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                driver.FindElement(By.CssSelector("li._10anr.vidHz._3asN5"));

                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript(
                  @"document.querySelector('input[type=""file""]').style.display = 'block';"

                     );

                //envia el mensaje primero luego la imagen
                driver.FindElement(By.CssSelector("._2S1VP.copyable-text.selectable-text")).SendKeys(textBox2.Text + OpenQA.Selenium.Keys.Enter);
                //espera necesaria para esperar que el inputfile sea alcanzable por el driver
                System.Threading.Thread.Sleep(2000);
                //envía la imagen
                driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/header/div[3]/div/div[2]/span/div/div/ul/li[1]/input")).SendKeys(rutaIma);
                //escribe el texto dentro del textbox en la imagen y luego envía           
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                driver.FindElement(By.CssSelector("img._1a4Ru"));                
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                driver.FindElement(By.CssSelector("div._3hV1n.yavlE")).Click();

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                driver.FindElement(By.CssSelector("._2MSJr")).Click();
                driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/label/input")).SendKeys("Jairo" + OpenQA.Selenium.Keys.Enter);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            }
            //vacía el chat para evitar errores
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/header/div[3]/div/div[3]")).Click();
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/header/div[3]/div/div[3]/span/div/ul/li[4]/div")).Click();
            driver.FindElement(By.CssSelector("div._1WZqU.PNlAR")).Click();
            this.Visible = true;
        }
    }
}


 

