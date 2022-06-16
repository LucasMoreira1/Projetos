using System;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace Ferramentas_Scheduler
{
    public partial class FormPrincipal : Form
    {
        public FormPrincipal()
        {
            InitializeComponent();
        }

        private void btnBaterPonto_Click(object sender, EventArgs e)
        {
            var ahgora = "https://www.ahgora.com.br/novabatidaonline/?defaultDevice=a774091";

            var driver = new ChromeDriver();

            driver.Navigate().GoToUrl(ahgora);

            
        }
    }
}
