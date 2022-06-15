using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;


namespace SCHEDULER
{
    public partial class FormInicial : Form
    {
        public FormInicial()
        {
            InitializeComponent();
        }
        private void FormInicial_Load(object sender, EventArgs e)
        {

        }

        private void btnArchiveDiario_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Archive diário em execução.", "Archive diário.",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnBaterPonto_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Batida de ponto em execução.", "Batida ponto.",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            string ahgora = "https://www.ahgora.com.br/novabatidaonline/?defaultDevice=a774091";
            var driver = new ChromeDriver();
            driver.Navigate().GoToUrl(ahgora);

            

            driver.Quit();
            
            
        }

        private void btnAlterarSenhaSAP_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Senhas SAP sendo alteradas.", "Senhas SAP.",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}