using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

using SAPFEWSELib;

namespace Ferramentas_Scheduler
{
    public partial class FormPrincipal : Form
    {
        public class SAPActive
        {
            public static GuiApplication SapGuiApp { get; set; }
            public static GuiConnection SapConnection { get; set; }
            public static GuiSession SapSession { get; set; }

            public static void OpenSap(string env)
            {
                SAPActive.SapGuiApp = new GuiApplication();

                string connectString = null;
                if (env.ToUpper().Equals("DEFAULT"))
                {
                    connectString = "1.0 Test ERP (DEFAULT)";
                }
                else
                {
                    connectString = env;
                }
                SAPActive.SapConnection = SAPActive.SapGuiApp.OpenConnection(connectString, Sync: true); //creates connection
                SAPActive.SapSession = (GuiSession)SAPActive.SapConnection.Sessions.Item(0); //creates the Gui session off the connection you made
            }
            public static void loginSap(string myclient, string mylogin, string mypass, string mylang)
            {
                GuiTextField client = (GuiTextField)SAPActive.SapSession.ActiveWindow.FindByName("RSYST-MANDT", "GuiTextField");
                GuiTextField login = (GuiTextField)SAPActive.SapSession.ActiveWindow.FindByName("RSYST-BNAME", "GuiTextField");
                GuiTextField pass = (GuiTextField)SAPActive.SapSession.ActiveWindow.FindByName("RSYST-BCODE", "GuiPasswordField");
                GuiTextField language = (GuiTextField)SAPActive.SapSession.ActiveWindow.FindByName("RSYST-LANGU", "GuiTextField");

                client.SetFocus();
                client.Text = myclient;
                login.SetFocus();
                login.Text = mylogin;
                pass.SetFocus();
                pass.Text = mypass;
                language.SetFocus();
                language.Text = mylang;

                //Press the green checkmark button which is about the same as the enter key 
                GuiButton btn = (GuiButton)SapSession.FindById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[0]");
                btn.SetFocus();
                btn.Press();
            }

        }
        public FormPrincipal()
        {
            InitializeComponent();
        }

        private void btnBaterPonto_Click(object sender, EventArgs e)
        {
            //var ahgora = "https://www.ahgora.com.br/novabatidaonline/?defaultDevice=a774091";
            var ahgora = "https://www.google.com";

            var driver = new ChromeDriver();

            driver.Navigate().GoToUrl(ahgora);

            
        }

        private void btnArchiveDiario_Click(object sender, EventArgs e)
        {
            AbrirSAP();
        }
        public void AbrirSAP(string[] args)
        {
            SAPActive.OpenSap("My connection string");
            SAPActive.loginSap("10", "lucas", "lmsferr", "PT");
            SAPActive.SapSession.StartTransaction("SM37");
        }
    }
}
