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
            SapROTWr.CSapROTWrapper sapROTWrapper = new SapROTWr.CSapROTWrapper();
            object SapGuilRot = sapROTWrapper.GetROTEntry("SAPGUI");
            object engine = SapGuilRot.GetType().InvokeMember("GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, SapGuilRot, null);
            GuiConnection connection = (engine as GuiApplication).OpenConnection("EBP - Produção Testes Integrados New Edwards");
            GuiSession session = connection.Children.ElementAt(0) as GuiSession;

            GuiTextField login = (GuiTextField)session.ActiveWindow.FindByName("RSYST-BNAME", "GuiTextField");
            GuiTextField pass = (GuiTextField)session.ActiveWindow.FindByName("RSYST-BCODE", "GuiPasswordField");
            GuiButton btn;
            btn = (GuiButton)session.FindById("wnd[0]/tbar[0]/btn[0]");
            
            login.SetFocus(); login.Text = "schedule2";
            pass.SetFocus(); pass.Text = "sCHeduLe2##";
            btn.SetFocus(); btn.Press();

            session.StartTransaction("Sara");

            //GuiTextField sara = (GuiTextField)session.ActiveWindow.FindById("wnd[0]/tbar[0]/okcd", "sara");
            //sara.SetFocus(); sara.Text = "SARA";
            btn.SetFocus(); btn.Press();
            GuiTextField idoc = (GuiTextField)session.ActiveWindow.FindById("wnd[0]/usr/ctxtARCH_TXT-OBJECT", "idoc");
            idoc.SetFocus(); idoc.Text = "idoc";
            GuiButton btnDelete;
            btnDelete = (GuiButton)session.FindById("wnd[0]/usr/btnDELETE");
            btnDelete.SetFocus(); btnDelete.Press();

            GuiButton btn1;
            btn1 = (GuiButton)session.FindById("wnd[0]/usr/btn%_AUTOTEXT009");
            btn1.SetFocus(); btn1.Press();

            GuiButton btn2;
            btn2 = (GuiButton)session.FindById("wnd[1]/tbar[0]/btn[8]");
            btn2.SetFocus(); btn2.Press();

            GuiButton btn9;
            btn9 = (GuiButton)session.FindById("wnd[1]/tbar[0]/btn[0]");
            btn9.SetFocus(); btn9.Press();


            GuiButton btn3;
            btn3 = (GuiButton)session.FindById("wnd[0]/usr/btn%_AUTOTEXT007");
            btn3.SetFocus(); btn3.Press();

            GuiButton btn4;
            btn4 = (GuiButton)session.FindById("wnd[1]/usr/btnSOFORT_PUSH");
            btn4.SetFocus(); btn4.Press();

            GuiButton btn5;
            btn5 = (GuiButton)session.FindById("wnd[1]/tbar[0]/btn[11]");
            btn5.SetFocus(); btn5.Press();

            GuiButton btn6;
            btn6 = (GuiButton)session.FindById("wnd[0]/usr/btn%_AUTOTEXT008");
            btn6.SetFocus(); btn6.Press();

            GuiTextField local = (GuiTextField)session.ActiveWindow.FindById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST", "local");
            local.SetFocus(); local.Text = "LOCAL";

            GuiButton btn7;
            btn7 = (GuiButton)session.FindById("wnd[1]/tbar[0]/btn[13]");
            btn7.SetFocus(); btn7.Press();

            GuiButton btn8;
            btn8 = (GuiButton)session.FindById("wnd[2]/tbar[0]/btn[0]");
            btn8.SetFocus(); btn8.Press();

            GuiButton btn10;
            btn10 = (GuiButton)session.FindById("wnd[0]/tbar[1]/btn[8]");
            btn10.SetFocus(); btn10.Press();
            //GuiButton btn9;
            //btn9 = (GuiButton)session.FindById("wnd[0]/usr/btnDELETE");


        }
    }
}
