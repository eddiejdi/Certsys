using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ribbons
{

    public class WebRobots
    {
        public WebRobots(string vUser, string vPass, string vProjeto, string vHI1, string vHI2, string vHI3, string vHF1, string vHF2, string vHF3)
        { // - Construtor
            User = vUser;
            Pass = vPass;
            Projeto = vProjeto;
            HI1 = vHI1;
            HI2 = vHI2;
            HI3 = vHI3;
            HF1 = vHF1;
            HF2 = vHF2;
            HF3 = vHF3;
        }

        public string User { get; set; }
        public string Pass { get; set; }
        public string Projeto { get; set; }
        public string HI1 { get; set; }
        public string HI2 { get; set; }
        public string HI3 { get; set; }
        public string HF1 { get; set; }
        public string HF2 { get; set; }
        public string HF3 { get; set; }

        public bool WebPonto()
        {
            IWebDriver browser = new ChromeDriver();
            //Firefox's proxy driver executable is in a folder already
            //  on the host system's PATH environment variable.
            //Maximizar
            browser.Manage().Window.Maximize();
            //Navegar para o Link
            browser.Navigate().GoToUrl("https://sccd.certsys.com.br/maximo/ui/?event=loadapp&value=activity&uniqueid=4805&uisessionid=929&csrftoken=9kkoh60l178mjha3dgk0vnebc5");
            //Usuário
            browser.FindElement(By.Name("j_username")).SendKeys(User);
            //Senha
            browser.FindElement(By.Name("j_password")).SendKeys(Pass);
            //Enter
            browser.FindElement(By.Name("j_password")).SendKeys(OpenQA.Selenium.Keys.Enter);
            //Click na Aba de lançameto de horas
            browser.FindElement(By.Id("mx298_anchor")).Click();


            //browser.Close();
            //browser.FindElement(By.Id("mx1537")).Click();
            return true;
        }
    }
}