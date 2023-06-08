using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

namespace Proyecto_Zendesk_Beta
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            radioButtonWppLua.Checked = true;
            radioButtonEdge.Checked = true;
            radioButtonSimultaneity3.Checked = true;
            System.Console.SetOut(System.IO.TextWriter.Null);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int simultaneidad;
            int timeSleep = 1000;
            bool refresh = true;
            string driverPath;
            IWebDriver driver;
            EdgeOptions options;
            WebDriverWait wait;
            Actions actions;

            // Configuracion de la simultaneidad
            simultaneidad = radioButtonSimultaneity3.Checked ? 3 : 4;

            // Configuracion del navegador
            if (radioButtonEdge.Checked)
            {
                driverPath = driverPath = @"\\Co0000fs0001\planeacion$\01_LATAM\02_REPORTING\07-Gestion en tiempo real\Fraily\Proyectos\Zendesk\driver\edge";
                options = new EdgeOptions();
                driver = new EdgeDriver(driverPath, options);
                actions = new Actions(driver);
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            }
            else
            {
                driverPath = driverPath = @"\\Co0000fs0001\planeacion$\01_LATAM\02_REPORTING\07-Gestion en tiempo real\Fraily\Proyectos\Zendesk\driver\chrome";
                options = new EdgeOptions();
                driver = new EdgeDriver(driverPath, options);
                actions = new Actions(driver);
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            }

            try
            {
                // Navegamos a la pagina web de Zendesk
                Navegar(driver);

                // Logueo del usuario
                Login(driver, wait);

                // Configuracion de vistas
                if (radioButtonWppVts.Checked)
                {
                    // //div[@class='sc-1ui74oj-0 bHAYWD'][normalize-space()='Whatsapp Ventas SSC - en Cola']
                    IWebElement ulElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[normalize-space()='Whatsapp Ventas SSC - en Cola']")));
                    ulElement.Click();
                }

                //// Filtro descendente
                //wait.Until(elementExist => driver.FindElements(By.CssSelector("button[data-garden-id='tables.sortable']")));
                //Thread.Sleep(timeSleep);
                //var desc = driver.FindElements(By.CssSelector("button[data-garden-id='tables.sortable']"));
                //desc[5].Click();

                // Abrimos y navegamos a una segunda pestaña
                Thread.Sleep(timeSleep);
                ((IJavaScriptExecutor)driver).ExecuteScript("window.open();");
                driver.SwitchTo().Window(driver.WindowHandles.Last());

                // Navegamos a la pagina web de Zendesk
                Navegar(driver);

                // Logueo del usuario
                Login(driver, wait);

                // Configuracion de vistas
                if (radioButtonWppLua.Checked)
                {
                    IWebElement wppLuaElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[normalize-space()='Asignacion LUA']")));
                    wppLuaElement.Click();
                }
                else
                {
                    IWebElement wppVtsElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[normalize-space()='Asignacion VTS']")));
                    wppVtsElement.Click();
                }

                // Ciclo de trabajo
                while (true)
                {
                    // Lista de agentes
                    List<Agente> agentesList = new List<Agente>();

                    // Listamos todos los agentes de la vista
                    GetList(wait, refresh, agentesList);

                    if (!refresh)
                    {
                        continue;
                    }

                    // Tab (Primera Pestaña)
                    Thread.Sleep(timeSleep);
                    driver.SwitchTo().Window(driver.WindowHandles.First());

                    // Asignamos casos
                    Thread.Sleep(timeSleep);
                    Asignar(driver, wait, timeSleep, agentesList, actions, simultaneidad);

                    // Tab (Ultima Pestaña)
                    Thread.Sleep(timeSleep);
                    driver.SwitchTo().Window(driver.WindowHandles.Last());

                    // Click en (Siguiente Pagina)
                    Thread.Sleep(timeSleep);
                    IWebElement nextElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("nav > button:nth-child(3) > span:nth-child(1)")));
                    nextElement.Click();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                //driver.Quit();
            }
        }

        private void Navegar(IWebDriver driver)
        {
            driver.Navigate().GoToUrl("http://casounico.app.lan.com/CASOUNICO-1.0");
        }

        private void Login(IWebDriver driver, WebDriverWait wait)
        {
            // Login del usuario
            IWebElement userElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("input[id='loginForm:userName']")));
            userElement.SendKeys(textBox_user.Text);

            IWebElement passwordElement = driver.FindElement(By.CssSelector("input[id='loginForm:userPassword']"));
            passwordElement.SendKeys(textBox_password.Text);
            passwordElement.SendKeys(OpenQA.Selenium.Keys.Enter);

            // Vistas de Zendesk
            IWebElement vistaElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("button[title='Vistas']")));
            vistaElement.Click();
        }

        private void GetList(WebDriverWait wait, bool refresh, List<Agente> agentesList)
        {
            // Listamos todos los agentes de la vista
            var tdElements = wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.CssSelector("table > tbody > tr > td[data-garden-id='tables.cell'][colspan='9']:nth-child(1)")));
            foreach (var tdElement in tdElements)
            {
                if (tdElement.Text == "Agente asignado: -")
                {
                    // Click en (Primera Pagina)
                    IWebElement firstElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("nav > nav > button > span")));
                    firstElement.Click();

                    refresh = false;
                    break;
                }

                var trElements = tdElement.FindElements(By.XPath("./../following-sibling::tr"));
                int tkt = trElements.Count(tr => tr.GetAttribute("data-garden-id") == "tables.row");

                var agente = new Agente
                {
                    Ticket = tkt,
                    Name = tdElement.Text.Replace("Agente asignado: ", "")
                };
                agentesList.Add(agente);
            }
        }

        private void Asignar(IWebDriver driver, WebDriverWait wait, int timeSleep, List<Agente> agentesList,
            Actions actions, int simultaneidad)
        {
            foreach (var agente in agentesList)
            {
                // Validacion (Agentes con simultaneidad)
                if (agente.Ticket >= simultaneidad)
                {
                    continue;
                }

                // Asignamos casos (Solo si hay casos presentes)
                var labelElements = wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.CssSelector("table > tbody:nth-child(2) > tr > td > div > label:nth-child(2)")));
                while (labelElements.Count > 0)
                {
                    int casos = Math.Min(simultaneidad - agente.Ticket, labelElements.Count);
                    for (int i = 0; i <= casos; i++)
                    {
                        labelElements[i].Click();
                        Thread.Sleep(timeSleep);
                    }

                    // Seleccionamos el boton (Editar)
                    IWebElement editarElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("button[data-test-id='bulk-actions-edit-button']")));
                    editarElement.Click();

                    // Escribimos el nombre del agente a asignar
                    wait.Until(elementExist => driver.FindElement(By.CssSelector("div[id='mn_16']")));
                    //Thread.Sleep(timeSleep);
                    var aggentAsigElement = driver.FindElement(By.CssSelector("div[id='mn_16']"));
                    aggentAsigElement.Click();

                    //Thread.Sleep(timeSleep);
                    var inputElement = aggentAsigElement.FindElement(By.CssSelector("input"));
                    actions.Click(inputElement);
                    foreach (char c in agente.Name)
                    {
                        actions.SendKeys(c.ToString()).Perform();
                    }
                    //Thread.Sleep(timeSleep);
                    //actions.SendKeys(agente.Name).Build().Perform();

                    // Seleccionamos el agente a asignar
                    Thread.Sleep(timeSleep);
                    var listaOrdenada = driver.FindElements(By.CssSelector("html>body>div:nth-of-type(7)>div:nth-of-type(2)>ul>li"));
                    foreach (var item in listaOrdenada)
                    {
                        if (radioButtonWppLua.Checked && item.Text == "WhatsApp SSC -AMC/" + agente.Name)
                        {
                            item.Click();
                        }

                        if (radioButtonWppVts.Checked && item.Text == "WhatsApp Ventas SSC/" + agente.Name)
                        {
                            item.Click();
                        }
                    }

                    // Presionamos el boton (Enviar)
                    IWebElement enviarElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("div[class='sc-1hg3so9-0 bhZefU']")));
                    enviarElement.Click();

                    // Esperamos antes de asignar nuevo caso
                    while (driver.FindElements(By.CssSelector("table:nth-of-type(1) tr:nth-of-type(2) td:nth-of-type(1) div > svg:nth-of-type(1)")).Count != 0)
                    {
                        Thread.Sleep(timeSleep);
                    }

                    //Thread.Sleep(timeSleep);
                    break;
                }

                // Refrezcamos la pagina (Solo si no hay casos presentes)
                while (driver.FindElements(By.CssSelector("table > tbody:nth-child(2) > tr > td > div > label:nth-child(2)")).Count < 1)
                {
                    driver.FindElement(By.CssSelector("button[data-test-id='views_views-list_header-refresh']")).Click();
                    Thread.Sleep(timeSleep);
                }
            }
        }
    }

    public class Agente
    {
        public string Name { get; set; }
        public int Ticket { get; set; }
    }
}
