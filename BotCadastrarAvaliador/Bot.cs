using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BotCadastrarAvaliador
{
    public class Bot
    {
        private IWebDriver? driver;

        public Bot()
        {
            FirefoxOptions options = new();
            options.AddArguments("--private", "--start-maximized", "--safe-mode", "--disable-component-update", "--no-default-browser-check", "--disable-gpu", "--ignore-certificate-errors");

            FirefoxDriverService service = FirefoxDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            driver = new FirefoxDriver(service, options);
        }

        public bool OpenPage(string url)
        {
            try
            {
                driver.Navigate().GoToUrl(url);
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                return false;
            }
        }

        public int FindChild(string xpath_row, string xpath_child, string value)
        {
            try
            {
                if (!WaitElement(By.XPath(xpath_row))) throw new Exception("O elemento não existe.");

                var row = driver.FindElements(By.XPath(xpath_row));

                for (int i=0; i < row.Count; i++)
                {
                    MessageBox.Show(i + ": " + row[i].FindElement(By.XPath(xpath_child)).Text);
                    if (row[i].FindElement(By.XPath(xpath_child)).Text.Contains(value)) return i + 1;
                }
            }
            catch (Exception err)
            {
                MessageBox.Show($"FindChild: {err.Message}");
            }

            return 0;
        }

        public bool FindAndClick(string xpath_row, string xpath_child, string value)
        {
            return false;
        }

        public bool WaitElement(By by_element, int tempo_espera = 30)
        {
            if (driver != null)
            {
                for (int i = 0; i < tempo_espera * 4; i++)
                {
                    try
                    {
                        if (driver.FindElement(by_element).Displayed) return true;
                    }
                    catch (Exception) { }
                    Thread.Sleep(250);
                }
            }

            return false;
        }

        public bool Click(By by_element)
        {
            try
            {
                if (!WaitElement(by_element, 10)) throw new Exception("Elemento não existe ou não foi carregado.");

                driver.FindElement(by_element).Click();
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                return false;
            }
        }

        public bool SendText(By by_element, object text)
        {
            try
            {
                if (!WaitElement(by_element, 10)) throw new Exception("Elemento não existe ou não foi carregado.");

                driver.FindElement(by_element).SendKeys(text.ToString());
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                return false;
            }
        }

        public void Close()
        {
            if (driver != null)
            {
                try
                {
                    driver.Close();
                    driver.Quit();
                }
                catch (Exception) { }
            }
        }

    }
}
