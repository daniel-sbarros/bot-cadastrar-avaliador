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
        public string Url { get { return driver.Url ?? null; } }

        public Bot()
        {
            FirefoxOptions options = new();
            options.AddArguments("--private", "--start-maximized", "--safe-mode", "--disable-component-update", "--no-default-browser-check", "--disable-gpu", "--ignore-certificate-errors");

            FirefoxDriverService service = FirefoxDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            driver = new FirefoxDriver(service, options);
        }

        public bool isChecked(By by_element)
        {
            return driver.FindElement(by_element).Selected;
        }

        public bool OpenPage(string url)
        {
            try
            {
                driver!.Navigate().GoToUrl(url);
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                return false;
            }
        }

        private string ReplaceEspecialChars(string str)
        {
            string[] acentos = new string[] { "ç", "Ç", "á", "é", "í", "ó", "ú", "ý", "Á", "É", "Í", "Ó", "Ú", "Ý", "à", "è", "ì", "ò", "ù", "À", "È", "Ì", "Ò", "Ù", "ã", "õ", "ñ", "ä", "ë", "ï", "ö", "ü", "ÿ", "Ä", "Ë", "Ï", "Ö", "Ü", "Ã", "Õ", "Ñ", "â", "ê", "î", "ô", "û", "Â", "Ê", "Î", "Ô", "Û" };
            string[] semAcento = new string[] { "c", "C", "a", "e", "i", "o", "u", "y", "A", "E", "I", "O", "U", "Y", "a", "e", "i", "o", "u", "A", "E", "I", "O", "U", "a", "o", "n", "a", "e", "i", "o", "u", "y", "A", "E", "I", "O", "U", "A", "O", "N", "a", "e", "i", "o", "u", "A", "E", "I", "O", "U" };

            for (int i = 0; i < acentos.Length; i++)
                str = str.Replace(acentos[i], semAcento[i]);

            return str.Trim();
        }

        public int FindChild(string xpath_row, string xpath_child, string value)
        {
            try
            {
                if (!WaitElement(By.XPath(xpath_row))) throw new Exception("O elemento não existe.");

                var row = driver!.FindElements(By.XPath(xpath_row));

                for (int i=0; i < row.Count; i++)
                {
                    if (ReplaceEspecialChars(row[i].FindElement(By.XPath(xpath_child)).Text).ToUpper().Contains(ReplaceEspecialChars(value).ToUpper())) return i + 1;
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

                driver!.FindElement(by_element).Click();
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

                driver!.FindElement(by_element).SendKeys(text.ToString());
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
