using OpenQA.Selenium;

namespace BotCadastrarAvaliador
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Bot bot = new();

            //bot.OpenPage("https://suap.ifma.edu.br/admin/pesquisa/comissaoeditalpesquisa/?ano=2022&descricao=143");
            bot.OpenPage("https://suap.ifma.edu.br/pesquisa/adicionar_comissao_por_area/187/");

            if (bot.WaitElement(By.Id("id_username")))
            {
                if(!bot.SendText(By.Id("id_username"), "2225318")) return;
                if(!bot.SendText(By.Id("id_password"), "difma+07")) return;
                Thread.Sleep(300);
                if(!bot.Click(By.XPath(@"//input[@value='Acessar']"))) return;

                if(!bot.WaitElement(By.ClassName("notifications"))) return;

                int row = bot.FindChild("//*[@id=\"bolsas_form\"]/table/tbody/tr", "td[2]", "Abel Batista de Oli");

                MessageBox.Show($"Row: {row}");

                if (row > 0)
                {
                    bot.Click(By.XPath($"//*[@id=\"bolsas_form\"]/table/tbody/tr[{row}]/td[1]/input"));
                    MessageBox.Show("Achei");
                }
                else
                {
                    MessageBox.Show("N�o Achei");
                }




                //driver.FindElement(By.Name("Salvar")).Click();
            }
            else
            {
            }

            //bot.Close();
        }
    }
}