using OpenQA.Selenium;

namespace BotCadastrarAvaliador
{
    public partial class Form1 : Form
    {
        string? user, pass;
        Bot bot;
        MSExcel excel;
        List<Avaliador> internos, externos;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            LerCredenciais();

            excel = excel ?? new MSExcel("cadastro-de-avaliador-interno-externo.xlsx");
            internos = internos ?? excel.GetListValues(1, "Avaliador Interno");
            externos = externos ?? excel.GetListValues(2, "Avaliador Externo");

            cbxTipo.SelectedIndex = 0;
        }

        public void Logs(object texto)
        {
            StreamReader sr = new(Application.StartupPath + @"\logs.txt");
            string line = sr.ReadLine();
            System.Text.StringBuilder conteudo = new();

            while (line != null) 
            {
                conteudo.AppendLine(line);
                line = sr.ReadLine();
            }
            sr.Close();
            
            StreamWriter sw = new(Application.StartupPath + @"\logs.txt");
            sw.Write(conteudo.ToString());
            sw.WriteLine($"{texto}");
            sw.Close();
        }

        private void CarregarDgv()
        {

            dataGridView1.DataSource = cbxTipo.SelectedIndex < 0 ? "" : (cbxTipo.SelectedIndex == 0 ? internos : externos);
            lblAviso.Text = $"Registros encontrados: {dataGridView1.RowCount}";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass))
            {
                MessageBox.Show("Usuario e/ou Senha em branco.");
                return;
            }

            if (bot == null)
            {
                bot = new();
                bot.OpenPage("https://suap.ifma.edu.br/pesquisa/adicionar_comissao_por_area/187/");
            }

            if (bot.WaitElement(By.Id("id_username")))
            {
                if(!bot.SendText(By.Id("id_username"), user)) return;
                if(!bot.SendText(By.Id("id_password"), pass)) return;
                Thread.Sleep(300);
                if(!bot.Click(By.XPath(@"//input[@value='Acessar']"))) return;

                if(!bot.WaitElement(By.ClassName("notifications"))) return;

                bot.Click(By.XPath("//*[@id=\"bolsas_form\"]/table/thead/tr/th[1]/input"));
                Thread.Sleep(300);
                bot.Click(By.XPath("//*[@id=\"bolsas_form\"]/table/thead/tr/th[1]/input"));
                Thread.Sleep(300);

                foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    int row = bot.FindChild("//*[@id=\"bolsas_form\"]/table/tbody/tr", "td[2]", r.Cells[0].Value.ToString());

                    if (row > 0)
                    {
                        bot.Click(By.XPath($"//*[@id=\"bolsas_form\"]/table/tbody/tr[{row}]/td[1]/input"));
                    }
                    else
                    {
                        Logs($"Não encontrou o avaliador {r.Cells[0].Value}; {DateTime.Now}");
                    }
                }
            }
            else
            {
                MessageBox.Show("A página não foi carregada.");
            }
        }

        private void LerCredenciais()
        {
            StreamReader? sr = new StreamReader(Application.StartupPath + @"\config.txt");
            user = sr.ReadLine().Trim();
            pass = sr.ReadLine().Trim();
            sr.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(bot != null) bot.Close();

            if (excel != null)
            {
                excel.Close();
            }
        }

        private void cbxTipo_SelectedIndexChanged(object sender, EventArgs e)
        {
            CarregarDgv();
        }
    }
}