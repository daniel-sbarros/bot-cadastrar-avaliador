using OpenQA.Selenium;
using System.Text.RegularExpressions;

namespace BotCadastrarAvaliador
{
    public partial class Form1 : Form
    {
        string? user, pass;
        Bot bot;
        MSExcel excel;
        List<Avaliador> internos, externos;
        List<string> avaliadores;
        FileStream logs;
        Thread thread2;
        int cont_erros = 0;

        public Form1()
        {
            avaliadores = null;
            InitializeComponent();
        }

        private void createLogsFile()
        {
            var now = DateTime.Now;
            logs = File.Create($"logs_-_{now.Year}-{My.FormatNumber(now.Month)}-{My.FormatNumber(now.Day)}_-_{My.FormatNumber(now.Hour)}-{My.FormatNumber(now.Minute)}-{My.FormatNumber(now.Second)}.txt");
            logs.Close();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            LerCredenciais();
            lblStatus.Text = "";

            excel = excel ?? new MSExcel("cadastro-de-avaliador-interno-externo.xlsx");
            internos = internos ?? excel.GetListValues(1);
            externos = externos ?? excel.GetListValues(2);

            cbxTipo.SelectedIndex = 0;
        }

        public void Logs(object texto, string name_file)
        {
            StreamReader sr = new(name_file);
            string line = sr.ReadLine();
            System.Text.StringBuilder conteudo = new();

            while (line != null) 
            {
                conteudo.AppendLine(line);
                line = sr.ReadLine();
            }
            sr.Close();
            
            StreamWriter sw = new(name_file);
            sw.Write(conteudo.ToString());
            sw.WriteLine($"{texto}");
            sw.Close();
        }

        private void CarregarDgv()
        {
            dataGridView1.DataSource = cbxTipo.SelectedIndex < 0 ? "" : (cbxTipo.SelectedIndex == 0 ? internos : externos);
            lblAviso.Text = $"Registros encontrados: {dataGridView1.RowCount}";
            dataGridView1.AutoResizeColumns();
            dataGridView1.AutoResizeRows();
        }

        private void pagNaoCarregada(string value)
        {
            Status(value, lblStatus);
            BotaoExec(button1, true);
            MessageBox.Show($"{value}");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Status("", lblStatus);

            if (cbxTipo.SelectedIndex < 0)
            {
                MessageBox.Show("Selecione o tipo de registro.");
                return;
            }

            string tipo = cbxTipo.Text.ToUpper();
            string url = @"https://suap.ifma.edu.br/pesquisa/adicionar_comissao_por_area/187/";
            List<Avaliador> lista = cbxTipo.SelectedIndex == 0 ? internos : externos;
            BotaoExec(button1, false);

            thread2 = new Thread(() =>
            {
                DateTime hr_inicio = DateTime.Now;

                if (string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass))
                {
                    //MessageBox.Show("Usuario e/ou Senha em branco.");
                    pagNaoCarregada("Usuario e/ou Senha em branco."); return;
                }

                if(logs == null) createLogsFile();

                bot = bot ?? new();
                if (bot.Url != url) bot.OpenPage(url);

                if (bot.Url != url)
                {
                    if (bot.Url.Contains(@"https://suap.ifma.edu.br/accounts/login/?"))
                    {
                        if (bot.WaitElement(By.Id("id_username")))
                        {
                            if (!bot.SendText(By.Id("id_username"), user) || !bot.SendText(By.Id("id_password"), pass))
                            {
                                pagNaoCarregada("Elementos id_username e/ou id_password não foram carregados."); return;
                            }
                            Thread.Sleep(300);
                            if (!bot.Click(By.XPath(@"//input[@value='Acessar']")))
                            {
                                pagNaoCarregada("Não foi possível clicar no botão salvar."); return;
                            }
                        }
                        else
                        {
                            pagNaoCarregada("O elemento id_username não foi carregado."); return;
                        }
                    }
                    else
                    {
                        pagNaoCarregada("A página de login não foi carregada."); return;
                    }
                }

                if (!bot.WaitElement(By.Name("Salvar"), 60))
                {
                    pagNaoCarregada("O botão Salvar não foi carregado."); return;
                }

                if (avaliadores == null)
                {
                    Status("Carregando Avaliadores da página Web...", lblStatus);
                    avaliadores = new();

                    for (int i = 1; i <= bot.getCount(By.XPath("//*[@id=\"bolsas_form\"]/table/tbody/tr")); i++)
                    {
                        avaliadores.Add(bot.getText(By.XPath($"//*[@id=\"bolsas_form\"]/table/tbody/tr[{i}]/td[2]")).ToUpper().Trim());
                    }
                    Status($"Avaliadores cadastrados.", lblStatus);
                }

                Logs($"\n>>>>>>>>>>>>>>>>>>>>>>>>>>>> {tipo} <<<<<<<<<<<<<<<<<<<<<<<<<<<<\n\n", logs.Name);

                for (int l = 0; l < lista.Count; l++)
                {
                    if (avaliadores.Count(av => av.Contains(lista[l].Nome.ToUpper())) > 0)
                    {
                        for (int r = 0; r < avaliadores.Count; r++)
                        {
                            if (avaliadores[r].Contains(lista[l].Nome.ToUpper()))
                            {
                                if (!bot.isChecked(By.XPath($"//*[@id=\"bolsas_form\"]/table/tbody/tr[{(r + 1)}]/td[1]/input")))
                                {
                                    bot.Click(By.XPath($"//*[@id=\"bolsas_form\"]/table/tbody/tr[{(r + 1)}]/td[1]/input"));
                                    Logs($">>>>>>>>>>>>>> {lista[l].Nome.ToUpper()} foi selecionado.\n", logs.Name);
                                }
                                break;
                            }
                        }
                    }
                    else
                    {
                        cont_erros++;
                        Logs($">> [ERROR] {DateTime.Now} -> Não encontrou o avaliador: \n{lista[l].Nome.Trim()}\n", logs.Name);
                    }
                }

                Logs($"\n>>>>>>>>>>>>>>>>>>>>>>>>>>>> TOTAL DE REGISTROS NÃO ENCOTRADOS: {cont_erros} <<<<<<<<<<<<<<<<<<<<<<<<<<<<\n\n", logs.Name);

                Thread.Sleep(100);

                TimeSpan tempo_limite = new TimeSpan(1, 25, 0);

                if (DateTime.Now < (hr_inicio + tempo_limite))
                {
                    bot.Click(By.Name("Salvar"));
                }
                
                Status($"Avaliadores {tipo} cadastrados.", lblStatus);
                BotaoExec(button1, true);
                MessageBox.Show("TAREFA CONCLUÍDA.");
            });
            thread2.Start();
        }

        private void Status(string value, Label label)
        {
            if (label.InvokeRequired)
            {
                label.Invoke(new MethodInvoker(() => label.Text = value));
            }
            else
            {
                label.Text = value;
            }
        }

        private void BotaoExec(Button btn, bool value)
        {
            if (btn.InvokeRequired)
                btn.Invoke(new MethodInvoker(() =>
                {
                    btn.Enabled = value;
                    btn.Cursor = value ? Cursors.Hand : Cursors.Default;
                }));
            else
            {
                btn.Enabled = value;
                btn.Cursor = value ? Cursors.Hand : Cursors.Default;
            }
        }

        private void LerCredenciais()
        {
            try
            {
                StreamReader? sr = new StreamReader(Application.StartupPath + @"\config.txt");
                user = sr.ReadLine().Trim();
                pass = sr.ReadLine().Trim();
                sr.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Não foi possível encontrar o arquivo config.txt, adicione o arquivo a pasta raiz do aplicativo.");
                this.Close();
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (thread2 != null && thread2.ThreadState == ThreadState.Running) thread2.Abort();
                if (bot != null) bot.Close();
                if (excel != null) excel.Close();
            }
            catch (Exception) { }
        }

        private void cbxTipo_SelectedIndexChanged(object sender, EventArgs e)
        {
            CarregarDgv();
        }
    }
}