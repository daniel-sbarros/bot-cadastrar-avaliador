using OpenQA.Selenium;
using System.Text.RegularExpressions;

namespace BotCadastrarAvaliador
{
    public partial class Form1 : Form
    {
        string? user, pass;
        Bot bot;
        MSExcel excel;
        List<Avaliador> avaliadores, internos, externos;
        FileStream logs;
        Thread thread2;

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

        private bool findAvaliador(List<Avaliador> avaliadores, string value)
        {
            foreach (var av in avaliadores)
            {
                if (av.Nome.Contains(value)) return true;
            }
            return false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (cbxTipo.SelectedIndex < 0)
            {
                MessageBox.Show("Selecione o tipo de registro.");
                return;
            }

            string url = @"https://suap.ifma.edu.br/pesquisa/adicionar_comissao_por_area/187/";
            List<Avaliador> lista = cbxTipo.SelectedIndex == 0 ? internos : externos;
            BotaoExec(button1, false);

            thread2 = new Thread(() =>
            {
                DateTime hr_inicio = DateTime.Now;

                if (string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass))
                {
                    MessageBox.Show("Usuario e/ou Senha em branco.");
                    return;
                }
                
                createLogsFile();

                if (bot == null)
                {
                    bot = new();
                    bot.OpenPage(url);
                }

                if (bot.WaitElement(By.Id("id_username")))
                {
                    if (bot.Url != url)
                    {
                        if (!bot.SendText(By.Id("id_username"), user)) return;
                        if (!bot.SendText(By.Id("id_password"), pass)) return;
                        Thread.Sleep(300);
                        if (!bot.Click(By.XPath(@"//input[@value='Acessar']"))) return;
                    }

                    if (!bot.WaitElement(By.ClassName("notifications"))) return;

                    if (avaliadores == null)
                    {
                        avaliadores = new();

                        for (int i = 1; i <= bot.getCount(By.XPath("//*[@id=\"bolsas_form\"]/table/tbody/tr")); i++)
                        {
                            avaliadores.Add(new Avaliador(bot.getText(By.XPath($"//*[@id=\"bolsas_form\"]/table/tbody/tr[{i}]/td[2]")).ToUpper().Trim()));
                        }
                    }

                    for (int i = 0; i < lista.Count; i++)
                    {
                        if ((i + 1) % 10 == 0) editTextbox($"Processando: '{i}'", lblAndamento);

                        if (findAvaliador(avaliadores, lista[i].Nome.ToUpper()))
                        {
                            for (int r = 0; r < avaliadores.Count; r++)
                            {
                                if (avaliadores[r].Nome.Contains(lista[i].Nome.ToUpper()))
                                {
                                    if (!bot.isChecked(By.XPath($"//*[@id=\"bolsas_form\"]/table/tbody/tr[{(r + 1)}]/td[1]/input")))
                                    {
                                        bot.Click(By.XPath($"//*[@id=\"bolsas_form\"]/table/tbody/tr[{(r + 1)}]/td[1]/input"));
                                    }
                                    break;
                                }
                            }
                        }
                        else
                        {
                            Logs($"{DateTime.Now} -> Não encontrou o avaliador: '{lista[i].Nome.Trim()}'.\n", logs.Name);
                        }
                    }

                    Thread.Sleep(100);
                    FormNaFrente();

                    TimeSpan tempo_limite = new TimeSpan(1, 25, 0);

                    if (DateTime.Now < (hr_inicio + tempo_limite))
                    {
                        bot.Click(By.Name("Salvar"));
                    }
                    
                    BotaoExec(button1, true);
                    MessageBox.Show("TAREFA CONCLUÍDA.");
                }
                else
                {
                    MessageBox.Show("A página não foi carregada.");
                }
            });
            thread2.Start();
        }

        private void editTextbox(string text, Label label)
        {
            if (label.InvokeRequired)
                label.Invoke(new MethodInvoker(() => label.Text = text));
            else
                label.Text = text;
        }

        private void FormNaFrente()
        {
            if (this.InvokeRequired)
                this.Invoke(new MethodInvoker(() => this.Activate()));
            else
                this.Activate();
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
            StreamReader? sr = new StreamReader(Application.StartupPath + @"\config.txt");
            user = sr.ReadLine().Trim();
            pass = sr.ReadLine().Trim();
            sr.Close();
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