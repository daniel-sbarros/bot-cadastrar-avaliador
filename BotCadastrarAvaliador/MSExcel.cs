using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BotCadastrarAvaliador
{
    class MSExcel
    {
        Excel.Application excel = null;
        Excel.Workbook book = null;
        Excel.Worksheet sheet;

        public MSExcel(string ArquivoOriginal = null)
        {
            try
            {
                excel = new Excel.Application();

                if (ArquivoOriginal != null)
                {
                    book = excel.Workbooks.Add(Application.StartupPath + @"\assets\" + ArquivoOriginal);
                }
                else
                {
                    book = excel.Workbooks.Add();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao carregar o arquivo.\n " + ex.Message);
            }
        }

        public bool ModificarValores(int PlanilhaIndex, string CelulaInicial, DataGridView Tabela, int TabelaColuna, int QuantidadeDeValores = 1, bool PreenchimentoVertical = true)
        {
            try
            {
                sheet = book.Sheets[PlanilhaIndex]; // SELECIONAR PLANILHA(ABA)
                sheet.Select();

                sheet.Range[CelulaInicial].Select(); // SELECIONAR CELULA

                for (int i = 0; i < QuantidadeDeValores; i++)
                {
                    excel.ActiveCell.Value = Tabela.Rows[i].Cells[TabelaColuna].EditedFormattedValue; // MODIFICA O VALOR DA CELULAR SELECIONADA

                    if (PreenchimentoVertical) excel.ActiveCell.Offset[1, 0].Select(); // SELECIONA A CELULA DE BAIXO
                    else excel.ActiveCell.Offset[0, 1].Select(); // SELECIONA A CELULA DA DIREITA
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("METODO: MODIFICAR VALORES.\n" + ex.Message);
                return false;
            }
        }

        public bool ModificarValor(int PlanilhaIndex, string Celula, object Valor)
        {
            try
            {
                sheet = book.Sheets[PlanilhaIndex]; // SELECIONAR PLANILHA(ABA)
                sheet.Select();

                sheet.Range[Celula].Select(); // SELECIONAR CELULA
                excel.ActiveCell.Value = Valor; // MODIFICA O VALOR DA CELULAR SELECIONADA

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("METODO: MODIFICAR VALORES.\n" + ex.Message);
                return false;
            }
        }

        public bool ModificarValor(int PlanilhaIndex, object[,] CelulaEValor)
        {
            try
            {
                sheet = book.Sheets[PlanilhaIndex]; // SELECIONAR PLANILHA(ABA)
                sheet.Select();

                for (int i = 0; i < (CelulaEValor.Length / 2); i++)
                {
                    sheet.Range[CelulaEValor[i, 0]].Select(); // SELECIONAR CELULA
                    excel.ActiveCell.Value = CelulaEValor[i, 1]; // MODIFICA O VALOR DA CELULAR SELECIONADA
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("METODO: MODIFICAR VALORES.\n" + ex.Message);
                return false;
            }
        }

        public void ProtegerPlanilha(string senha)
        {
            for (int i = 1; i <= excel.Sheets.Count; i++)
            {
                sheet = book.Sheets[i];
                sheet.Select();
                sheet.Protect(senha);
            }
        }

        public void ExibirExcel()
        {
            excel.Visible = true;
        }

        public static void criarPlanilha(DataGridView Tabela, string[] colunas = null)
        {
            // CRIAR OBJETOS
            Excel.Application app = new Excel.Application();
            Excel.Workbook book;
            Excel.Worksheet sheet;

            // ATRIBUIR WORKBOOK
            book = app.Workbooks.Add();

            // SELECIONAR PLANILHA
            sheet = book.Sheets[1];
            sheet.Select();

            // PREENCHER COLUNAS TÍTULOS **********************************************************************************
            sheet.Range["A1"].Select();

            for (int i = 0; i < Tabela.ColumnCount; i++)
            {
                app.ActiveCell.Value = Tabela.Columns[i].Name;
                app.ActiveCell.Offset[0, 1].Select();
            }

            // PREENCHER CÉLULAS ******************************************************************************************
            sheet.Range["A2"].Select();

            foreach (DataGridViewRow r in Tabela.Rows)
            {
                for (int i = 0; i < Tabela.Columns.Count; i++)
                {
                    app.ActiveCell.Offset[0, i].Value = r.Cells[i].EditedFormattedValue;
                }
                app.ActiveCell.Offset[1, 0].Select();
            }

            app.Columns["A:Z"].EntireColumn.AutoFit(); // AJUSTA A LARGURA DAS COLUNAS

            // DEIXA O EXCEL VISIVEL **************************************************************************************
            app.Visible = true;
        }
    }
}
