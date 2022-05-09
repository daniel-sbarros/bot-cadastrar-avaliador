using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BotCadastrarAvaliador
{
    internal class Avaliador
    {
        public string? Nome { get; private set; }

        public Avaliador(string nome)
        {
            Nome = nome;
        }
    }
}
