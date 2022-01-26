using System;
using System.Globalization;

namespace SegundoProjeto
{
    class Program
    {
        static void Main(string[] args)
        {
            int idade = 32;
            string nome = "Maria";
            double saldo = 10.35784;

            Console.WriteLine(nome + " tem " + idade + " anos e tem um saldo de " + saldo.ToString("F2", CultureInfo.InvariantCulture) + " reais");
        }
    }
}
