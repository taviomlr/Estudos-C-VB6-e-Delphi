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

            Console.WriteLine("{0} tem {1} anos e tem um saldo de {2:F2} reais", nome, idade, saldo);
        }
    }
}