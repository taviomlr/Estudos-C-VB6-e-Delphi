using System;

namespace Exercicio2
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Informe um número inteiro:");
            int x = int.Parse(Console.ReadLine());

            if (x % 2 == 0)
            {
                Console.WriteLine("PAR");
            }
            else
            {
                Console.WriteLine("IMPAR");
            }
        }
    }
}
