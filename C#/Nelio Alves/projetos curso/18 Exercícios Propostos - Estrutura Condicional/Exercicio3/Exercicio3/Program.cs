using System;

namespace Exercicio3
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Informe o primeiro número inteiros:");
            int A = int.Parse(Console.ReadLine());

            Console.WriteLine("Informe o segundo número inteiros:");
            int B = int.Parse(Console.ReadLine());

            if (A % B == 0 || B % A == 0)
            {
                Console.WriteLine("São Multiplos");
            }
            else
            {
                Console.WriteLine("Não São Multiplos");
            }
        }
    }
}
