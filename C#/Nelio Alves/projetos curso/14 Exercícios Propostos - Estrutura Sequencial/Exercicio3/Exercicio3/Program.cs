using System;

namespace Exercicio3
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Informe o valor para A:");
            int A = int.Parse(Console.ReadLine());

            Console.WriteLine("Informe o valor para B:");
            int B = int.Parse(Console.ReadLine());

            Console.WriteLine("Informe o valor para C:");
            int C = int.Parse(Console.ReadLine());

            Console.WriteLine("Informe o valor para D:");
            int D = int.Parse(Console.ReadLine());

            int diferenca = (A * B) - (C * D);

            Console.WriteLine($"A diferença é: {diferenca}");
        }
    }
}
