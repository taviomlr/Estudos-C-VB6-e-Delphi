using System;

namespace Exercicio1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Informe o valor 1:");
            int valor1 = int.Parse(Console.ReadLine());
            Console.WriteLine("Informe o valor 2:");
            int valor2 = int.Parse(Console.ReadLine());

            int soma = valor1 + valor2;
            Console.WriteLine($"O resultado da soma é: {soma}");
        }
    }
}
