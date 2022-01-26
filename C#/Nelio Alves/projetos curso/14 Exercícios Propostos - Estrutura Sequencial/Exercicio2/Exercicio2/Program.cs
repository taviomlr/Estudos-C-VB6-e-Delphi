using System;
using System.Globalization;

namespace Exercicio2
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Informe o valor do raio do círculo:");
            double valor = double.Parse(Console.ReadLine(), CultureInfo.InvariantCulture);
            double pi = 3.14159;
            double potencia = Math.Pow(valor, 2.0);
            double area = pi * potencia;

            Console.WriteLine($"O valor da área é: {area.ToString("F4", CultureInfo.InvariantCulture)}");
            
        }
    }
}
