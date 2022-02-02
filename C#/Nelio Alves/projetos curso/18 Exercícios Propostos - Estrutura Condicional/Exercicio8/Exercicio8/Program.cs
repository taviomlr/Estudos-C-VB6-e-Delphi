using System;
using System.Globalization;

namespace Exercicio8
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Informe o valor so seu salário:");
            double salario = double.Parse(Console.ReadLine(), CultureInfo.InvariantCulture);
            double imposto;

            if (salario < 2000.0)
            {
                Console.WriteLine("Isento");
            }
            else if (salario > 2000.0 && salario < 3000.0)
            {
                imposto = (salario - 2000.0) * 0.08;
                Console.WriteLine("R$ " + imposto.ToString("F2", CultureInfo.InvariantCulture));
            }
            else if (salario > 3000.01 && salario < 4500.0)
            {
                imposto = (salario - 3000.0) * 0.18 + 1000.0 * 0.08;
                Console.WriteLine("R$ " + imposto.ToString("F2", CultureInfo.InvariantCulture));
            }
            else
            {
                imposto = (salario - 4500.0) * 0.28 + 1500.0 * 0.18 + 1000.0 * 0.08;
                Console.WriteLine("R$ " + imposto.ToString("F2", CultureInfo.InvariantCulture));
            }
        }
    }
}
