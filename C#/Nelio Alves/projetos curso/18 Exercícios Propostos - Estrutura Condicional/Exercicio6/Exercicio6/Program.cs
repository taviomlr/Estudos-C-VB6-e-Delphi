using System;
using System.Globalization;

namespace Exercicio6
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Digite um número:");
            double val = double.Parse(Console.ReadLine(), CultureInfo.InvariantCulture);
            
            if (val > 0.00 && val <= 25.00)
            {
                Console.WriteLine("Intervalo (0, 25)");
            }
            else if (val > 25.00 && val <= 50.00)
            {
                Console.WriteLine("Intervalo (25, 50)");
            }
            else if (val > 50.00 && val <= 75.00)
            {
                Console.WriteLine("Intervalo (50, 75)");
            }
            else if (val > 75.00 && val <= 100.00)
            {
                Console.WriteLine("Intervalo (75, 100)");
            }
            else
            {
                Console.WriteLine("Fora do Intervalo");
            }

        }
    }
}
