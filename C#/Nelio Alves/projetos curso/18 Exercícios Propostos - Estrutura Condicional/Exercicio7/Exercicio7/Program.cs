using System;
using System.Globalization;

namespace Exercicio7
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Informe os valores do eixo X e Y:");
            string[] val = Console.ReadLine().Split(' ');
            double x = double.Parse(val[0], CultureInfo.InvariantCulture);
            double y = double.Parse(val[1], CultureInfo.InvariantCulture);

            if (x == 0.0 && y == 0.0)
            {
                Console.WriteLine("Origem");
            }
            else if (x == 0.0)
            {
                Console.WriteLine("Eixo y");
            }
            else if (y == 0.0)
            {
                Console.WriteLine("Eixo x");
            }

            else if (x > 0.0 && y > 0.0)
            {
                Console.WriteLine("Q1");
            }
            else if (x < 0.0 && y > 0.0)
            {
                Console.WriteLine("Q2");
            }
            else if (x < 0.0 && y < 0.0)
            {
                Console.WriteLine("Q3");
            }
            else
            {
            Console.WriteLine("Q4");
            }
        }
    }
}
