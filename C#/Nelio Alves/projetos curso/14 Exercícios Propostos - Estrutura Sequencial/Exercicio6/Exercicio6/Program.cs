using System;
using System.Globalization;

namespace Exercicio6
{
    class Program
    {
        static void Main(string[] args)
        {
            double A, B, C, areaTriangulo, areaCirculo, areaTrapezio, areaQuadrado, areaRetangulo;
            double pi = 3.14159;

            Console.WriteLine("Informe os valores de A, B e C:");
            string[] v = Console.ReadLine().Split(' ');
            A = double.Parse(v[0], CultureInfo.InvariantCulture);
            B = double.Parse(v[1], CultureInfo.InvariantCulture);
            C = double.Parse(v[2], CultureInfo.InvariantCulture);

            areaTriangulo = (A * C) / 2;
            areaCirculo = pi * Math.Pow(C, 2.0);
            areaTrapezio = ((A + B) * C) / 2;
            areaQuadrado = B * B;
            areaRetangulo = A * B;

            Console.WriteLine();
            Console.WriteLine("----------------------");
            Console.WriteLine("TRIÂNGULO: " + areaTriangulo.ToString("F3", CultureInfo.InvariantCulture));
            Console.WriteLine("CÍRCULO: " + areaCirculo.ToString("F3", CultureInfo.InvariantCulture));
            Console.WriteLine("TRAPÉZIO: " + areaTrapezio.ToString("F3", CultureInfo.InvariantCulture));
            Console.WriteLine("QUADRADO: " + areaQuadrado.ToString("F3", CultureInfo.InvariantCulture));
            Console.WriteLine("RETÂNGULO: " + areaRetangulo.ToString("F3", CultureInfo.InvariantCulture));
            Console.WriteLine("----------------------");
        }
    }
}
