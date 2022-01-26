using System;
using System.Globalization;

namespace Exercicio4
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Informe o número do funcionário:");
            int numeroFunc = int.Parse(Console.ReadLine());

            Console.WriteLine("Informe o número de horas trabalhadas:");
            int numHoras = int.Parse(Console.ReadLine());

            Console.WriteLine("Informe o valor da hora:");
            double valHora = double.Parse(Console.ReadLine(), CultureInfo.InvariantCulture);

            double salario = numHoras * valHora;

            Console.WriteLine($"Número: {numeroFunc}");
            Console.WriteLine("Salário: R$ " + salario.ToString("F2", CultureInfo.InvariantCulture));
        }
    }
}
