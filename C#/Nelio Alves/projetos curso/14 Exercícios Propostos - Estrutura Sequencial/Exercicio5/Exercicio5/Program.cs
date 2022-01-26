using System;
using System.Globalization;

namespace Exercicio5
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Informe os dados da peça 1:");
            string[] vet1 = Console.ReadLine().Split(' ');
            int codPeca1 = int.Parse(vet1[0]);
            int qtdPeca1 = int.Parse(vet1[1]);
            double valPeca1 = double.Parse(vet1[2], CultureInfo.InvariantCulture);

            Console.WriteLine("Informe os dados da peça 1:");
            string[] vet2 = Console.ReadLine().Split(' ');
            int codPeca2 = int.Parse(vet2[0]);
            int qtdPeca2 = int.Parse(vet2[1]);
            double valPeca2 = double.Parse(vet2[2], CultureInfo.InvariantCulture);

            double vrPgto = (qtdPeca1 * valPeca1) + (qtdPeca2 * valPeca2);

            Console.WriteLine("Valor a pagar: R$ " + vrPgto.ToString("F2", CultureInfo.InvariantCulture));
        }
    }
}
