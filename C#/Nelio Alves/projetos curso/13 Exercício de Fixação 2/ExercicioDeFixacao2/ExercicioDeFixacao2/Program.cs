using System;
using System.Globalization;

namespace ExercicioDeFixacao2
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Entre com seu nome completo:");
            string nomeCompleto = Console.ReadLine();
            Console.WriteLine("Quantos quartos tem na sua casa?");
            int qtdQuartos = int.Parse(Console.ReadLine());
            Console.WriteLine("Entre com o preço do produto:");
            double precoProduto = double.Parse(Console.ReadLine());
            Console.WriteLine("Informe seu último nome, idade e altura (na mesma linha):");
            string[] vet = Console.ReadLine().Split(' ');
            string ultimoNome = vet[0];
            int idade = int.Parse(vet[1]);
            double altura = double.Parse(vet[2], CultureInfo.InvariantCulture);

            Console.WriteLine();
            Console.WriteLine("---------------------------------------");
            Console.WriteLine($"Nome Completo: {nomeCompleto}");
            Console.WriteLine($"Quantidade de quartos: {qtdQuartos}");
            Console.WriteLine($"Preço dos produtos: {precoProduto.ToString("F2", CultureInfo.InvariantCulture)}");
            Console.WriteLine(ultimoNome);
            Console.WriteLine($"Idade: {idade}");
            Console.WriteLine($"Altura: {altura.ToString("F2", CultureInfo.InvariantCulture)}");
            Console.WriteLine("---------------------------------------");
        }
    }
}
