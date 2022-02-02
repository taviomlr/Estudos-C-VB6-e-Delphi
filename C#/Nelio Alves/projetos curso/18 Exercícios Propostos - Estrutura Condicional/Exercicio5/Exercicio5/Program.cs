using System;
using System.Globalization;

namespace Exercicio5
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Digite o código e a quantidade do produto: ");
            string[] v = Console.ReadLine().Split(' ');
            int codProduto = int.Parse(v[0]);
            int qtdProduto = int.Parse(v[1]);

            Console.WriteLine();

            double vrProduto1 = 4.00;
            double vrProduto2 = 4.50;
            double vrProduto3 = 5.00;
            double vrProduto4 = 2.00;
            double vrProduto5 = 1.50;

            double vrTotal;

            if (codProduto == 1)
            {
                vrTotal = qtdProduto * vrProduto1;
                Console.WriteLine("Quantidade: " + qtdProduto +  " | Produto: Cachorro Quente | Preço: R$ " + vrProduto1.ToString("F2", CultureInfo.InvariantCulture));
                Console.WriteLine("Valor Total: R$ " + vrTotal.ToString("F2", CultureInfo.InvariantCulture));
            }
            else if (codProduto == 2)
            {
                vrTotal = qtdProduto * vrProduto2;
                Console.WriteLine("Quantidade: " + qtdProduto + " | Produto: X-Salada | Preço: R$ " + vrProduto2.ToString("F2", CultureInfo.InvariantCulture));
                Console.WriteLine("Valor Total: R$ " + vrTotal.ToString("F2", CultureInfo.InvariantCulture));
            }
            else if (codProduto == 3)
            {
                vrTotal = qtdProduto * vrProduto3;
                Console.WriteLine("Quantidade: " + qtdProduto + " | Produto: X-Bacon | Preço: R$ " + vrProduto3.ToString("F2", CultureInfo.InvariantCulture));
                Console.WriteLine("Valor Total: R$ " + vrTotal.ToString("F2", CultureInfo.InvariantCulture));
            }
            else if (codProduto == 4)
            {
                vrTotal = qtdProduto * vrProduto4;
                Console.WriteLine("Quantidade: " + qtdProduto + " | Produto: Torrada simples | Preço: R$ " + vrProduto4.ToString("F2", CultureInfo.InvariantCulture));
                Console.WriteLine("Valor Total: R$ " + vrTotal.ToString("F2", CultureInfo.InvariantCulture));
            }
            else
            {
                vrTotal = qtdProduto * vrProduto5;
                Console.WriteLine("Quantidade: " + qtdProduto + " | Produto: Refrigerante | Preço: R$ " + vrProduto5.ToString("F2", CultureInfo.InvariantCulture));
                Console.WriteLine("Valor Total: R$ " + vrTotal.ToString("F2", CultureInfo.InvariantCulture));
            }
        }
    }
}
