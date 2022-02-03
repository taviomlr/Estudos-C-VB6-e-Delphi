
using System;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Informe o Valor de A:");
            int A = int.Parse(Console.ReadLine());
            Console.WriteLine("Informe o Valor de A:");
            int B = int.Parse(Console.ReadLine());

            int x = A + B;
            Console.WriteLine("x = " + x);
        }
    }
}
