﻿using System;
using System.Globalization;

namespace SegundoProjeto
{
    class Program
    {
        static void Main(string[] args)
        {
            char genero = 'F';
            int idade = 32;
            string nome = "Maria";

            double saldo = 10.35784;

            Console.WriteLine("Bom dia!");
            Console.Write("Boa tarde!");
            Console.WriteLine("Boa noite!");
            Console.WriteLine("----------------------");
            Console.WriteLine(genero);
            Console.WriteLine(idade);
            Console.WriteLine(nome);
            Console.WriteLine("----------------------");
            Console.WriteLine(saldo);
            Console.WriteLine(saldo.ToString("F2"));
            Console.WriteLine(saldo.ToString("F4"));
            Console.WriteLine(saldo.ToString("F4", CultureInfo.InvariantCulture));

        }
    }
}