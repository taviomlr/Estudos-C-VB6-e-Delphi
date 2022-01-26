﻿using System;

namespace OperadoresAritimeticos
{
    class Program
    {
        static void Main(string[] args)
        {
            int n1 = 3 + 4 * 2;
            int n2 = (3 + 4) * 2;
            int n3 = 17 % 3;
            int n4 = 10 / 8;
            double n5 = 10 / 8;
            double n6 = 10.0 / 8.0;

            double a = 1.0, b = -3.0, c = -4.0;
            double delta = Math.Pow(b, 2.0) - 4.0 * a * c;
            double x1 = (-b + Math.Sqrt(delta)) / (2.0 * a);
            double x2 = (-b - Math.Sqrt(delta)) / (2.0 * a);


            Console.WriteLine($"n1 = {n1}");
            Console.WriteLine($"n2 = {n2}");
            Console.WriteLine($"n3 = {n3}");
            Console.WriteLine($"n4 = {n4}");
            Console.WriteLine($"n5 = {n5}");
            Console.WriteLine($"n6 = {n6}");
            Console.WriteLine($"delta = {delta}");
            Console.WriteLine($"x1 = {x1}");
            Console.WriteLine($"x2 = {x2}");
        }
    }
}
