using System;

namespace OperadoresLogicos
{
    class Program
    {
        static void Main(string[] args)
        {
            bool c1 = 4 != 5; //true
            bool c2 = 2 > 3 && 4 != 5; // false
            bool c3 = 2 > 3 || 4 != 5; //true
            bool c4 = !(2 > 3) && 4 != 5; //true
            bool c5 = 10 < 5; //false
            bool c6 = c3 || c4 && c5; //true

            Console.WriteLine(c1);
            Console.WriteLine(c2);
            Console.WriteLine(c3);
            Console.WriteLine(c4);
            Console.WriteLine("----------------------");
            Console.WriteLine(c5);
            Console.WriteLine(c6);
        }
    }
}
