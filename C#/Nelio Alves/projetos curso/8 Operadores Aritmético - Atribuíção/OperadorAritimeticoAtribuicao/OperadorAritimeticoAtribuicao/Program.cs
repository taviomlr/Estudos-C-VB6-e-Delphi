using System;

namespace OperadorAritimeticoAtribuicao
{
    class Program
    {
        static void Main(string[] args)
        {
            int a = 10;
            a++;
            Console.WriteLine(a);
            
            int b = 10;
            b--;
            Console.WriteLine(b);

            Console.WriteLine("----------------");

            int c = 10;
            int d = c++;
            int e = ++c;
            Console.WriteLine(c);
            Console.WriteLine(d);
            Console.WriteLine(e);

        }
    }
}
