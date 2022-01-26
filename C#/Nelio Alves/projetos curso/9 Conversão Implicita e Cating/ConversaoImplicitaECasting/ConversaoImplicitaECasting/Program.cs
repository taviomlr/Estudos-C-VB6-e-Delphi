using System;

namespace ConversaoImplicitaECasting
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("----------FLOAT PARA DOUBLE (CONVERSÃO IMPLICITA)----------");
            
            float x = 4.5f;
            double y = x;
            Console.WriteLine(y);

            Console.WriteLine("----------DOUBLE PARA FLOAT (CASTING - CONVERSÃO EXPLICITA)----------");
            
            double a;
            float b;
            a = 5.1;
            b =  (float)a;
            Console.WriteLine(b);

            Console.WriteLine("----------DOUBLE PARA INT(CASTING - CONVERSÃO EXPLICITA - VALOR TRUNCADO)----------");

            double c;
            int d;

            c = 5.1;
            d = (int)c;
            Console.WriteLine(d);

            Console.WriteLine("----------INT PARA DOUBLE(CASTING - CONVERSÃO EXPLICITA)----------");

            int e = 5;
            int f = 2;

            double resultado = (double)e / f;

            Console.WriteLine(resultado);
        }
    }
}
