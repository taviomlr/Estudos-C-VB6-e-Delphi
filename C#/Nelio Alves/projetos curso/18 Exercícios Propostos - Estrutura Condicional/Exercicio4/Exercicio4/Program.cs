using System;

namespace Exercicio3
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Informe a hora de início do Jogo (de 0 a 24)");
            int horaInicio = int.Parse(Console.ReadLine());

            Console.WriteLine("Informe a hora de término do Jogo (de 0 a 24)");
            int horaTermino = int.Parse(Console.ReadLine());

            int mesmoDia = horaTermino - horaInicio;
            int outroDia = (24 - horaInicio) + horaTermino;

            if (horaInicio > horaTermino)
            {
                Console.WriteLine("O JOGO DUROU " + outroDia + " HORA(S)");
            }
            else if (horaInicio < horaTermino)
            {
                Console.WriteLine("O JOGO DUROU " + mesmoDia + " HORA(S)");
            }
            else
            {
                Console.WriteLine("O JOGO DUROU 24 HORA(S)");
            }
        }
    }
}
