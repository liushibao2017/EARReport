using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                throw new NullReferenceException("C");
                Console.WriteLine("A");
            }
            catch(ArithmeticException e)
            {
                Console.WriteLine("B");
            }
            Console.ReadLine();
        }
    }
}
