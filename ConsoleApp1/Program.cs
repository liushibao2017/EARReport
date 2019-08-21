using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            int a = 100;
            int b=66;
            Fun(ref a, ref b);
            Console.WriteLine("a:{0},b{1}",a,b);

        }
        static void Fun(ref int a, ref int b)
        {
            a = a + b;
            b = 1;
        }
    }
}
