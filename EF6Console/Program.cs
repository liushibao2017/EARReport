using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EF6Console
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var ctx = new SchoolContext())
            {
                Student student = new Student() { StudentName = "bill" };
                ctx.Students.Add(student);
                ctx.SaveChanges();
                Console.WriteLine("初始化完成");
            }
        }
    }
}
