using System;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            
            DateTime recordDate = DateTime.Parse("9/4/2021");
           
            if (recordDate.Date > DateTime.Now.Date )
            {
                Console.WriteLine("Hello World!");
            }
        }
    }
}
