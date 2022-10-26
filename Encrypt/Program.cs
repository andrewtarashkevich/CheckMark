using System;
using System.Linq;
using System.Collections.Generic;

namespace Encrypt
{
    class Program
    {
        static void Main(string[] args)
        {
            int[] a = new int[] {11,18,56,2,8,9};
            int result = a[0];

            for (int i = 1; i < a.Length; ++i)
            {
                result = gcd(result, a[i]);                
            }
            Console.WriteLine(result);    
        }
        static int gcd(int a, int b)
        {
            while (b != 0)
                b = a % (a = b);

            return a;
        }
    }
}


