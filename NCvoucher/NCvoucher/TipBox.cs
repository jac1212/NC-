using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace NCvoucher
{
    class TipBox
    {
        static void Main1(string[] args)
        {
            test1();
            Console.WriteLine("123");
        }
        private static async Task test1()
        {
            await test2();
            Console.WriteLine("456");
        }
        // 模拟耗时操作（异步方法）
        private static async Task test2()
        {
            for (int i = 0; i < 5; ++i)
            {
                Console.WriteLine("耗时操作{0}", i);
                await Task.Delay(1);
            }
        }

    }
    
}
