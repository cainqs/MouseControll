using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Timers;

namespace MouseControl
{
    public class Program
    {
        private static Timer timer1 = new Timer();
        private static int interval = 120;
        private static int SW_SHOWNOMAL = 1;
        private static Point p;

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out Point pt);
        [DllImport("User32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int cmdShow);
        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("User32.dll")]
        static extern void mouse_event(int flags, int dX, int dY, int buttons, int extraInfo);
        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;                            
            public int Top;                             
            public int Right;                           
            public int Bottom;                        
        }

        const int MOUSEEVENTF_MOVE = 0x1;
        const int MOUSEEVENTF_LEFTDOWN = 0x2;
        const int MOUSEEVENTF_LEFTUP = 0x4;
        const int MOUSEEVENTF_RIGHTDOWN = 0x8;
        const int MOUSEEVENTF_RIGHTUP = 0x10;
        const int MOUSEEVENTF_MIDDLEDOWN = 0x20;
        const int MOUSEEVENTF_MIDDLEUP = 0x40;
        const int MOUSEEVENTF_WHEEL = 0x800;
        const int MOUSEEVENTF_ABSOLUTE = 0x8000;


        public static void Main(string[] args)
        {
            InitProcess();
        }

        private static void InitProcess()
        {
            Console.WriteLine("输入间隔时间（单位秒），默认2分钟");
            var count = Console.ReadLine();

            if (!int.TryParse(count, out interval))
            {
                interval = 120;
                Console.WriteLine("输入无效采用默认2分钟");
            }

            Console.WriteLine(string.Format("间隔时间{0}秒", interval));

            timer1.Interval = 1000 * interval;
            timer1.Enabled = true;
            timer1.Elapsed += new ElapsedEventHandler(timer1_Tick);

            Console.WriteLine("按任意键退出");
            Console.ReadKey();
        }

        private static void timer1_Tick(object source, ElapsedEventArgs e)
        {
            GetCursorPos(out p);
            var currentP = Process.GetCurrentProcess();
            var process = GetCurrentProcess();

            foreach (var p in process)
            {
                HandleRunningInstance(p);
                var tempP = GetCurrentPosition();

                mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE, (tempP.X + 500) * 65536 / 1920, (tempP.Y + 500) * 65536 / 1080, 0, 0);
                mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);

                System.Threading.Thread.Sleep(500);
            }

            SetForegroundWindow(currentP.MainWindowHandle);
            mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE, (p.X) * 65536 / 1920, (p.Y) * 65536 / 1080, 0, 0);
        }

        private static List<Process> GetCurrentProcess()
        {
            Process[] procs = Process.GetProcessesByName("mstsc");
            bool hasWindow = true;

            if (procs.Length <= 0)
                return null;

            foreach (Process proc in procs)
            {
                if (proc.MainWindowHandle == IntPtr.Zero)
                {
                    hasWindow = false;
                }
            }

            if (hasWindow)
            {
                return procs.ToList();
            }
            else
            {
                return null;
            }
        }

        private static void HandleRunningInstance(Process instance)
        {
            ShowWindowAsync(instance.MainWindowHandle, SW_SHOWNOMAL);//显示
            SetForegroundWindow(instance.MainWindowHandle);//当到最前端
        }

        private static Point GetCurrentPosition()
        {
            IntPtr awin = GetForegroundWindow();    //获取当前窗口句柄
            RECT rect = new RECT();
            GetWindowRect(awin, ref rect);
            int width = rect.Right - rect.Left;                        //窗口的宽度
            int height = rect.Bottom - rect.Top;                   //窗口的高度
            int x = rect.Left;
            int y = rect.Top;

            return new Point
            {
                X = x,
                Y = y
            };
        }
    }
}
