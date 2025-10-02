using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.CompilerServices;
using System.Reflection.Metadata;
using WindowsInput;
using WindowsInput.Native;

//using WindowsInput;
//using WindowsInput.Native;

namespace excel_clicker
{
    class Program
    {

        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point); 
        

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);


        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);


        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        // Delegate to filter which windows to include 
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        static void Main(string[] args)
        {
            Process[] ps = Process.GetProcessesByName("EXCEL"); // Get the name of the process (Task Manager > Details tab)
            Process excelProcess = ps.FirstOrDefault();
            InputSimulator isim = new InputSimulator(); // Used bc SendKeys.SendWait() doesn't always work


            ///*

            if (excelProcess != null) {
                // Bring Excel into focus
                Console.WriteLine("Bringing Excel into focus");
                IntPtr h = excelProcess.MainWindowHandle;
                SetForegroundWindow(h); // Works if the window is NOT minimized

                // Copy contents. For some reason ^C, ^{C}, and ^(C) don't work. Guess it's not capital C, maybe because default is lowercase c?
                SendKeys.SendWait("^c"); 
                Console.WriteLine("Copying contents...");

            } else {            
                Console.WriteLine("Couldn't find the Excel application.");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }

            //*/

            // Next step: Get the Kronos browser
            Process[] kr = Process.GetProcessesByName("msedge");
            Process kronosProcess = kr.FirstOrDefault();

            if (kronosProcess != null) {
                // Bring Edge into focus - only whatever the current tab is, does not find Kronos specifically
                Console.WriteLine("Bringing Microsoft Edge into focus");
                IntPtr k = kronosProcess.MainWindowHandle;
                SetForegroundWindow(k); // Works if the window is NOT minimized
            } else {
                Console.WriteLine("Couldn't find an open Microsoft Edge window.");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }

            ///*

            // Paste contents into search box. Prerequisite: Activity Name selected
            // Weird error: instead of the project name being copied, `SendKeys.SendWait("^v");*` was copied instead lol
            copyExcelContents(isim);

            //SendKeys.SendWait("^a");
            //Thread.Sleep(1000);
            //SendKeys.SendWait("^v*{TAB}");
            //Console.WriteLine("Pasting contents...");
            //isim.Keyboard.KeyPress(VirtualKeyCode.RETURN);

            //*/



            /********* TODO: Hard coded to leftmost monitor on in-office setup, try to find HTML element *********/

            // Open & Close the project
            closeProject(isim);

            //// Open project 
            //isim.Mouse.MoveMouseToPositionOnVirtualDesktop(2000, 27050);
            //Thread.Sleep(1000);
            //isim.Mouse.LeftButtonDoubleClick();
            //Thread.Sleep(1000);

            //// Move mouse to Status dropdown
            //isim.Mouse.MoveMouseToPositionOnVirtualDesktop(15000, 26000);
            //Thread.Sleep(1000);
            //isim.Mouse.LeftButtonClick();

            //// Get Complete option
            //isim.Mouse.MoveMouseToPositionOnVirtualDesktop(15000, 29000);
            //Thread.Sleep(1000);
            //isim.Mouse.LeftButtonClick();

            //// Click Save & Close
            //isim.Mouse.MoveMouseToPositionOnVirtualDesktop(13000, 38000);
            //Thread.Sleep(1000);
            //isim.Mouse.LeftButtonClick();


            // Return to the Excel Sheet



        }

        static void copyExcelContents(InputSimulator isim)
        {
            SendKeys.SendWait("^a");
            Thread.Sleep(1000);
            SendKeys.SendWait("^v*{TAB}");
            Console.WriteLine("Pasting contents...");
            isim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        }

        static void closeProject(InputSimulator isim)
        {
            // Open project 
            isim.Mouse.MoveMouseToPositionOnVirtualDesktop(2000, 27050);
            Thread.Sleep(1000);
            isim.Mouse.LeftButtonDoubleClick();
            Thread.Sleep(1000);

            // Move mouse to Status dropdown
            isim.Mouse.MoveMouseToPositionOnVirtualDesktop(15000, 26000);
            Thread.Sleep(1000);
            isim.Mouse.LeftButtonClick();

            // Get Complete option
            isim.Mouse.MoveMouseToPositionOnVirtualDesktop(15000, 29000);
            Thread.Sleep(1000);
            isim.Mouse.LeftButtonClick();

            // Click Save & Close
            isim.Mouse.MoveMouseToPositionOnVirtualDesktop(13000, 38000);
            Thread.Sleep(1000);
            isim.Mouse.LeftButtonClick();
        }
    }
}