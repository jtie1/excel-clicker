using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Text;
using WindowsInput;
using WindowsInput.Native;


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
            InputSimulator isim = new InputSimulator();         // Used bc SendKeys.SendWait() doesn't always work

            // Get the Kronos browser
            Process[] kr = Process.GetProcessesByName("msedge");
            Process kronosProcess = kr.FirstOrDefault();

            // Check that both Excel and Edge are open
            if (excelProcess == null)
            {
                Console.WriteLine("Couldn't find the Excel application.");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                Environment.Exit(1);
            }

            if (kronosProcess == null)
            {
                Console.WriteLine("Couldn't find an open Microsoft Edge window.");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                Environment.Exit(1);
            }

            // Close all projects
            //closeAllProjects(excelProcess, kronosProcess, isim);
            close2ndLineProjects(excelProcess, kronosProcess, isim);
            
        }

        static void closeAllProjects(Process excelProcess, Process kronosProcess, InputSimulator isim)
        {
            for (int i = 0; i < 30; i++)
            {
                // Bring Excel into focus
                Console.WriteLine("Bringing Excel into focus");
                _bringIntoFocus(excelProcess);

                // Get the next item in the list
                SendKeys.SendWait("{DOWN}");
                Thread.Sleep(1000);

                // Copy
                _copyContents();

                // Bring Edge into focus
                Console.WriteLine("Bringing Microsoft Edge into focus");
                _bringIntoFocus(kronosProcess);

                // Paste
                _pasteExcelContents(isim);

                // Close
                _closeProject(isim);
            }
        }

        // Some projects have multiple function codes. This takes care of the second function code
        // TODO: hard coded to positions on the screen. see if it can be dynamic
        static void close2ndLineProjects(Process excelProcess, Process kronosProcess, InputSimulator isim)
        {
            for (int i = 0; i < 5; i++)
            {
                // Bring Excel into focus
                Console.WriteLine("Bringing Excel into focus");
                _bringIntoFocus(excelProcess);

                // Get the next item in the list
                SendKeys.SendWait("{DOWN}");
                Thread.Sleep(1000);

                // Copy
                _copyContents();

                // Bring Edge into focus
                Console.WriteLine("Bringing Microsoft Edge into focus");
                _bringIntoFocus(kronosProcess);

                // Paste
                _pasteExcelContents(isim);

                // Close
                _close2ndLine(isim);
            }
        }

        // Bring the application into focus before taking further actions
        private static void _bringIntoFocus(Process process)
        {
            IntPtr p = process.MainWindowHandle;
            SetForegroundWindow(p);
            Thread.Sleep(1000);
        }

        // Copy contents from Excel
        private static void _copyContents()
        {
            // Copy contents. For some reason ^C, ^{C}, and ^(C) don't work. Guess it's not capital C, maybe because default is lowercase c?
            SendKeys.SendWait("^c");
            Console.WriteLine("Copying contents...");
        }

        // Paste contents into search box. Prerequisite: Activity Name selected
        // Weird error: instead of the project name being copied, `SendKeys.SendWait("^v");*` was copied instead lol
        private static void _pasteExcelContents(InputSimulator isim)
        {
            SendKeys.SendWait("^a");
            Thread.Sleep(1000);
            SendKeys.SendWait("^v*"); // removed {TAB}
            Console.WriteLine("Pasting contents...");
            isim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        }

        /********* TODO: Hard coded to leftmost monitor on in-office setup, try to find HTML element *********/
        private static void _closeProject(InputSimulator isim)
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

        private static void _close2ndLine(InputSimulator isim)
        {
            // Open project 
            isim.Mouse.MoveMouseToPositionOnVirtualDesktop(2000, 28750);
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