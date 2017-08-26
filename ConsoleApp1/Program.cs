using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using IWshRuntimeLibrary;
using System.Windows.Automation;
using InTheHand;
using InTheHand.Net.Bluetooth;
using InTheHand.Net.Ports;
using InTheHand.Net.Sockets;
using NDde.Client;
using System.Text;
using MimeTypes;
namespace ConsoleApp1
{
    class Program
    {


        static void Main(string[] args)
        {
            
            BluetoothServer bs = new BluetoothServer();
            


          //  var path = bs.getFilePath();
           // Console.WriteLine(path);
        }
        private class BluetoothServer{
            Guid muid = new Guid("{28931028-9310-0000-0000-289310289310}");
            private Boolean listening;
            BluetoothRadio btRadio = BluetoothRadio.PrimaryRadio;
            BluetoothListener btListener;
            BluetoothClient btClient;
            Stream ns;

            public BluetoothServer()
            {
                connectAsServer();
            }
            public void closeConnection()
            {
                try
                {
                    Console.WriteLine("Closing BluetoothClient");
                    btClient.Close();
                }
                catch
                {
                    Console.WriteLine("Error closing bluetoothclient");
                }
            }

            public void connectAsServer()
            {
                if (btRadio == null)
                {
                    Console.WriteLine("Bluetooth not supported");
                    return;
                }
                if (btRadio.Mode != RadioMode.Discoverable)
                    btRadio.Mode = RadioMode.Discoverable;              
                Console.WriteLine("BT device name: "+btRadio.Name.ToString());
                btListener = new BluetoothListener(muid);
                btListener.ServiceName = "Tranny";
                Console.WriteLine(btListener.ServiceName);
                btListener.Start();
                listening = true;
                //  Thread t = new Thread(new ThreadStart(listenLoop));
                listenLoop();
            }

            private void listenLoop()
            {
                
                while (listening)
                {
                    
                    try
                    {
                        Console.WriteLine("ListenLoop activated ");
                        btClient  = btListener.AcceptBluetoothClient();
                        
                        if (btClient.Connected)
                        {
                            Console.WriteLine("Connected to " + btClient.RemoteMachineName);
                        }
                        ns = btClient.GetStream();
                        if (ns != null && ns.CanWrite)
                        {
                            var path = getFilePath();
                            if (path != null && path.Contains("http"))
                            {
                                Console.WriteLine(path);

                                Byte[] bytes = new Byte[path.Length * sizeof(char)];
                                bytes = Encoding.UTF8.GetBytes(path);
                                Console.WriteLine(bytes.Length);
                                ns.Write(bytes, 0, bytes.Length);

                                //  closeConnection();
                                //   connDev = bc.RemoteMachineName;
                                //  Thread.Sleep(10000);
                            }
                            else if (path != null)
                            {
                                Stopwatch sw = new Stopwatch();
                                sw.Start();

                                Console.WriteLine(path);
                                int tmp = path.LastIndexOf('\\');
                                string fileName = path.Substring(tmp + 1);
                                Byte[] bytes = new Byte[fileName.Length * sizeof(char)];

                                /* string t = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+"\\TrannyTemp";
                                 if (!Directory.Exists(t))
                                 {
                                     Directory.CreateDirectory(t);
                                 }
                                 t = Path.Combine(t, getFileName(path));
                                 if (!System.IO.File.Exists(t))
                                 {
                                     System.IO.File.Create(t);
                                 }

                                 Console.WriteLine("t = " + t);
                                 System.IO.File.Copy(path, t, true);
                                 */

                                //byte[] outBuffer = System.IO.File.ReadAllBytes(path);

                                Stream s = new FileStream(path,
                                 FileMode.Open,
                                 FileAccess.Read,
                                 FileShare.ReadWrite);
                                var outBuffer = ReadFully(s);


                                Console.WriteLine(outBuffer.Length);
                                fileName = fileName + "---" + outBuffer.Length;
                                bytes = Encoding.UTF8.GetBytes(fileName);
                                Console.WriteLine(bytes.Length);
                                ns.Write(bytes, 0, bytes.Length);
                                Console.WriteLine(fileName);
                                ns.ReadTimeout = 2000;
                                if (!ns.CanRead)
                                {
                                    Console.WriteLine("Unable to read Stream");
                                }
                                else
                                {
                                    Console.WriteLine("Attempting to read");
                                    try
                                    {
                                        int resp = ns.ReadByte();
                                        Console.WriteLine("Response: " + resp);
                                        if (resp == 255)
                                        {

                                            for (int os = 0; os < outBuffer.Length; os += 1228800)
                                            {
                                                if (outBuffer.Length - os < 1228800)
                                                {
                                                    ns.Write(outBuffer, os, outBuffer.Length - os);
                                                }
                                                else
                                                    ns.Write(outBuffer, os, 1228800);
                                                Console.WriteLine(os);
                                            }

                                            Console.WriteLine("Transfer Finished in " + sw.Elapsed.ToString());
                                            ns.Close();
                                            sw.Stop();
                                           
                                        }
                                        else
                                        {
                                            Console.WriteLine("Not sending File");
                                            ns.Close();
                                        }
                                    }
                                    catch(Exception e)
                                    {
                                        
                                        Console.WriteLine(e.StackTrace);

                                        ns.Close();
                                    }
                                    
                                }
                            }          
                            }
                        
                    }
                    catch(Exception e)
                    {
                        printStackTrace(e);
                        if (e.Message.Contains("The process cannot access the file"))
                        {
                            
                        }
                        else
                            break;
                
                    }
                   
                }
                
            }

            public string GetChromeURL(Process proc)
            {
                AutomationElement elm = AutomationElement.FromHandle(proc.MainWindowHandle);
                
                string name = elm.GetCurrentPropertyValue(AutomationElement.NameProperty) as string;
                
                AutomationElement elmUrlBar = elm.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, "Address and search bar"));

                
                if (elmUrlBar != null)
                {
                    AutomationPattern[] patterns = elmUrlBar.GetSupportedPatterns();
                    if (patterns.Length > 0)
                    {
                        ValuePattern val = (ValuePattern)elmUrlBar.GetCurrentPattern(patterns[0]);
                        String url = val.Current.Value;
                        if (!url.Contains("http"))
                        {
                            url = "http://" + url;
                        }
                        return url;
                    }
                }
                return null;
            }
            public string GetFirefoxURL()
            {
                DdeClient dde = new DdeClient("firefox", "WWW_GetWindowInfo");
                dde.Connect();
                string url = dde.Request("URL", int.MaxValue);
                dde.Disconnect();
                if (!url.Equals(""))
                {
                    var split = url.Split('"');

                    return split[1].Trim();
                }
                return null;
            }
            public string getFilePath()
            {
               // Thread.Sleep(3000); // Test it with 5 Seconds, set a window to foreground, and you see it works!
                IntPtr hWnd = GetForegroundWindow();
                uint procId = 0;
                GetWindowThreadProcessId(hWnd, out procId);
                var proc = Process.GetProcessById((int)procId);
                Console.WriteLine(proc.ProcessName);
                if (proc.ProcessName.Contains("firefox"))
                    return GetFirefoxURL();
                else if (proc.ProcessName.Contains("chrome"))
                    return GetChromeURL(proc);

                string path = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
                var files = Directory.EnumerateFiles(path);
                WshShell shell = new WshShell();
                Console.WriteLine(proc.MainWindowTitle);
                string title = getImptTitle(proc);
                Console.WriteLine(title);
                foreach (var i in files)
                {
                    String[] s = i.Split('.');
                    s = s[0].Split('\\');
                    //Console.WriteLine("s[last]=" + s[]);
                    if (!i.Contains(".lnk") && !i.Contains(".url"))
                        continue;
                    else if (s[s.Length-1].Equals(title))
                    {
                        //Console.WriteLine(i);
                           
                        IWshShortcut link = (IWshShortcut)shell.CreateShortcut(i);
                        Console.WriteLine("TargetPath="+link.TargetPath);
                        try
                        {
                            var executable = FindExecutable(link.TargetPath);
                            Console.WriteLine("Executable = " + executable);
                            executable = proc.MainModule.FileName;
                            Console.WriteLine("Mainmodule Executable = " + executable);
                        }
                        catch(Exception e)
                        {
                            printStackTrace(e);
                            
                        }
                        return link.TargetPath;
                    }
                }
                return null;
            }
            public byte[] ReadFully(Stream input)
            {
                byte[] buffer = new byte[16 * 1024];
                using (MemoryStream ms = new MemoryStream())
                {
                    int read;
                    while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        ms.Write(buffer, 0, read);
                    }
                    return ms.ToArray();
                }
            }


            public string getFileName(string path)
            {
                int t = path.LastIndexOf('\\');
                var filename = path.Substring(t + 1);
                return filename;
            }
            public void printStackTrace(Exception e)
            {

                Console.WriteLine(e.Source);
                Console.WriteLine(e.TargetSite);
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);

            }

            private string getImptTitle(Process proc)
            {
                var windowTitle = proc.MainWindowTitle;
                Console.WriteLine(windowTitle);
                int idx=0;
                if(windowTitle.Contains("-"))
                     idx = windowTitle.LastIndexOf('-');
                Console.WriteLine(idx);
                string title=windowTitle;
                if(idx!=0)
                    title = windowTitle.Substring(0, idx);
                //windowTitle.LastIndexOf(new string[] {" - "}, StringSplitOptions.None);
                if (title.Contains("."))
                {
                    var splitTitle = title.Split('.');
                    return splitTitle[0].Trim();
                }

                return title.Trim();
            }

            public bool HasExecutable(string path)
            {
                var executable = FindExecutable(path);
                return !string.IsNullOrEmpty(executable);
            }

            private string FindExecutable(string path)
            {
                var executable = new StringBuilder(1024);
                FindExecutable(path, string.Empty, executable);
                return executable.ToString();
            }

            [DllImport("shell32.dll", EntryPoint = "FindExecutable")]
            private static extern long FindExecutable(string lpFile, string lpDirectory, StringBuilder lpResult);

        }

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    }
    
}
