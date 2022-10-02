using System;
using System.ComponentModel;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LibUsbDotNet;
using LibUsbDotNet.Main;
using LibUsbDotNet.Info;
using System.Collections.ObjectModel;
using System.Net.Sockets;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;
using System.Collections.Generic;
using System.Threading;

namespace BarcodeScanner
{
    public partial class Main : Form
    {
        private const int myPID = 0xA002;  //產品ID
        private const int myVID = 0x065A;  //供應商ID
        private static UsbDeviceFinder MyUsbFinder;
        private static UsbDevice MyUsbDevice;
        private static IUsbDevice wholeUsbDevice;
        private TcpClient localTcp;
        private TcpClient EPTcp;
        Socket socketListen;
        Socket socketConnect;
        string RemoteEndPoint;    
        Dictionary<string, Socket> dicClient = new Dictionary<string, Socket>();
        Dictionary<string, Label> dicClientLabel = new Dictionary<string, Label>();
        bool IsSocketConnect = false;
        string PCN = "";
        bool workFinished = true;
        bool receiverPLCFinished = true;
        bool writerFinished = true;
        /* Main Event (Start)   */
        public Main()
        {
            InitializeComponent();
            StartSocket();
        }
        private void Main_Load(object sender, EventArgs e)
        {
            textBox_RMR1.Focus();
        }
        private void Main_Activated(object sender, EventArgs e)
        {
            textBox_RMR1.Focus();
        }
        /* Main Event (End)     */

        /* BackgroundWorker */
        private void Local_PLC_DoWork(object sender, DoWorkEventArgs e)
        {
            ConnectLocalPLC("192.168.1.100", 1026);
            while (true)
            {
                if (ScanOrNot())
                {
                    string barcode = BarcodeScan();
                    Thread.Sleep(200);
                    ScanSuccess();
                    Thread.Sleep(200);
                    string position = GetPosition();
                    Thread.Sleep(200);
                    string AAnum = AAnumber();
                    Thread.Sleep(200);
                    WriteOutDataView(position, AAnum, barcode);
                    while (!SendOrNot())
                    {
                        Thread.Sleep(300);
                    }
                    AsyncSend(dicClient[position]);
                    Thread.Sleep(200);
                    SendSuccess();
                }
                Thread.Sleep(200);
            }

        }
        private void EPPLC_DoWork(object sender, DoWorkEventArgs e)
        {
            ConnectEPPLC("192.168.1.110", 1025);
            int countIn = 0;
            int indexIn = 0;
            receiverPLCFinished = true;
            while (workFinished)
            {
                try
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        indexIn = dataGridViewIn.Rows.Count;
                    }));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                if (indexIn != countIn)
                {
                    string position = "", result = "";
                    for (int i = countIn; i < indexIn; i++)
                    {
                        this.Invoke(new MethodInvoker(delegate
                        {
                            position = dataGridViewIn.Rows[i].Cells[1].Value.ToString();
                            result = dataGridViewIn.Rows[i].Cells[5].Value.ToString();
                        }));
                        byte resultByte = 0x00;
                        if (result.Equals("OK"))
                        {
                            resultByte = 0x01;
                        }
                        else if (result.Equals("NG"))
                        {
                            resultByte = 0x02;
                        }
                        switch (position)
                        {
                            case "1":
                                WriteAA(resultByte, 0xF3);
                                break;
                            case "2":
                                WriteAA(resultByte, 0xF4);
                                break;
                            case "3":
                                WriteAA(resultByte, 0xF5);
                                break;
                            case "4":
                                WriteAA(resultByte, 0xF6);
                                break;
                            case "5":
                                WriteAA(resultByte, 0xF7);
                                break;
                            default:
                                MessageBox.Show("抓不到判斷AA機是哪台");
                                break;
                        }
                    }
                    countIn = indexIn;
                }
                Thread.Sleep(500);
            }
            MessageBox.Show("停止傳輸給收料機");
            receiverPLCFinished = false;
        }
        private void backgroundWorkerWriter_DoWork(object sender, DoWorkEventArgs e)
        {
            NewExcel();
            int countOut = 0;
            int countIn = 0;
            int indexOut = 0;
            int indexIn = 0;
            writerFinished = true;
            while (workFinished)
            {
                try
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        indexOut = dataGridViewOut.Rows.Count;
                    }));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                if (indexOut != countOut)
                {
                    WriteDataOut(countOut, indexOut);
                    countOut = indexOut;
                }
                try
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        indexIn = dataGridViewIn.Rows.Count;
                    }));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                if (indexIn != countIn)
                {
                    WriteDataIn(countIn, indexIn);
                    countIn = indexIn;
                }
            }
            MessageBox.Show("停止寫入excel");
            writerFinished = false;
        }

        /* BackgroundWorker (End)   */

        /* AA Server */
        public void StartSocket()
        {
            if (!IsSocketConnect)
            {
                try
                {
                   
                    IPEndPoint ipe = new IPEndPoint(IPAddress.Parse(textBox_ip.Text), int.Parse(textBox_port.Text));
                    socketListen = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                    socketListen.Bind(ipe);
                    socketListen.Listen(10);
                    //連線客戶端
                    AsyncConnect(socketListen);
                    IsSocketConnect = true;
                }
                catch (SocketException e)
                {
                    MessageBox.Show("Server無法監聽，請確認IP設置是否正確");
                    IsSocketConnect = false;
                    textBox_ip.Enabled = true;
                    textBox_port.Enabled = true;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    IsSocketConnect = false;
                }
            }
        }
        private void AsyncConnect(Socket socket)
        {
            try
            {
                socket.BeginAccept(asyncResult =>
                {
                    if (IsSocketConnect)
                    {
                        //receive info from clients
                        socketConnect = socket.EndAccept(asyncResult);
                        FirstReceive(socketConnect);
                        AsyncConnect(socketListen);
                    }
                }, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void FirstReceive(Socket socket)
        {

            byte[] data = new byte[4096];
            try
            {
                //start to receive info
                socket.BeginReceive(data, 0, data.Length, SocketFlags.None,
                asyncResult =>
                {
                    try
                    {
                        int length = socket.EndReceive(asyncResult);
                        DecodeData message = new DecodeData(data);
                        string position = message.getPosition();
                        this.Invoke(new MethodInvoker(delegate
                        {
                            Label[] AA_state = { label8, label9, label10, label11, label12 };
                            int label_num = Convert.ToInt32(position) - 1;
                            AA_state[label_num].Text = "連線";
                            AA_state[label_num].ForeColor = Color.Green;
                            dicClientLabel.Add(position, AA_state[label_num]);
                        }));
                        dicClient.Add(position, socket);
                        //setText(BitConverter.ToString(data).Replace("-", ""));
                    }
                    catch (Exception)
                    {
                        AsyncReceive(socket);
                    }
                    AsyncReceive(socket);
                }, null);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void AsyncReceive(Socket socket)
        {
            byte[] data = new byte[4096];
            try
            {
                //start to receive info
                socket.BeginReceive(data, 0, data.Length, SocketFlags.None,
                asyncResult =>
                {
                    try
                    {
                        int length = socket.EndReceive(asyncResult);
                        DecodeData message = new DecodeData(data);
                        WriteInDataView(message);
                        //setText(BitConverter.ToString(data).Replace("-", ""));
                    }
                    catch (Exception)
                    {
                        AsyncReceive(socket);
                    }
                    AsyncReceive(socket);
                }, null);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void AsyncSend(Socket client)
        {
            if (client == null) return;
            //encoding
            byte[] header = { 0x55, 0x00 };
            byte[] datetime = GetTime();
            String[] info = new String[4];
            int index = dataGridViewOut.Rows.Count - 1;
            for (int k = 0; k < 4; k++)
            {
                info[k] = dataGridViewOut.Rows[index].Cells[k + 1].Value.ToString();
            }
            byte[] Position = GetPosition(Convert.ToInt16(info[0]).ToString("0000"));
            byte[] LogNo = GetLogNo(Convert.ToInt16(info[1]).ToString("0000"));
            byte[] SendPCN = new byte[128];
            byte[] PCNByte = Encoding.ASCII.GetBytes(info[2]);
            Array.Copy(PCNByte, 0, SendPCN, 0, PCNByte.Length);
            byte[] SendSN = new byte[128];
            byte[] SNByte = Encoding.ASCII.GetBytes(info[3]);
            Array.Copy(SNByte, 0, SendSN, 0, SNByte.Length);
            byte[] data = new byte[4096];
            Array.Copy(header, 0, data, 0, header.Length);
            Array.Copy(datetime, 0, data, header.Length, datetime.Length);
            Array.Copy(Position, 0, data, header.Length + datetime.Length, Position.Length);
            Array.Copy(LogNo, 0, data, header.Length + datetime.Length + Position.Length, LogNo.Length);
            Array.Copy(SendPCN, 0, data, header.Length + datetime.Length + Position.Length + LogNo.Length, SendPCN.Length);
            Array.Copy(SendSN, 0, data, header.Length + datetime.Length + Position.Length + LogNo.Length + SendPCN.Length, SendSN.Length);
            try
            {
                client.BeginSend(data, 0, data.Length, SocketFlags.None, asyncResult =>
                {
                    int length = client.EndSend(asyncResult);
                }, null);
            }
            catch (Exception ex)
            {
                //failed and delete
                string deleteClient = client.RemoteEndPoint.ToString();
                dicClient.Remove(deleteClient);
                MessageBox.Show(ex.Message);
            }
        }
        private void AsyncSend1(Socket client)
        {
            if (client == null) return;
            //encoding
            byte[] header = { 0x55, 0x00 };
            byte[] datetime = GetTime();
            byte[] Position = GetPosition("0001");
            byte[] LogNo = GetLogNo("0002");
            byte[] SendPCN = new byte[128];
            byte[] PCNByte = Encoding.ASCII.GetBytes(PCN);
            Array.Copy(PCNByte, 0, SendPCN, 0, PCNByte.Length);
            Array.Reverse(SendPCN);
            byte[] SNByte = Encoding.ASCII.GetBytes("0123456789");
            Array.Reverse(SNByte);
            byte[] data = new byte[4096];
            Array.Copy(header, 0, data, 0, header.Length);
            Array.Copy(datetime, 0, data, header.Length, datetime.Length);
            Array.Copy(Position, 0, data, header.Length + datetime.Length, Position.Length);
            Array.Copy(LogNo, 0, data, header.Length + datetime.Length + Position.Length, LogNo.Length);
            Array.Copy(SendPCN, 0, data, header.Length + datetime.Length + Position.Length + LogNo.Length, SendPCN.Length);
            Array.Copy(SNByte, 0, data, header.Length + datetime.Length + Position.Length + LogNo.Length + SendPCN.Length, SNByte.Length);
            try
            {
                client.BeginSend(data, 0, data.Length, SocketFlags.None, asyncResult =>
                {
                    int length = client.EndSend(asyncResult);
                }, null);
            }
            catch (Exception ex)
            {
                string deleteClient = client.RemoteEndPoint.ToString();
                dicClient.Remove(deleteClient);
                MessageBox.Show(ex.Message);
            }
        }
        /* AA Server (End)*/

        /* Function     */
        private void ConnectLocalPLC(string hostName, int port)
        {
            localTcp = new TcpClient();
            try
            {
                if(localTcp.Connected == false)
                {
                    localTcp.Connect(hostName, port);
                    label_localPLC.ForeColor = Color.Green;
                    label_localPLC.Text = "連線";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                label_localPLC.ForeColor = Color.Red;
                label_localPLC.Text = "斷線";
            }
        }
        private void ConnectEPPLC(string hostName, int port)
        {
            EPTcp = new TcpClient();
            try
            {
                if(EPTcp.Connected == false)
                {
                    EPTcp.Connect(hostName, port);
                    label_EPPLC.ForeColor = Color.Green;
                    label_EPPLC.Text = "連線";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                label_EPPLC.ForeColor = Color.Red;
                label_EPPLC.Text = "斷線";
            }
        }
        private bool GetValidatedIP(string ipStr)
        {
            string validatedIP = string.Empty;
            IPAddress ip;
            if (IPAddress.TryParse(ipStr, out ip))
            {
                return true;
            }
            return false;
        }
        private string BarcodeScan()
        {
            //Open the device
            MyUsbFinder = new UsbDeviceFinder(myVID, myPID);
            MyUsbDevice = UsbDevice.OpenUsbDevice(MyUsbFinder);
            // If the device is open and ready  \0n
            if (MyUsbDevice == null)
            {
                MessageBox.Show("無此USB裝置");
                throw new Exception("Device Not Found.");
            }
            wholeUsbDevice = MyUsbDevice as IUsbDevice;
            if (!ReferenceEquals(wholeUsbDevice, null))
            {
                //This is a"whole" USB device. Before it can be used, 
                //the desired configuration and interface must be selected.
                //I think those do make some difference...
                //Select config #1
                wholeUsbDevice.SetConfiguration(1);
                //Claim interface #0.
                wholeUsbDevice.ClaimInterface(0);
            }
            //Open up the endpoints
            UsbEndpointWriter writer = MyUsbDevice.OpenEndpointWriter(WriteEndpointID.Ep01);
            UsbEndpointReader reader = MyUsbDevice.OpenEndpointReader(ReadEndpointID.Ep02);
            //Create a buffer with some data in it
            byte[] buffer = new byte[3];
            buffer[0] = 0x1B;
            buffer[1] = 0x5A;
            buffer[2] = 0x0D;
            //Write three bytes
            ErrorCode ec = ErrorCode.None;
            int bytesWritten;
            ec = writer.Write(buffer, 3000, out bytesWritten);
            if (ec != ErrorCode.None) throw new Exception(UsbDevice.LastErrorString);
            //Read some data 
            byte[] readBuffer = new byte[64];
            ec = reader.Read(readBuffer, 3000, out var readBytes);
            if (ec != ErrorCode.None) throw new Exception(UsbDevice.LastErrorString);
            // Write that output to the console.
            String barcode = Encoding.UTF8.GetString(readBuffer);
            barcode = barcode.Substring(0, 12);
            Console.WriteLine(barcode);
            return barcode;
        }
        private void WriteOutDataView(string position, string AA, string barcode)
        {
            this.Invoke(new MethodInvoker(delegate
            {
                DataGridViewRow dgvr = new DataGridViewRow();
                dgvr.CreateCells(dataGridViewOut);
                dgvr.Cells[0].Value = DateTime.Now;
                dgvr.Cells[1].Value = position;
                dgvr.Cells[2].Value = AA;
                dgvr.Cells[3].Value = PCN;
                dgvr.Cells[4].Value = barcode;
                dataGridViewOut.Rows.Add(dgvr);
                int i = dataGridViewOut.Rows.Count - 1;
                dataGridViewOut.CurrentCell = dataGridViewOut.Rows[i].Cells[0];
            }));
        }
        private void WriteInDataView(DecodeData data)
        {
            this.Invoke(new MethodInvoker(delegate
            {
                DataGridViewRow dgvr = new DataGridViewRow();
                dgvr.CreateCells(dataGridViewIn);
                dgvr.Cells[0].Value = data.getTime();
                dgvr.Cells[1].Value = data.getPosition();
                dgvr.Cells[2].Value = data.getLog();
                dgvr.Cells[3].Value = data.getPCN();
                dgvr.Cells[4].Value = data.getSN();
                dgvr.Cells[5].Value = data.getRlt();
                dgvr.Cells[6].Value = data.getNum();
                dgvr.Cells[7].Value = data.getRltsub();
                dataGridViewIn.Rows.Add(dgvr);
                int i = dataGridViewIn.Rows.Count - 1;
                dataGridViewIn.CurrentCell = dataGridViewIn.Rows[i].Cells[0];
            }));
        }
        private byte[] GetTime()
        { //certain spec
            DateTime dt = DateTime.Now;
            ushort year = ushort.Parse(dt.ToString("yyyy"));
            ushort month = ushort.Parse(dt.ToString("MM"));
            ushort week = ushort.Parse(dt.DayOfWeek.ToString("d"));
            ushort day = ushort.Parse(dt.ToString("dd"));
            ushort hour = ushort.Parse(dt.ToString("HH"));
            ushort min = ushort.Parse(dt.ToString("mm"));
            ushort sec = ushort.Parse(dt.ToString("ss"));
            ushort millsec = ushort.Parse(dt.ToString("fff"));
            byte[] yearByte = BitConverter.GetBytes(year);
            byte[] monthByte = BitConverter.GetBytes(month);
            byte[] weekByte = BitConverter.GetBytes(week);
            byte[] dayByte = BitConverter.GetBytes(day);
            byte[] hourByte = BitConverter.GetBytes(hour);
            byte[] minByte = BitConverter.GetBytes(min);
            byte[] secByte = BitConverter.GetBytes(sec);
            byte[] millsecByte = BitConverter.GetBytes(millsec);
            byte[] time = new byte[32];
            Array.Copy(yearByte, 0, time, 0, 2);
            Array.Copy(monthByte, 0, time, 2, 2);
            Array.Copy(weekByte, 0, time, 4, 2);
            Array.Copy(dayByte, 0, time, 6, 2);
            Array.Copy(hourByte, 0, time, 8, 2);
            Array.Copy(minByte, 0, time, 10, 2);
            Array.Copy(secByte, 0, time, 12, 2);
            Array.Copy(millsecByte, 0, time, 14, 2);
            Array.Copy(yearByte, 0, time, 16, 2);
            Array.Copy(monthByte, 0, time, 18, 2);
            Array.Copy(weekByte, 0, time, 20, 2);
            Array.Copy(dayByte, 0, time, 22, 2);
            Array.Copy(hourByte, 0, time, 24, 2);
            Array.Copy(minByte, 0, time, 26, 2);
            Array.Copy(secByte, 0, time, 28, 2);
            Array.Copy(millsecByte, 0, time, 30, 2);
            return time;
        }
        private byte[] GetPosition(string position)
        {
            ushort position1 = ushort.Parse(position.Substring(2, 2));
            ushort position2 = ushort.Parse(position.Substring(0, 2));
            byte[] p1 = BitConverter.GetBytes(position1);
            byte[] p2 = BitConverter.GetBytes(position2);
            byte[] p = new byte[4];
            Array.Copy(p1, 0, p, 0, 2);
            Array.Copy(p2, 0, p, 2, 2);
            return p;
        }
        private byte[] GetLogNo(string log)
        {
            ushort log1 = ushort.Parse(log.Substring(2, 2));
            ushort log2 = ushort.Parse(log.Substring(0, 2));
            byte[] l1 = BitConverter.GetBytes(log1);
            byte[] l2 = BitConverter.GetBytes(log2);
            byte[] l = new byte[4];
            Array.Copy(l1, 0, l, 0, 2);
            Array.Copy(l2, 0, l, 2, 2);
            return l;
        }
        private void NewExcel()
        {
            CreateFolder();
            try
            {
                string filePath = @"C:\BCS&MSS\" + PCN;
                if (!File.Exists(filePath + ".xlsx"))
                {
                    Excel.Workbook wBook;
                    Excel.Worksheet wSheet;
                    Excel.Application excelApp;
                    excelApp = new Excel.Application();
                    // tre to open workbook
                    excelApp.Workbooks.Add();
                    // excel property *****/
                    excelApp.Visible = false; 
                    excelApp.DisplayAlerts = false;
                    wBook = excelApp.Workbooks[1];
                    excelApp.Worksheets.Add();
                    wBook.Activate();
                    wSheet = wBook.Worksheets[1];
                    wSheet.Name = "Send";
                    wSheet.Activate();
                    string[] info = { "Time", "Position", "LogNo", "PCN", "SN" };
                    for (int i = 0; i < info.Length; i++)
                    {
                        excelApp.Cells[1, i + 1] = info[i];

                    }
                    wSheet = wBook.Worksheets[2];
                    wSheet.Name = "Receive";
                    wSheet.Activate();
                    string[] rltString = {"StattTime", "EndTime", "TBD1", "TBD2",
                    "Item","Result","Num","TestValue","Note" };
                    string[] rlt = new string[21*9];
                    for(int i = 0; i < 21; i++)
                    {
                        for(int j = 0; j < 9; j++)
                        {
                            rlt[9 * i + j] = rltString[j];
                        }
                    }
                    string[] info2 = { "Time", "Position", "LogNo", "PCN", "SN", "Result", "RltNum" };
                    string[] allString = new string[9*21+7];
                    Array.Copy(info2, 0, allString, 0, 7);
                    Array.Copy(rlt, 0, allString, 7, 9*21);
                    for (int i = 0; i < allString.Length; i++)
                    {
                        excelApp.Cells[1, i + 1] = allString[i];

                    }
                    
                    wBook.SaveAs(filePath);
                    wBook.Close(true);
                    excelApp.Quit(); 
                                     
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    wBook = null;
                    wSheet = null;
                    excelApp = null;
                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private void CreateFolder()
        {
            try
            {
                string filePath = @"C:\BCS&MSS";
                bool exists = System.IO.Directory.Exists(filePath);
                if (!exists)
                    Directory.CreateDirectory(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void WriteDataOut(int count, int index)
        {
            try
            {
                string filePath = @"C:\BCS&MSS\" + PCN;
                if (File.Exists(filePath + ".xlsx"))
                {
                    Excel.Workbook wBook;
                    Excel.Worksheet wSheet;
                    Excel.Range wRange;
                    Excel.Application excelApp;
                    excelApp = new Excel.Application();
                    
                    excelApp.Application.Workbooks.Open(filePath);
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;
                    wBook = excelApp.Workbooks[1];
                    wBook.Activate();
                    wSheet = wBook.Worksheets["Send"];
                    wSheet.Activate();
                    //last row
                    Excel.Range lastRow;
                    int lastUsedRow;
                    lastRow = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                    lastUsedRow = lastRow.Row + 1;
                    Console.WriteLine("Out row: " + lastUsedRow);
                    for (int j = count; j < index; j++)
                    {
                        string[] info = new string[5];
                        for (int k = 0; k < 5; k++)
                        {
                            if (k == 4)
                            {
                                info[k] = "'" + dataGridViewOut.Rows[j].Cells[k].Value.ToString();
                            }
                            else
                            {
                                info[k] = dataGridViewOut.Rows[j].Cells[k].Value.ToString();
                            }

                        }
                        for (int i = 0; i < info.Length; i++)
                        {
                            excelApp.Cells[lastUsedRow, i + 1] = info[i];

                        }
                        lastUsedRow++;
                    }
                    wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 2]];
                    wRange.Select();
                    wRange.Columns.AutoFit();
                    wBook.SaveAs(filePath);
                    wBook.Close(true);
                    excelApp.Quit();
                    //release
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    Console.WriteLine("end Out");
                    wBook = null;
                    wSheet = null;
                    wRange = null;
                    excelApp = null;
                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void WriteDataIn(int count, int index)
        {
            try
            {
                string filePath = @"C:\BCS&MSS\" + PCN;
                if (File.Exists(filePath + ".xlsx"))
                {
                    Excel.Workbook wBook;
                    Excel.Worksheet wSheet;
                    Excel.Range wRange;
                    Excel.Application excelApp;
                    excelApp = new Excel.Application();
                    excelApp.Application.Workbooks.Open(filePath);
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;
                    wBook = excelApp.Workbooks[1];
                    wBook.Activate();
                    wSheet = wBook.Worksheets["Receive"];
                    wSheet.Activate();
                    //last row
                    Excel.Range lastRow;
                    int lastUsedRow;
                    lastRow = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                    lastUsedRow = lastRow.Row + 1;
                    for (int j = count; j < index; j++)
                    {
                        String[] info = new String[7];
                        for (int k = 0; k < 7; k++)
                        {
                            info[k] = dataGridViewIn.Rows[j].Cells[k].Value.ToString();
                        }
                        for (int i = 0; i < info.Length; i++)
                        {
                            excelApp.Cells[lastUsedRow, i + 1] = info[i];
                        }
                        string[] rltsub = dataGridViewIn.Rows[j].Cells[7].Value.ToString().Split('?');
                        int rltsubIndex = 8;
                        foreach (string s in rltsub)
                        {
                            excelApp.Cells[lastUsedRow, rltsubIndex] = s;
                            rltsubIndex++;
                        }
                        lastUsedRow++;
                    }
                    wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 2]];
                    wRange.Select();
                    wRange.Columns.AutoFit();
                    wBook.SaveAs(filePath);
                    wBook.Close(true);
                    excelApp.Quit();
                    //release
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    Console.WriteLine("end In");
                    wBook = null;
                    wSheet = null;
                    wRange = null;
                    excelApp = null;
                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private bool ScanOrNot()
        {
            byte[] myBytes = { 0x50, 0x00, 0x00, 0xFF, 0xFF, 0x03, 0x00, 0x0C, 0x00, 0x00, 0x00, 0x01, 0x04, 0x00, 0x00, 0xF2, 0x03, 0x00, 0xA8, 0x01, 0x00 };
            localTcp.GetStream().Write(myBytes, 0, myBytes.Length);
            int bufferSize = localTcp.ReceiveBufferSize;
            Console.WriteLine("bufferSize: " + bufferSize);
            byte[] myBufferBytes = new byte[bufferSize];
            localTcp.GetStream().Read(myBufferBytes, 0, bufferSize);
            return BitConverter.ToInt32(myBufferBytes, 11) == 0 ? false : true;
        }
        private void ScanSuccess()
        {
            byte[] myBytes = { 0x50, 0x00,
            0x00,
            0xFF,
            0xFF, 0x03,
            0x00,
            0x0E, 0x00,
            0x00, 0x00, 0x01, 0x14, 0x00, 0x00, 0xF2, 0x03, 0x00, 0xA8, 0x01, 0x00, 0x00, 0x00 };
            localTcp.GetStream().Write(myBytes, 0, myBytes.Length);
            int bufferSize = localTcp.ReceiveBufferSize;
            byte[] myBufferBytes = new byte[bufferSize];
            localTcp.GetStream().Read(myBufferBytes, 0, bufferSize);
        }
        private string GetPosition()
        {
            byte[] myBytes = { 0x50, 0x00, 0x00, 0xFF, 0xFF, 0x03, 0x00, 0x0C, 0x00, 0x00, 0x00, 0x01, 0x04, 0x00, 0x00, 0xE8, 0x03, 0x00, 0xA8, 0x01, 0x00 };
            localTcp.GetStream().Write(myBytes, 0, myBytes.Length);
            int bufferSize = localTcp.ReceiveBufferSize;
            byte[] myBufferBytes = new byte[bufferSize];
            localTcp.GetStream().Read(myBufferBytes, 0, bufferSize);
            return BitConverter.ToInt32(myBufferBytes, 11).ToString();
        }
        private string AAnumber()
        {
            byte[] myBytes = { 0x50, 0x00, 0x00, 0xFF, 0xFF, 0x03, 0x00, 0x0C, 0x00, 0x00, 0x00, 0x01, 0x04, 0x00, 0x00, 0xEA, 0x03, 0x00, 0xA8, 0x01, 0x00 };
            localTcp.GetStream().Write(myBytes, 0, myBytes.Length);
            int bufferSize = localTcp.ReceiveBufferSize;
            byte[] myBufferBytes = new byte[bufferSize];
            localTcp.GetStream().Read(myBufferBytes, 0, bufferSize);
            return BitConverter.ToInt32(myBufferBytes, 11).ToString();
        }
        private bool SendOrNot()
        {
            byte[] myBytes = { 0x50, 0x00, 0x00, 0xFF, 0xFF, 0x03, 0x00, 0x0C, 0x00, 0x00, 0x00, 0x01, 0x04, 0x00, 0x00, 0xFC, 0x03, 0x00, 0xA8, 0x01, 0x00 };
            localTcp.GetStream().Write(myBytes, 0, myBytes.Length);
            int bufferSize = localTcp.ReceiveBufferSize;
            Console.WriteLine("bufferSize: " + bufferSize);
            byte[] myBufferBytes = new byte[bufferSize];
            localTcp.GetStream().Read(myBufferBytes, 0, bufferSize);
            return BitConverter.ToInt32(myBufferBytes, 11) == 0 ? false : true;
        }
        private void SendSuccess()
        {
            byte[] myBytes = { 0x50, 0x00,
            0x00,
            0xFF,
            0xFF, 0x03,
            0x00,
            0x0E, 0x00,
            0x00, 0x00, 0x01, 0x14, 0x00, 0x00, 0xFC, 0x03, 0x00, 0xA8, 0x01, 0x00, 0x00, 0x00 };
            localTcp.GetStream().Write(myBytes, 0, myBytes.Length);
            int bufferSize = localTcp.ReceiveBufferSize;
            byte[] myBufferBytes = new byte[bufferSize];
            localTcp.GetStream().Read(myBufferBytes, 0, bufferSize);
        }
        private void WriteAA(byte result, byte position)
        {
            byte[] myBytes = { 0x50, 0x00,
            0x00,
            0xFF,
            0xFF, 0x03,
            0x00,
            0x0E, 0x00,
            0x00, 0x00, 0x01, 0x14, 0x00, 0x00, position, 0x03, 0x00, 0xA8, 0x01, 0x00, result , 0x00};
            EPTcp.GetStream().Write(myBytes, 0, myBytes.Length);
            int bufferSize = EPTcp.ReceiveBufferSize;
            Console.WriteLine("bufferSize: " + bufferSize);
            byte[] myBufferBytes = new byte[bufferSize];
            EPTcp.GetStream().Read(myBufferBytes, 0, bufferSize);
        }

        private void CleanDataView()
        {
            dataGridViewOut.DataSource = null;
            dataGridViewOut.Rows.Clear();
            dataGridViewIn.DataSource = null;
            dataGridViewIn.Rows.Clear();
        }
        /* Function (End)*/

        /*  Button Zone */
        private void StartButton_Click(object sender, EventArgs e)
        {
            if (!textBox_RMR1.Text.Equals(""))
            {
                if (!(textBox_ip.Text.Equals("") || textBox_port.Text.Equals("")))
                {
                    if (GetValidatedIP(textBox_ip.Text))
                    {
                        PCN = textBox_RMR1.Text + ";" + textBox_AFM.Text;
                        backgroundWorkerWriter.RunWorkerAsync();
                        workFinished = true;
                        if (!IsSocketConnect)
                        {
                            StartSocket();
                        }
                        Local_PLC.RunWorkerAsync();
                        EPPLC.RunWorkerAsync();
//Loadbutton.Enabled = false;
                        StartButton.Enabled = false;
                        buttonChangeIP.Enabled = false;
                        StopButton.Enabled = true;
                    }
                    else
                    {
                        MessageBox.Show("請輸入正確的IP number");
                    }
                }
                else
                {
                    MessageBox.Show("請先輸入IP及Port number");
                }
            }
            else
            {
                MessageBox.Show("請先輸入PCN Code");
            }
        }
        private void StopButton_Click(object sender, EventArgs e)
        {
            workFinished = false;
            //while(writerFinished || receiverPLCFinished)
            //{
            //    Thread.Sleep(200);
            //}
            CleanDataView();
//Loadbutton.Enabled = true;
            StartButton.Enabled = true;
            buttonChangeIP.Enabled = true;
            textBox_RMR1.Text = "";
            textBox_AFM.Text = "";
            Thread.Sleep(500);
            backgroundWorkerWriter.CancelAsync();
            EPPLC.CancelAsync();
        }
        private void buttonChangeIP_Click(object sender, EventArgs e)
        {
            IsSocketConnect = false;
            socketListen.Close();
            textBox_ip.Enabled = true;
            textBox_port.Enabled = true;
        }
        private void buttonTest_Click(object sender, EventArgs e)
        {
//Form f = new test();
            //f.Visible = true;
        }
        private void buttonRMR1_Click(object sender, EventArgs e)
        {
            textBox_RMR1.Text = BarcodeScan();
        }
        private void buttonAFM_Click(object sender, EventArgs e)
        {
            textBox_AFM.Text = BarcodeScan();
        }
        /*  Button Zone (End) */

        /* textbox event */
        private void textBox_port_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
                e.Handled = true;
        }


        /* textbox event (End)*/


        /* PLC Function*/
        private void buttonWrtieD_Click_1(object sender, EventArgs e)
        {
            String a = "00";
            String b = "00";
            this.Invoke(new MethodInvoker(delegate
            {
                //button1.PerformClick();
                //a = textBox1.Text.Substring(0, 2);
                //b = textBox1.Text.Substring(2, 2);
            }));
            byte c = Convert.ToByte(a, 16);
            byte d = Convert.ToByte(b, 16);
            byte[] myBytes = { 0x50, 0x00,
            0x00,
            0xFF,
            0xFF, 0x03,
            0x00,
            0x0E, 0x00,
            0x00, 0x00, 0x01, 0x14, 0x00, 0x00, 0xDE, 0x00, 0x00, 0xA8, 0x01, 0x00, c, d };
            localTcp.GetStream().Write(myBytes, 0, myBytes.Length);
            int bufferSize = localTcp.ReceiveBufferSize;
            Console.WriteLine("bufferSize: " + bufferSize);
            byte[] myBufferBytes = new byte[bufferSize];
            localTcp.GetStream().Read(myBufferBytes, 0, bufferSize);
            //encoding
            int s = BitConverter.ToInt32(myBufferBytes, 9);
            if (s == 0)
            {
                MessageBox.Show("寫入成功");
            }
            else
            {
                MessageBox.Show("寫入失敗");
            }

        }
        private void buttonReadD_Click_1(object sender, EventArgs e)
        {
            byte[] myBytes = { 0x50, 0x00, 0x00, 0xFF, 0xFF, 0x03, 0x00, 0x0C, 0x00, 0x00, 0x00, 0x01, 0x04, 0x00, 0x00, 0xDE, 0x00, 0x00, 0xA8, 0x01, 0x00 };

            localTcp.GetStream().Write(myBytes, 0, myBytes.Length);
            int bufferSize = localTcp.ReceiveBufferSize;
            Console.WriteLine("bufferSize: " + bufferSize);
            byte[] myBufferBytes = new byte[bufferSize];
            localTcp.GetStream().Read(myBufferBytes, 0, bufferSize);
            MessageBox.Show(BitConverter.ToInt32(myBufferBytes, 11).ToString());
            //for (int i = 0; i < 20; i++)
            //{
            //    WriteIn(i.ToString(), BitConverter.ToInt32(myBufferBytes, i).ToString());
            //}
        }
        private void buttonConnect_Click_1(object sender, EventArgs e)
        {
            string hostName = "192.168.1.100";
            int connectPort = 1026;
            localTcp = new TcpClient();
            try
            {
                localTcp.Connect(hostName, connectPort);
            }
            catch (Exception ex)
            {
                Console.WriteLine("主機 {0} 通訊埠 {1} 無法連接  !!", hostName, connectPort);
                MessageBox.Show(ex.Message);
            }
        }
        private void buttonOnM_Click(object sender, EventArgs e)
        {   
            byte[] myBytes = { 0x50, 0x00, 0x00, 0xFF, 0xFF, 0x03, 0x00, 0x0E, 0x00, 0x00, 0x00, 0x01, 0x14, 0x01, 0x00, 0x64, 0x00, 0x00, 0x90, 0x04, 0x00, 0x11, 0x11 };
            localTcp.GetStream().Write(myBytes, 0, myBytes.Length);
            int bufferSize = localTcp.ReceiveBufferSize;
            Console.WriteLine("bufferSize: " + bufferSize);
            byte[] myBufferBytes = new byte[bufferSize];
            localTcp.GetStream().Read(myBufferBytes, 0, bufferSize);
        }
        private void buttonOffM_Click_1(object sender, EventArgs e)
        {
            byte[] myBytes = { 0x50, 0x00,
            0x00,
            0xFF,
            0xFF, 0x03,
            0x00,
            0x0E, 0x00,
            0x00, 0x00, 0x01, 0x14, 0x01, 0x00, 0x64, 0x00, 0x00, 0x90, 0x04, 0x00, 0x00, 0x00};
            localTcp.GetStream().Write(myBytes, 0, myBytes.Length);
            int bufferSize = localTcp.ReceiveBufferSize;
            Console.WriteLine("bufferSize: " + bufferSize);
            byte[] myBufferBytes = new byte[bufferSize];
            localTcp.GetStream().Read(myBufferBytes, 0, bufferSize);
        }
        private void buttonLookM_Click(object sender, EventArgs e)
        {
            byte[] myBytes = { 0x50, 0x00, 0x00, 0xFF, 0xFF, 0x03, 0x00, 0x0C, 0x00, 0x00, 0x00, 0x01, 0x04, 0x00, 0x00, 0x64, 0x00, 0x00, 0x90, 0x04, 0x00 };
            localTcp.GetStream().Write(myBytes, 0, myBytes.Length);
            int bufferSize = localTcp.ReceiveBufferSize;
            Console.WriteLine("bufferSize: " + bufferSize);
            byte[] myBufferBytes = new byte[bufferSize];
            localTcp.GetStream().Read(myBufferBytes, 0, bufferSize);
            //label15.Text = BitConverter.ToInt32(myBufferBytes, 11) == 0 ? "0" : "1";
            //for (int i = 0; i < 20; i++)
            //{
            //    WriteIn(i.ToString(), BitConverter.ToInt32(myBufferBytes, i).ToString());
            //}
        }

        private void textBox_RMR1_TextChanged(object sender, EventArgs e)
        {

        }



        /* PLC Function*/
    }
    class DecodeData
    {
        private byte[] data = new byte[4096];
        private int rltNum = 0;
        public DecodeData()
        {


        }
        public DecodeData(byte[] data)
        {
            this.data = data;
        }
        public string getTime()
        {
            byte[] time = new byte[16];
            Array.Copy(this.data, 2, time, 0, 16);
            ushort year = BitConverter.ToUInt16(time, 0);
            ushort month = BitConverter.ToUInt16(time, 2);
            ushort week = BitConverter.ToUInt16(time, 4);
            ushort day = BitConverter.ToUInt16(time, 6);
            ushort hour = BitConverter.ToUInt16(time, 8);
            ushort min = BitConverter.ToUInt16(time, 10);
            ushort sec = BitConverter.ToUInt16(time, 12);
            ushort millsec = BitConverter.ToUInt16(time, 14);
            return (year + "/" + month.ToString("00") + "/" + day + "/" + week + " " + hour.ToString("00") + ":" + min.ToString("00") + ":" + sec.ToString("00") + " " + millsec.ToString("000"));
        }
        public string getPosition()
        {
            byte[] position = new byte[4];
            Array.Copy(this.data, 34, position, 0, 4);
            ushort location1 = BitConverter.ToUInt16(position, 0);
            ushort location2 = BitConverter.ToUInt16(position, 2);
            int location = location2 * 100 + location1;
            return location.ToString();
        }
        public int getLog()
        {
            byte[] log = new byte[4];
            Array.Copy(this.data, 38, log, 0, 4);
            ushort logNo1 = BitConverter.ToUInt16(log, 0);
            ushort logNo2 = BitConverter.ToUInt16(log, 2);
            int logNo = logNo2 * 100 + logNo1;
            return logNo;
        }
        public string getPCN()
        {
            byte[] PCN = new byte[128];
            Array.Copy(this.data, 42, PCN, 0, 128);
            return Encoding.ASCII.GetString(PCN);
        }
        public string getSN()
        {
            byte[] SN = new byte[128];
            Array.Copy(this.data, 170, SN, 0, 128);
            return Encoding.ASCII.GetString(SN);
        }
        public string getRlt()
        {
            byte[] rlt = new byte[4];
            Array.Copy(this.data, 298, rlt, 0, 4);
            ushort rlt1 = BitConverter.ToUInt16(rlt, 0);
            ushort rlt2 = BitConverter.ToUInt16(rlt, 2);
            int r = rlt2 * 100 + rlt1;
            if (r <= 0)
            {
                return "";
            }
            else if (r <= 15)
            {
                return "NG";
            }
            else if (r == 16)
            {
                return "OK";
            }
            return "";
        }
        public int getNum()
        {
            byte[] num = new byte[4];
            Array.Copy(this.data, 302, num, 0, 4);
            ushort m1 = BitConverter.ToUInt16(num, 0);
            ushort m2 = BitConverter.ToUInt16(num, 2);
            int m = m2 * 100 + m1;
            this.rltNum = m;
            return m;
        }
        private string getStartTime(int start)
        {
            byte[] time = new byte[16];
            Array.Copy(this.data, start, time, 0, 16);
            ushort year = BitConverter.ToUInt16(time, 0);
            ushort month = BitConverter.ToUInt16(time, 2);
            ushort week = BitConverter.ToUInt16(time, 4);
            ushort day = BitConverter.ToUInt16(time, 6);
            ushort hour = BitConverter.ToUInt16(time, 8);
            ushort min = BitConverter.ToUInt16(time, 10);
            ushort sec = BitConverter.ToUInt16(time, 12);
            ushort millsec = BitConverter.ToUInt16(time, 14);
            return (year + "/" + month.ToString("00") + "/" + day + "/" + week + "," + hour.ToString("00") + ":" + min.ToString("00") + ":" + sec.ToString("00") + "," + millsec.ToString("000"));
        }
        private string getEndTime(int start)
        {
            byte[] time = new byte[16];
            Array.Copy(this.data, start, time, 0, 16);
            ushort year = BitConverter.ToUInt16(time, 0);
            ushort month = BitConverter.ToUInt16(time, 2);
            ushort week = BitConverter.ToUInt16(time, 4);
            ushort day = BitConverter.ToUInt16(time, 6);
            ushort hour = BitConverter.ToUInt16(time, 8);
            ushort min = BitConverter.ToUInt16(time, 10);
            ushort sec = BitConverter.ToUInt16(time, 12);
            ushort millsec = BitConverter.ToUInt16(time, 14);
            return (year + "/" + month.ToString("00") + "/" + day + "/" + week + "," + hour.ToString("00") + ":" + min.ToString("00") + ":" + sec.ToString("00") + "," + millsec.ToString("000"));
        }
        private int getTBD1(int start)
        {
            byte[] TBD = new byte[4];
            Array.Copy(this.data, start, TBD, 0, 4);
            ushort t1 = BitConverter.ToUInt16(TBD, 0);
            ushort t2 = BitConverter.ToUInt16(TBD, 2);
            int TBD1 = t2 * 100 + t1;
            return TBD1;
        }
        private int getTBD2(int start)
        {
            byte[] TBD = new byte[4];
            Array.Copy(this.data, start, TBD, 0, 4);
            ushort t1 = BitConverter.ToUInt16(TBD, 0);
            ushort t2 = BitConverter.ToUInt16(TBD, 2);
            int TBD2 = t2 * 100 + t1;
            return TBD2;
        }
        private int getItem(int start)
        {
            byte[] Item = new byte[4];
            Array.Copy(this.data, start, Item, 0, 4);
            ushort i1 = BitConverter.ToUInt16(Item, 0);
            ushort i2 = BitConverter.ToUInt16(Item, 2);
            int i = i2 * 100 + i1;
            return i;
        }
        private string getResult(int start)
        {
            byte[] Result = new byte[4];
            Array.Copy(this.data, start, Result, 0, 4);
            ushort i1 = BitConverter.ToUInt16(Result, 0);
            ushort i2 = BitConverter.ToUInt16(Result, 2);
            int i = i2 * 100 + i1;
            if (i <= 0)
            {
                return "";
            }
            else if (i <= 15)
            {
                return "NG";
            }
            else if (i == 16)
            {
                return "OK";
            }
            return "";
        }
        private int getN(int start)
        {
            byte[] num = new byte[4];
            Array.Copy(this.data, start, num, 0, 4);
            ushort m1 = BitConverter.ToUInt16(num, 0);
            ushort m2 = BitConverter.ToUInt16(num, 2);
            int m = m2 * 100 + m1;
            return m;
        }
        private string getTestValue(int start, int end)
        {
            int num = end;
            byte[] rSub = new byte[num];
            Array.Copy(this.data, start, rSub, 0, num);
            //Array.Reverse(Note);
            return BitConverter.ToString(rSub).Replace("-", "");
        }
        private string getNote(int start)
        {
            byte[] Note = new byte[128];
            Array.Copy(this.data, start, Note, 0, 128);
            return Encoding.ASCII.GetString(Note);
        }

        public string getRltsub()
        {
            int rltNum = this.rltNum;
            string startTime, endTime, result, testValue, note;
            int TBD1, TBD2, item, testNum;
            int dataIndex = 306;
            string rltsub = "";
            try
            {
                for (int i = 0; i < rltNum; i++)
                {
                startTime = getStartTime(dataIndex);
                dataIndex = dataIndex + 16;
                endTime = getEndTime(dataIndex);
                dataIndex = dataIndex + 16;
                TBD1 = getTBD1(dataIndex);
                dataIndex = dataIndex + 4;
                TBD2 = getTBD2(dataIndex);
                dataIndex = dataIndex + 4;
                item = getItem(dataIndex);
                dataIndex = dataIndex + 4;
                result = getResult(dataIndex);
                dataIndex = dataIndex + 4;
                testNum = getN(dataIndex);
                dataIndex = dataIndex + 4;
                testValue = getTestValue(dataIndex, dataIndex+testNum);
                dataIndex = dataIndex + testNum ;
                note = getNote(dataIndex);
                dataIndex = dataIndex + 128;
                rltsub = rltsub + startTime + "?" + endTime + "?" + TBD1.ToString() + "?" + TBD2.ToString() + "?" + item.ToString() + "?" + result + "?" + testNum.ToString() + "?" + testValue + "?" + note + "?";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return rltsub;
        }
    }
}