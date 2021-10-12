using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.Security.Permissions;
using Microsoft.Win32;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Linq;
using System.Diagnostics;

namespace decrypter
{
    

    public partial class Form1 : Form
    {

        private string anotherKey = "abcdefghijklmnop";
        private string filePath = ""; // empty string
        private bool stuff = true;


        public Form1()
        {
            InitializeComponent();


            button1.Click += new EventHandler(button1_Click);

            button2.Click += new EventHandler(button2_Click);

            button3.Click += new EventHandler(button3_Click);

            
            
        }


        private void decryptStuffAsync()
        {

        }
        // INFO https://www.youtube.com/watch?v=2moh18sh5p4
        // File decryption using Asynchronous programming
        private async Task button1_ClickAsync()
        {
            // decrypt_file button

            var sw = new Stopwatch();
            sw.Start();

            textBox2.Text = "";
            textBox4.Text = "Processing. . . . .";
             

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            // Excel.Range range;

            xlWorkBook = xlApp.Workbooks.Open(filePath);
            xlWorkSheet = xlWorkBook.Worksheets[1];

            //in what row and columnm to begin decrypting on
            int i = 2; //row
            int j = 2; //column

            string decryptedData = "";


            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook newXLWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet newXLWorksheet = (Excel.Worksheet)newXLWorkbook.Sheets.Add();

            //in what row and column to where to place the decypted file on the new excel file
            int row = 1;
            int column = 1;

            
            do
            {
                if (xlWorkSheet.Cells[i, j].Text != "NULL")
                {
                    decryptedData = xlWorkSheet.Cells[i, j].Value2;
                    decryptedData = await Task.Run(() => DecryptString(decryptedData));
                    newXLWorksheet.Cells[row, column] = decryptedData;
                    row++;
                }
                else
                {
                    decryptedData = "NULL";
                    newXLWorksheet.Cells[row, column] = decryptedData;
                    row++;
                }


                i++;

            } while (xlWorkSheet.Cells[i, j].Text != string.Empty);
            

            // location and file name of where the program will save the decypted information
            excelApp.ActiveWorkbook.SaveAs(@"C:\Users\sxiong\desktop\decrypttest.xls", Excel.XlFileFormat.xlWorkbookNormal);

            newXLWorkbook.Close();
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(newXLWorksheet);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(newXLWorkbook);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            textBox2.Text = "Decryption Done";

            textBox3.Text = sw.ElapsedMilliseconds.ToString();

            textBox4.Text = "Done";
        }


        // File decryption using Parallel programming
        private async Task button1_ClickParallel()
        {
            // decrypt_file button

            var sw = new Stopwatch();
            sw.Start();

            textBox4.Text = "Processing. . . . .";


            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            // Excel.Range range;

            xlWorkBook = xlApp.Workbooks.Open(filePath);
            xlWorkSheet = xlWorkBook.Worksheets[1];

            //in what row and columnm to begin decrypting on
            int i = 2; //row
            int j = 2; //column

            string decryptedData = "";


            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook newXLWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet newXLWorksheet = (Excel.Worksheet)newXLWorkbook.Sheets.Add();

            //in what row and column to where to place the decypted file on the new excel file
            int row = 1;
            int column = 1;

            List<Task<string>> task = new List<Task<string>>();

            do
            {
                if (xlWorkSheet.Cells[i, j].Text != "NULL")
                {
                    decryptedData = xlWorkSheet.Cells[i, j].Value2;
                    //decryptedData = await Task.Run(() => DecryptString(decryptedData));
                    task.Add(Task.Run(() => DecryptString(decryptedData)));
                    //newXLWorksheet.Cells[row, column] = decryptedData;
                    //row++;
                }
                else
                {
                    decryptedData = "NULL";
                    task.Add(Task.Run(() => DecryptString(decryptedData)));
                    //newXLWorksheet.Cells[row, column] = decryptedData;
                    //row++;
                }

                //newXLWorksheet.Cells[row, column] = decryptedData;
                //row++;

                i++;

            } while (xlWorkSheet.Cells[i, j].Text != string.Empty);

            var results = await Task.WhenAll(task);
            
            foreach(var item in results)
            {
                newXLWorksheet.Cells[row, column] = item;
                row++;
            }
            


            // location and file name of where the program will save the decypted information
            excelApp.ActiveWorkbook.SaveAs(@"C:\Users\sxiong\desktop\decrypttest.xls", Excel.XlFileFormat.xlWorkbookNormal);

            newXLWorkbook.Close();
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(newXLWorksheet);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(newXLWorkbook);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            textBox2.Text = "Decryption Done";

            textBox3.Text = sw.ElapsedMilliseconds.ToString();

            textBox4.Text = "Done";
        }


        private async void button1_Click(object sender, EventArgs e)
        {
            await button1_ClickAsync();
        }

        // File decryption using synchronous programming
        private async void button1_ClickSynchronous(object sender, EventArgs e)
        {   // decrypt_file button

            //await button1_ClickAsync();

            textBox4.Text = "Processing. . . . .";

            var sw = new Stopwatch();
            sw.Start();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
           // Excel.Range range;

            xlWorkBook = xlApp.Workbooks.Open(filePath);
            xlWorkSheet = xlWorkBook.Worksheets[1];

            //in what row and columnm to begin decrypting on
            int i = 2; //row
            int j = 2; //column

            string decryptedData = "" ;


            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook newXLWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet newXLWorksheet = (Excel.Worksheet)newXLWorkbook.Sheets.Add();

            //in what row and column to where to place the decypted file on the new excel file
            int row = 1;
            int column = 1;

           
            do
            {
                if(xlWorkSheet.Cells[i,j].Text != "NULL")
                {
                    decryptedData = xlWorkSheet.Cells[i, j].Value2;
                    decryptedData = DecryptString(decryptedData);
                    newXLWorksheet.Cells[row, column] = decryptedData;
                    row++;
                }
                else
                {
                    decryptedData = "NULL";
                    newXLWorksheet.Cells[row, column] = decryptedData;
                    row++;
                }
                
                
                i++;

            } while (xlWorkSheet.Cells[i, j].Text != string.Empty);

            // location and file name of where the program will save the decypted information
            excelApp.ActiveWorkbook.SaveAs(@"C:\Users\sxiong\desktop\decrypttest.xls", Excel.XlFileFormat.xlWorkbookNormal);

            newXLWorkbook.Close();
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(newXLWorksheet);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(newXLWorkbook);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            textBox2.Text = "Decryption Done";

            textBox3.Text = sw.ElapsedMilliseconds.ToString();

            textBox4.Text = "DONE";

        }

        private async void button2_Click(object sender, EventArgs e)
        {
            //textBox2.Text = DecryptString(textBox1.Text);

            await button1_ClickAsync();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Text = EncryptData(textBox1.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
          
            filePath = load_ExcelFile();
        }

       

        private string load_ExcelFile()
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.InitialDirectory = "C:\\";
            openFile.Filter = "Excel (*.xlsx, *.xls)| *.xls;*.xlsx";
            openFile.FilterIndex = 0;
            openFile.RestoreDirectory = true;

            textBox1.Text = "File Loaded";
            if(openFile.ShowDialog() == DialogResult.OK)
            {
                string path = openFile.FileName;
                //textBox1.Text = path;
                return path;
            }

            return null;
        }

        public static string EncryptData(string message)
        {
            string passKey = string.Empty;

            //passKey = "anotherKey";
            passKey = GetKey();
        
            RijndaelManaged aes256 = new RijndaelManaged();
            //4qFdu0feymPJkO6aJnFK1IEFgK/BF2EJAq/o9qxQT3Q=

            aes256.KeySize = 256;
            aes256.BlockSize = 256;
            aes256.Padding = PaddingMode.PKCS7;
            aes256.Mode = CipherMode.ECB;
            aes256.Key = Encoding.ASCII.GetBytes(passKey);
            aes256.GenerateIV();

            ICryptoTransform encryptor = aes256.CreateEncryptor();
            MemoryStream ms = new MemoryStream();
            CryptoStream cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write);
            StreamWriter mSWriter = new StreamWriter(cs);
            mSWriter.Write(message);
            mSWriter.Flush();
            cs.FlushFinalBlock();
            byte[] cypherTextBytes = ms.ToArray();

            ms.Close();
            return Convert.ToBase64String(cypherTextBytes);
        }

        public static string DecryptString(string text)
        {
            string passKey = string.Empty;

            if(text == "NULL")
            {
                return "NULL";
            }

            passKey = GetKey();

            RijndaelManaged aes256 = new RijndaelManaged();

            aes256.KeySize = 256;
            aes256.BlockSize = 256;
            aes256.Padding = PaddingMode.PKCS7;
            aes256.Mode = CipherMode.ECB;
            aes256.Key = Encoding.ASCII.GetBytes(passKey);
            aes256.GenerateIV();

            byte[] encryptedData = Convert.FromBase64String(text);
            ICryptoTransform transform = aes256.CreateDecryptor();
            byte[] plainText = transform.TransformFinalBlock(encryptedData, 0, encryptedData.Length);

            return Encoding.UTF8.GetString(plainText);
        }

      

        public static string GetKey()
        {
            string registryValue = string.Empty;

            // input encryption key here, but be 16-byte key ex."abcdefghijklmnop"
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(@""); // <-- location of key

            try
            {
                if (registryKey != null)
                {
                   
                    object keyValue = registryKey.GetValue("sN");
                    registryValue = keyValue.ToString();
                }
            }
            catch (Exception ex)
            {
                //string[] arrayParams = new string[] { ex.ToString(), "F002", "getRegistryKey", "TryCatch", DateTime.Now.ToString() };
                //string strsql = "insert into data_ErrorLog (Error,FormNumber,FunctionName,ADDBY,ADDDTTM) values (@Param1,@Param2,@Param3,@Param4,@Param5)";
                //_dbConnection.ExecuteCommandNonQuery(strsql, arrayParams);
            }
            return registryValue;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

    }
}
