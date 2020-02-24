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

namespace decrypter
{
    public partial class Form1 : Form
    {

        private string anotherKey = "abcdefghijklmnop";
        private string filePath = ""; // empty string


        public Form1()
        {
            InitializeComponent();


            button1.Click += new EventHandler(button1_Click);

            button2.Click += new EventHandler(button2_Click);

            button3.Click += new EventHandler(button3_Click);
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
           // Excel.Range range;

            xlWorkBook = xlApp.Workbooks.Open(filePath);
            xlWorkSheet = xlWorkBook.Worksheets[1];

            int i = 2; //row
            int j = 5; //column

            string decryptedData = "" ;


            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook newXLWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet newXLWorksheet = (Excel.Worksheet)newXLWorkbook.Sheets.Add();

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

        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Text = DecryptString(textBox1.Text);
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
            passKey = GetKey();
            RijndaelManaged aes256 = new RijndaelManaged();

            aes256.KeySize = 256;
            aes256.BlockSize = 256;
            aes256.Padding = PaddingMode.None;
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
            passKey = GetKey();
            RijndaelManaged aes256 = new RijndaelManaged();

            aes256.KeySize = 256;
            aes256.BlockSize = 256;
            aes256.Padding = PaddingMode.None;
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
                    registryValue = registryKey.GetValue("sN").ToString();
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
    }
}
