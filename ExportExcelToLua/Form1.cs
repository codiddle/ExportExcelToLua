using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel;
using System.IO;

namespace ExportExcelToLua
{
    public partial class Form1 : Form
    {
        public const int STARTING_ROW = 3; // 0 based so row 4 == index 3
        public int[] LINES_TO_SKIP = new int[]{
            2, 13, 24, 35, 46, 57, 68, 79
        };
        private IExcelDataReader excelReader { get; set; }
        private FileStream objFileStream { get; set; }
        private string path { get; set; }

        public Form1()
        {
            InitializeComponent();
        }

        public string luaText = "";



        private void button1_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.Title = "Select Excel file.";
            var dialogResult = ofd.ShowDialog();
            if(dialogResult == DialogResult.OK)
            {
                var fileObj = File.OpenRead(ofd.FileName);
                var binaryExcel = ReadToEnd(fileObj);
                var ds = GetDataTableFromBinaryFile(binaryExcel);
                string raceAndJob = "";
                string race = "";
                string job = "";
                string[] raceAndJobArray = null;
                foreach (DataTable table in ds.Tables)
                {
                    foreach(DataRow row in table.Rows)
                    {
                        if (LINES_TO_SKIP.Contains(table.Rows.IndexOf(row) + 1))
                        {
                            continue;
                        }
                        if (table.Rows.IndexOf(row) == 0)
                        {
                            // this tells us race/job
                            raceAndJob = row.ItemArray[0].ToString();
                            raceAndJobArray = raceAndJob.Split(' ');
                            race = raceAndJobArray[0];
                            job = raceAndJobArray[2];
                            if(raceAndJobArray.Length > 3)
                            {
                                job += raceAndJobArray[3];
                            }
                            luaText += "[\"" + job + "\"] = {";
                        }
                        if (table.Rows.IndexOf(row) > 1)
                        {
                            luaText += "[";

                            var level = row.ItemArray[0].ToString();
                            luaText += level;
                            luaText += "] = {";
                            var hp = row.ItemArray[1].ToString();
                            luaText += "[\"HP\"] =" + hp + ",";
                            var mp = row.ItemArray[2].ToString();
                            luaText += "[\"MP\"] =" + mp + ",";
                            var str = row.ItemArray[3].ToString();
                            luaText += "[\"STR\"] =" + str + ",";
                            var dex = row.ItemArray[4].ToString();
                            luaText += "[\"DEX\"] =" + dex + ",";
                            var vit = row.ItemArray[5].ToString();
                            luaText += "[\"VIT\"] =" + vit + ",";
                            var agi = row.ItemArray[6].ToString();
                            luaText += "[\"AGI\"] =" + agi + ",";
                            var intl = row.ItemArray[7].ToString();
                            luaText += "[\"INT\"] =" + intl + ",";
                            var mnd = row.ItemArray[8].ToString();
                            luaText += "[\"MND\"] =" + mnd + ",";
                            var chr = row.ItemArray[9].ToString();
                            luaText += "[\"CHR\"] =" + chr;
                            luaText += "},";
                        }

                        
                    }
                    luaText += "},";
                }
                textBox1.Text = luaText;
                luaText = "";
                fileObj.Close();
            }
            
        }

        private static byte[] ReadToEnd(System.IO.Stream stream)
        {
            long originalPosition = 0;

            if (stream.CanSeek)
            {
                originalPosition = stream.Position;
                stream.Position = 0;
            }

            try
            {
                byte[] readBuffer = new byte[4096];

                int totalBytesRead = 0;
                int bytesRead;

                while ((bytesRead = stream.Read(readBuffer, totalBytesRead, readBuffer.Length - totalBytesRead)) > 0)
                {
                    totalBytesRead += bytesRead;

                    if (totalBytesRead == readBuffer.Length)
                    {
                        int nextByte = stream.ReadByte();
                        if (nextByte != -1)
                        {
                            byte[] temp = new byte[readBuffer.Length * 2];
                            Buffer.BlockCopy(readBuffer, 0, temp, 0, readBuffer.Length);
                            Buffer.SetByte(temp, totalBytesRead, (byte)nextByte);
                            readBuffer = temp;
                            totalBytesRead++;
                        }
                    }
                }

                byte[] buffer = readBuffer;
                if (readBuffer.Length != totalBytesRead)
                {
                    buffer = new byte[totalBytesRead];
                    Buffer.BlockCopy(readBuffer, 0, buffer, 0, totalBytesRead);
                }
                return buffer;
            }
            finally
            {
                if (stream.CanSeek)
                {
                    stream.Position = originalPosition;
                }
            }
        }

        private DataSet GetDataTableFromBinaryFile(byte[] binaryExcel)
        {
            path = Path.GetTempFileName();
            objFileStream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite);
            objFileStream.Write(binaryExcel, 0, binaryExcel.Length);
            excelReader = ExcelReaderFactory.CreateOpenXmlReader(objFileStream);
            return excelReader.AsDataSet();
        }

    }
}
