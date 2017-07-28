using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;


namespace Excel.Component.Library
{
    public class ExcelMap
    {
        private  ExcelPackage _pack;
        #region Properties


        #endregion

        #region Constructors

        public ExcelMap()
        {
        }

        #endregion

        #region Methods


        public byte[] Extract()
        {
            return Process();
        }

        public void Extract(string path, string filename, bool openFile = false)
        {

            byte[] result = Process();
            var longFilename = Path.Combine(path, filename + ".xlsx");
            BinaryWriter bw = new BinaryWriter(File.Open(longFilename, FileMode.CreateNew));
            bw.Write(result);
            bw.Flush();
            bw.Close();

            if (openFile)
            {
                System.Diagnostics.Process.Start(longFilename);
            }
        }

        public ExcelWorksheet GetExcelWorkSheet(string label)
        {
            try
            {
                _pack = new ExcelPackage();
                return _pack.Workbook.Worksheets.Add(label);
            }
            finally { }
        }

        private byte[] Process()
        {
            byte[] result = null;

            result = _pack.GetAsByteArray();

            return result;
        }

        #endregion
    }
}