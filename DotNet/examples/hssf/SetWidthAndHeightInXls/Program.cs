/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;


namespace SetWidthAndHeightInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            NPOI.SS.UserModel.ISheet sheet1 = hssfworkbook.CreateSheet("Sheet1");
            //set the width of columns
            sheet1.SetColumnWidth(0,50 * 256);
            sheet1.SetColumnWidth(1, 100 * 256);
            sheet1.SetColumnWidth(2, 150 * 256);

            //set the width of height
            sheet1.CreateRow(0).Height = 100*20;
            sheet1.CreateRow(1).Height = 200*20;
            sheet1.CreateRow(2).Height = 300*20;

            sheet1.DefaultRowHeightInPoints = 50;

            WriteToFile();
        }


        static NPOI.HSSF.UserModel.HSSFWorkbook hssfworkbook;

        static void WriteToFile()
        {
            //Write the stream data of workbook to the root directory
            System.IO.FileStream file = new System.IO.FileStream(@"test.xls", System.IO.FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }

        static void InitializeWorkbook()
        {
            hssfworkbook = new NPOI.HSSF.UserModel.HSSFWorkbook();

            //create a entry of DocumentSummaryInformation
            NPOI.HPSF.DocumentSummaryInformation dsi = NPOI.HPSF.PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            //create a entry of SummaryInformation
            NPOI.HPSF.SummaryInformation si = NPOI.HPSF.PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;
        }
    }
}
