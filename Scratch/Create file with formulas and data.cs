        static void Main(string[] args)
        {
            InitializeWorkbook();

            Dictionary<String, ICellStyle> styles = CreateStyles(hssfworkbook);

            ISheet sheet = hssfworkbook.CreateSheet("Timesheet");
            IPrintSetup printSetup = sheet.PrintSetup;
            printSetup.Landscape = true;
            sheet.FitToPage=(true);
            sheet.HorizontallyCenter=(true);

            //title row
            IRow titleRow = sheet.CreateRow(0);
            titleRow.HeightInPoints=(45);
            ICell titleCell = titleRow.CreateCell(0);
            titleCell.SetCellValue("Weekly Timesheet");
            titleCell.CellStyle= (styles["title"]);
            sheet.AddMergedRegion(CellRangeAddress.ValueOf("$A$1:$L$1"));

            //header row
            IRow headerRow = sheet.CreateRow(1);
            headerRow.HeightInPoints = (40);
            ICell headerCell;
            for (int i = 0; i < titles.Length; i++)
            {
                headerCell = headerRow.CreateCell(i);
                headerCell.SetCellValue(titles[i]);
                headerCell.CellStyle = (styles["header"]);
            }


            int rownum = 2;
            for (int i = 0; i < 10; i++)
            {
                IRow row = sheet.CreateRow(rownum++);
                for (int j = 0; j < titles.Length; j++)
                {
                    ICell cell = row.CreateCell(j);
                    if (j == 9)
                    {
                        //the 10th cell contains sum over week days, e.g. SUM(C3:I3)
                        String reference = "C" + rownum + ":I" + rownum;
                        cell.CellFormula = ("SUM(" + reference + ")");
                        cell.CellStyle = (styles["formula"]);
                    }
                    else if (j == 11)
                    {
                        cell.CellFormula = ("J" + rownum + "-K" + rownum);
                        cell.CellStyle = (styles["formula"]);
                    }
                    else
                    {
                        cell.CellStyle = (styles["cell"]);
                    }
                }
            }

            //row with totals below
            IRow sumRow = sheet.CreateRow(rownum++);
            sumRow.HeightInPoints = (35);
            ICell cell1 = sumRow.CreateCell(0);
            cell1.CellStyle = (styles["formula"]);

            ICell cell2 = sumRow.CreateCell(1);
            cell2.SetCellValue("Total Hrs:");
            cell2.CellStyle=(styles["formula"]);

            for (int j = 2; j < 12; j++) {
                ICell cell = sumRow.CreateCell(j);
                String reference = (char)('A' + j) + "3:" + (char)('A' + j) + "12";
                cell.CellFormula = ("SUM(" + reference + ")");
                if(j >= 9)
                    cell.CellStyle = (styles["formula_2"]);
                else
                    cell.CellStyle = (styles["formula"]);
            }

            rownum++;
            sumRow = sheet.CreateRow(rownum++);
            sumRow.HeightInPoints = 25;
            ICell cell3 = sumRow.CreateCell(0);
            cell3.SetCellValue("Total Regular Hours");
            cell3.CellStyle = styles["formula"];
            cell3 = sumRow.CreateCell(1);
            cell3.CellFormula = ("L13");
            cell3.CellStyle=styles["formula_2"];
            sumRow = sheet.CreateRow(rownum++);
            sumRow.HeightInPoints = (25);
            cell3 = sumRow.CreateCell(0);
            cell3.SetCellValue("Total Overtime Hours");
            cell3.CellStyle = styles["formula"];
            cell3 = sumRow.CreateCell(1);
            cell3.CellFormula = ("K13");
            cell3.CellStyle = styles["formula_2"];

                    //set sample data
            for (int i = 0; i < sample_data.GetLength(0); i++)
            {
                IRow row = sheet.GetRow(2 + i);
                for (int j = 0; j < sample_data.GetLength(1); j++)
                {
                    if (sample_data[i,j] == null)
                        continue;

                    if (sample_data[i,j] is String)
                    {
                        row.GetCell(j).SetCellValue((String)sample_data[i,j]);
                    }
                    else
                    {
                        row.GetCell(j).SetCellValue((Double)sample_data[i,j]);
                    }
                }
            }

                    //finally set column widths, the width is measured in units of 1/256th of a character width
        sheet.SetColumnWidth(0, 30*256); //30 characters wide
        for (int i = 2; i < 9; i++) {
            sheet.SetColumnWidth(i, 6*256);  //6 characters wide
        }
        sheet.SetColumnWidth(10, 10*256); //10 characters wide


            WriteToFile();
        }
