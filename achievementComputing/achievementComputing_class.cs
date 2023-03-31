using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace achievementComputing
{
    class achievementComputing_class
    {
        internal string excelFilePath;
        internal int sheetColCount;
        internal int sheetRowCount;
        internal ExcelPackage EPexcel;
        //internal string sheetName;
        internal ExcelWorksheet sheet;
        internal string sheetName;

        table_class courseTable;
        table_class courseWeightSumTable;
        table_class courseAchievementTable;
        table_class indicatorTable;
        table_class indicatorWeightSumTable;
        table_class indicatorAchievementTalbe;
        table_class indicatorWeightTable;
        table_class indicatorRatingTable;
        table_class meanScoreTable;
        table_class scoreTable;
        
        //internal achievementComputing_class() /////测试用
        //{
        //    //////////////////////////////////////测试设置
        //    excelFilePath = "达成度.xlsx";
        //    sheetColCount = 16;
        //    sheetRowCount = 102;
        //    string sheetName = "达成度";
        //    //sheetName = "Sheet1";
        //    ////////////////////////////////////////
        //    FileInfo excelFileInfo = new FileInfo(excelFilePath);
        //    EPexcel = new ExcelPackage(excelFileInfo);
        //    sheet = EPexcel.Workbook.Worksheets[sheetName];
        //    computeAchievement();
        //}
        internal achievementComputing_class(string FilePath)
        {
            excelFilePath = FilePath;
            EPexcel = new ExcelPackage(new FileInfo(excelFilePath));
        }

        internal List<string> getSheetsNames()
        {
            List<string> sheetNameList = new List<string>();
            for (int i = 1; i <= EPexcel.Workbook.Worksheets.Count; i++)
            {
                sheetNameList.Add(EPexcel.Workbook.Worksheets[i].Name);
            }
            return (sheetNameList);
        }

        internal DataTable getDataTableFromSheet(string sheetname)
        {
			sheetName = sheetname;
			sheet = EPexcel.Workbook.Worksheets[sheetName];
            if (sheet.Dimension == null)
                return (null);
            sheetColCount = sheet.Dimension.End.Column;
            sheetRowCount = sheet.Dimension.End.Row;
            DataTable dt = new DataTable();
            for (int j = 1; j <= sheetColCount; j++)
            {
                dt.Columns.Add();
            }
            //Get Row Data of Excel
            for (int i = 1; i <= sheetRowCount; i++) //Loop for available row of excel data
            {
                DataRow row = dt.NewRow(); //assign new row to DataTable
                for (int j = 1; j <= sheetColCount; j++) //Loop for available column of excel data
                {
                    row[j - 1] = sheet.Cells[i, j].Value;
                }
                dt.Rows.Add(row); //add row to DataTable
            }
            return (dt);
        }

        internal void computeAchievement()
        {
            try
            {
                if (sheet.Dimension == null)
                {
                    showErrorInfo("空表！");
                }
                //设置区域
                getTables();
                //检测错误
                checkTables();
                //设置计算公式
                setFormula();
                //设置格式
                borderTables();
                colorTables();
            }
            catch(abortException ae)
			{
                EPexcel = new ExcelPackage(new FileInfo(excelFilePath));
                sheet= EPexcel.Workbook.Worksheets[sheetName];
                sheetColCount = sheet.Dimension.End.Column;
                sheetRowCount = sheet.Dimension.End.Row;
                return;
			}
        ////保存
        rewrite:
            try
            {
                EPexcel.Save();
                MessageBox.Show("保存成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                DialogResult r = MessageBox.Show("文件被其它程序打开，无法保存。",
                    "错误", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                if (r.ToString() == "Retry")
                    goto rewrite;
            }
        }

        static DialogResult showErrorInfo(string info, MessageBoxButtons buttonType = MessageBoxButtons.OK)
        {
            DialogResult r = MessageBox.Show(info, "错误",
                buttonType, MessageBoxIcon.Error);
            throw new abortException();
            //if (isExit)
            //    System.Environment.Exit(0);
            return (r);
        }

        table_class findTable(string startString, string endString) //用于获取课程目标表和指标点表
        {
            table_class table = new table_class();
            ExcelRange rg = findRangeFromCol(sheet.Cells[1, 1, sheetRowCount+1, sheetColCount], startString);
            if (rg == null)
            {
                showErrorInfo("没有找到\"" + startString + "\"");
            }
            table.startCol = 2;
            table.startRow = rg.Start.Row + 1;
            rg = findRangeFromRow(sheet.Cells[rg.Start.Row, 1, rg.Start.Row, sheetColCount+1], "");
            table.endCol = rg.End.Column - 1;
            rg = findRangeFromCol(sheet.Cells[table.startRow, 1, sheetRowCount+1, 1], endString);
            table.endRow = rg.End.Row - 1;
            table.computeWidthAndHeight();
            table.recaptureRange(sheet);
            return (table);
        }
        table_class findNextTable(table_class lastTable)
        {
            table_class table = new table_class();
            table.startRow = lastTable.endRow + 1;
            table.endRow = lastTable.endRow + 1;
            table.startCol = lastTable.startCol;
            table.endCol = lastTable.endCol;
            table.computeWidthAndHeight();
            table.recaptureRange(sheet);
            return (table);
        }
        table_class findMeanScoreTable()
        {
            table_class msTable = new table_class();
            msTable.startCol = 2;
            ExcelRange rg = findRangeFromCol(sheet.Cells[1, 1, sheetRowCount+1, sheetColCount], "平均成绩");
            if (rg == null)
            {
                showErrorInfo("没有找到\"平均成绩\"");
                //MessageBox.Show(, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //System.Environment.Exit(0);
            }
            msTable.startRow = rg.Start.Row;
            msTable.endRow = rg.Start.Row;
            rg = findRangeFromRow(sheet.Cells[msTable.startRow + 1, 1, msTable.startRow + 1, sheetColCount+1], "");
            msTable.endCol = rg.End.Column - 1;
            msTable.computeWidthAndHeight();
            msTable.recaptureRange(sheet);
            return (msTable);
        }
        table_class findScoreTable()
        {
            table_class sTable = new table_class();
            sTable.startRow = meanScoreTable.endRow + 2;
            sTable.startCol = 1;
            ExcelRange rg = findRangeFromRow
                (sheet.Cells[sTable.startRow - 1, 1, sTable.startRow - 1, sheetColCount+1], "");
            sTable.endCol = rg.End.Column - 1;
            rg = findRangeFromCol(sheet.Cells[sTable.startRow - 1, 1, sheetRowCount+1, 1], "");
            sTable.endRow = rg.End.Row - 1;
            sTable.computeWidthAndHeight();
            sTable.recaptureRange(sheet);
            return (sTable);
        }
        void getTables()
        {
            courseTable = findTable("课程目标", "课程目标总权重");
            courseWeightSumTable = findNextTable(courseTable);
            courseAchievementTable = findNextTable(courseWeightSumTable);
            indicatorTable = findTable("指标点", "指标点总权重");
            indicatorWeightSumTable = findNextTable(indicatorTable);
            indicatorAchievementTalbe = findNextTable(indicatorWeightSumTable);
            indicatorWeightTable = findNextTable(indicatorAchievementTalbe);
            indicatorRatingTable = findNextTable(indicatorWeightTable);
            meanScoreTable = findMeanScoreTable();
            scoreTable = findScoreTable();
        }
        internal static ExcelRange findRangeFromCol(ExcelRange sourceRange, string s)
        {
            ExcelRange range = null;
            int colIndex = sourceRange.Start.Column;
            int startRowID = sourceRange.Start.Row;
            int endRowID = sourceRange.End.Row;
            for (int i = startRowID; i <= endRowID; i++)
            {
                if (sourceRange[i, colIndex].Text == s)
                    return (sourceRange[i, colIndex]);
            }
            return (range);
        }
        internal static ExcelRange findRangeFromRow(ExcelRange sourceRange, string s)
        {
            ExcelRange range = null;
            int rowIndex = sourceRange.Start.Row;
            int startColID = sourceRange.Start.Column;
            int endColID = sourceRange.End.Column;
            for (int i = startColID; i <= endColID; i++)
            {
                if (sourceRange[rowIndex, i].Text == s)
                    return (sourceRange[rowIndex, i]);
            }
            return (range);
        }

        void colorTables()
        {
			//Color blue = Color.FromArgb(240, 248, 255);
			Color blue = Color.FromArgb(221, 235, 247);
			Color green = Color.FromArgb(226, 250, 218);
            Color yellow = Color.FromArgb(255, 255, 220);
            courseTable.colour(blue);
            indicatorTable.colour(blue);
            courseWeightSumTable.colour(green);
            indicatorWeightSumTable.colour(green);
            courseAchievementTable.colour(yellow);
            indicatorAchievementTalbe.colour(yellow);
            indicatorWeightTable.colour(blue);
            meanScoreTable.colour(yellow);
            indicatorRatingTable.colour(yellow);

        }
        void borderTables()
        {
            table_class tb1 = new table_class();
            tb1.startCol = 1;
            tb1.startRow = courseTable.startRow - 1;
            tb1.endCol = courseTable.endCol;
            tb1.endRow = courseTable.endRow + 2;
            tb1.recaptureRange(sheet);
            tb1.Border();
            table_class tb2 = new table_class();
            tb2.startCol = 1;
            tb2.startRow = indicatorTable.startRow - 1;
            tb2.endCol = indicatorTable.endCol;
            tb2.endRow = indicatorTable.endRow + 4;
            tb2.recaptureRange(sheet);
            tb2.Border();
            meanScoreTable.startCol = 1;
            meanScoreTable.recaptureRange(sheet);
            meanScoreTable.Border();
            meanScoreTable.startCol = 2;
            meanScoreTable.recaptureRange(sheet);
        }

        void checkTables()
        {
            checkWeight(courseWeightSumTable, courseTable);
            checkWeight(indicatorWeightSumTable, indicatorTable);
            checkCourseObjects();
        }
        void checkWeight(table_class table, table_class sourceTable)
        {
            int r = table.startRow;
            int STheight = sourceTable.height;
            for (int c = table.startCol; c <= table.endCol; c++)
            {
                string f = string.Format("SUM(R[-{0}]C:R[-1]C)", sourceTable.height);
                sheet.Cells[r, c].FormulaR1C1 = f;
                sheet.Cells[r, c].Calculate();
                //if (sheet.Cells[r, c].Text != "1")
                if (Math.Round((double)sheet.Cells[r, c].Value, 8) != 1.0)
                {
                    showErrorInfo("总权重单元格 " + sheet.Cells[r, c].Address + " !=1");
                }
            }
        }
        void checkCourseObjects()
        {
            for (int i = 0; i < courseTable.width; i++)
            {
                string objstr1 = sheet.Cells[courseTable.startRow - 1, courseTable.startCol + i].Text;
                string objstr2 = sheet.Cells[indicatorTable.startRow + i, indicatorTable.startCol - 1].Text;
                if (objstr1 != objstr2)
                {
                    showErrorInfo(objstr1 + " " + objstr2 + " 在两个表中不一致。");
                }
            }
        }
        int findColFromMeanScoreTable(string Item)
        {
            for (int i = 0; i < meanScoreTable.width; i++)
            {
                if (sheet.Cells[meanScoreTable.startRow + 1, meanScoreTable.startCol + i].Text == Item)
                {
                    return (meanScoreTable.startCol + i);
                }
            }
            return (0);
        }

        void setFormula()
        {
            setMeanScoreFormula();
            setIndicatorRatingFormula();
            setIndicatorAchievementFormula();
            setcourseAchievementFormula();
        }
        void setMeanScoreFormula()
        {
            for (int i = 0; i < courseTable.height; i++)
            {
                string scoreItem = sheet.Cells[courseTable.startRow + i, courseTable.startCol - 1].Text;
                int scoreItemCol = findColOfScoreItemFromScoreTable(scoreItem);
                string f = string.Format("AVERAGE(R[2]C:R[{0}]C)", scoreTable.height + 1);
				sheet.Cells[meanScoreTable.startRow, scoreItemCol].FormulaR1C1 = f;
                sheet.Cells[meanScoreTable.startRow - 2, scoreItemCol].FormulaR1C1 = "R[2]C/R[1]C*100";
                if (sheet.Cells[meanScoreTable.startRow - 1, scoreItemCol].Text == "")
                {
                    showErrorInfo(sheet.Cells[meanScoreTable.startRow - 1, scoreItemCol].Address
                        + " 没有设置考核项的满分值，会出现除零错误。");
				}
			}
		}
        int findColOfScoreItemFromScoreTable(string scoreItem)
		{
            int scoreItemCol = findColFromMeanScoreTable(scoreItem);
            if (scoreItemCol == 0)
            {
                showErrorInfo(scoreItem + " 不能在成绩表中发现该项！");
            }
            return (scoreItemCol);
        }
        void setIndicatorRatingFormula()
        {
            for (int i = 0; i < indicatorRatingTable.width; i++)
            {
                int row = indicatorRatingTable.startRow;
                int col = indicatorRatingTable.startCol + i;
                sheet.Cells[row, col].FormulaR1C1 = "R[-2]C*R[-1]C";
            }
        }
        void setIndicatorAchievementFormula()
        {
            int row = indicatorAchievementTalbe.startRow;
            for (int i = 0; i < indicatorAchievementTalbe.width; i++)
            {
                string f = buildIndicatorAchievementFormula(row, indicatorAchievementTalbe.startCol + i);
                sheet.Cells[row, indicatorAchievementTalbe.startCol + i].FormulaR1C1 = f;
            }
        }
        string buildIndicatorAchievementFormula(int row, int col)
        {
            string f = "";
            for (int i = 0; i < courseAchievementTable.width; i++)
            {
                string f1 = string.Format("+R{0}C{1}*R[-{2}]C",
                    courseAchievementTable.startRow,
                    courseAchievementTable.startCol + i,
                    indicatorTable.height + 1 - i);
                f = f + f1;
            }
            f = f.Substring(1);
            return (f);
        }
        void setcourseAchievementFormula()
        {
            //string[] scoreItems = getScoreItems();
            int row = courseAchievementTable.startRow;
            for (int i = 0; i < courseAchievementTable.width; i++)
            {
                string f = buildCourseAchievementFormula(row, courseAchievementTable.startCol + i);
                sheet.Cells[row, courseAchievementTable.startCol + i].FormulaR1C1 = f;
            }
        }
        string buildCourseAchievementFormula(int row, int col)
        {
            string f = "";
            for (int i = 0; i < courseTable.height; i++)
            {
                string scoreItem = sheet.Cells[courseTable.startRow + i, courseTable.startCol - 1].Text;
                int scoreItemCol = findColOfScoreItemFromScoreTable(scoreItem);
                string f1 = string.Format("+R[-{0}]C*R{1}C{2}",
                    courseTable.height + 1 - i, meanScoreTable.startRow-2, scoreItemCol);
                f = f + f1;
            }
            f = f.Substring(1);
            return (f);
        }
    }
    class abortException:Exception
	{

	}
}
