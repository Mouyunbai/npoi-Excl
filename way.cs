using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data;

namespace opexcl
{
    class way
    {
        FileStream filer;
        HSSFWorkbook wbr = new HSSFWorkbook();//excl文件定义
        int k = 0;//k=0时文件已存在  k=1时文件不存在要新建
        Sheet sheetr;//表sheetr数据表，sheetr1铭牌表sheetr2模块表
        DataTable dt,dt2, dt1;
        int xx = 0;//表行位
        /// <summary>
        /// 将excl中的数据写入到datatable中
        /// </summary>
        /// <returns>datatable数据</returns>
        public DataTable getwritetodt()
        {
            dt= ExcelToDataTable(sheetr, false);
            return dt;
        }
        /// <summary>
        /// 获取表行
        /// </summary>
        /// <returns>返回行数</returns>
        public string geth()
        {
            var firstRow = sheetr.GetRow(0);
            if(firstRow == null)
                        return 0+"";
            int cellCount = firstRow.LastCellNum;
            return cellCount + "";
        }
        /// <summary>
        /// 获取列数
        /// </summary>
        /// <returns>返回并数</returns>
        public string getl()
        {
            int rowCount = sheetr.LastRowNum;
            return rowCount+ 1 + "";
        }
        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(Sheet sheetName, bool isFirstRowColumn)
        {
            Sheet sheet = sheetName;
            var data = new DataTable();
            data.TableName = "sheet1";
            int startRow = 0;
            try
            {
                if (sheet != null)
                {
                    var firstRow = sheet.GetRow(0);
                    if (firstRow == null)
                        return data;
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数
                    startRow = isFirstRowColumn ? sheet.FirstRowNum + 1 : sheet.FirstRowNum;
                    
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        //.StringCellValue;
                        var column = new DataColumn(Convert.ToChar(((int)'A') + i).ToString());
                        /*数据库列名
                        if (isFirstRowColumn)
                        {
                            //var columnName = firstRow.GetCell(i).StringCellValue;
                            //column = new DataColumn(columnName);
                            var columnName = sheet.GetRow(0).GetCell(i).ToString();
                            column = new DataColumn(columnName);
                        }
                        */
                        data.Columns.Add(column);
                    }
                    data.Columns.Add(new DataColumn("行数"));
                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        Row row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　

                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                dataRow[j] = row.GetCell(j, MissingCellPolicy.RETURN_NULL_AND_BLANK).ToString();
                        }
                        dataRow[cellCount] = i;
                        data.Rows.Add(dataRow);
                    }
                }
                else throw new Exception("Don not have This Sheet");

                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 打开excl将excl中的文件与表信息读取
        /// </summary>
        /// <param name="openaddr">地址</param>
    public void openingexcl(string openaddr)
        {
            filer = new FileStream(openaddr, FileMode.Open, FileAccess.ReadWrite);
            wbr = new HSSFWorkbook(filer);
            sheetr = wbr.GetSheet("Sheet1");
        }
        public void closing()
        {
            filer.Close();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns>表中现在有的数据</returns>
        public string getnumber()
        {
            return xx+"";
        }


        /// <summary>
        /// 用于查看文件是否已存在
        /// </summary>
        /// <param name="openaddr">文件的地址</param>
        /// <returns>0说明不存在，1说明存在</returns>
        public int getexcl(string openaddr)
        {
            try
            {
                filer = new FileStream(openaddr, FileMode.Open, FileAccess.ReadWrite);
            }
            catch (Exception e)
            {
                return 0;
            }
            filer.Close();
            return 1;
        }
        /// <summary>
        /// 创建excl
        /// </summary>
        /// <param name="openaddr">创建的位置</param>
        /// <returns>1为创建成功，0创建不功能</returns>
        public int createexcl(string openaddr)
        {
            try
            {
                sheetr = wbr.CreateSheet("Sheet1");

                filer = new FileStream(openaddr, FileMode.Create);
                wbr.Write(filer);
                filer.Close();
            }
            catch (Exception e)
            {
                return 0;
            }
            return 1;
        }
        /// <summary>
        /// 向excl中写入标题
        /// </summary>
        /// <param name="li">需要写入的标题列</param>
        /// <returns>1为写入成功；0为写入失败</returns>
        public int WriteTitle(List<string>li)
        {
            int cou = 0,i=0;
            //cou = li.Count;
            //creecxlroll(0,cou);
            foreach (string s in li)
            {
                //sheetr.GetRow(0).GetCell(i++).SetCellValue(s);
                sheetr.CreateRow(0).CreateCell(i++).SetCellValue(s);
            }
            return 1;
        }
        /// <summary>
        /// 向excl中写入数据
        /// </summary>
        /// <param name="li">需要写入的数据</param>
        /// <returns>1为写入成功；0为写入失败</returns>
        public int WriteData(List<string> li)
        {
            int cou = 0, i = 0,k=0;
            //cou = li.Count;
            k=Convert.ToInt32(getl());
            //creecxlroll(k, cou);
            foreach (string s in li)
            {
                //sheetr.GetRow(k).GetCell(i++).SetCellValue(s);
                sheetr.CreateRow(k).CreateCell(i++).SetCellValue(s);
            }
            return 1;
        }
        /// <summary>
        /// 保存并关闭excl
        /// </summary>
        /// <param name="saveaddr">保存的地址</param>
        public void saveexcl(string saveaddr)
        {
            filer = new FileStream(saveaddr, FileMode.Create);
            wbr.Write(filer);
            filer.Close();

        }
        /// <summary>
        /// 删除指定的行
        /// </summary>
        /// <param name="h">行号</param>
        /// <returns>返回1删除成功</returns>
        public int RemoveH(int h)
        {
            var row = sheetr.GetRow(h);
            sheetr.RemoveRow(row);
            return 1;
        }
        /// <summary>
        /// 查找指定数据所在的行
        /// </summary>
        /// <param name="code">要查找的数据</param>
        /// <param name="title">列标题</param>
        /// <returns>返回所在的行</returns>
        public int SelectH(string code,string title)
        {
            DataRow[] rows =dt.Select(title+"='" + code + "'");

            return Convert.ToInt32(rows[0][5]);
        }
        /// <summary>
        /// 预备创建行列
        /// </summary>
        /// <param name="h">行位</param>
        /// <param name="l">列数</param>
        void creecxlroll(int h,int l)
        {
            sheetr.CreateRow(h);
            for (int i = 0; i < l; i++)
                sheetr.GetRow(h).CreateCell(i);
        }

        /// <summary>
        /// 把DataTable的数据写入到指定的excel文件中
        /// </summary>
        /// <param name="TargetFileNamePath">目标文件excel的路径</param>
        /// <param name="sourceData">要写入的数据</param>
        /// <param name="sheetName">excel表中的sheet的名称，可以根据情况自己起</param>
        /// <param name="IsWriteColumnName">是否写入DataTable的列名称</param>
        /// <returns>返回写入的行数</returns>
        public int DataTableToExcel(string TargetFileNamePath, DataTable sourceData, string sheetName, bool IsWriteColumnName)
        {

            //数据验证
            if (!File.Exists(TargetFileNamePath))
            {
                //excel文件的路径不存在
                throw new ArgumentException("excel文件的路径不存在或者excel文件没有创建好");
            }
            if (sourceData == null)
            {
                throw new ArgumentException("要写入的DataTable不能为空");
            }

            if (sheetName == null && sheetName.Length == 0)
            {
                throw new ArgumentException("excel中的sheet名称不能为空或者不能为空字符串");
            }



            //根据Excel文件的后缀名创建对应的workbook
            HSSFWorkbook workbook = null;
            if (TargetFileNamePath.IndexOf(".xlsx") > 0)
            {  //2007版本的excel
                workbook = new HSSFWorkbook();
            }
            else if (TargetFileNamePath.IndexOf(".xls") > 0) //2003版本的excel
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                return -1;    //都不匹配或者传入的文件根本就不是excel文件，直接返回
            }



            //excel表的sheet名
            Sheet sheet = workbook.CreateSheet(sheetName);
            if (sheet == null) return -1;   //无法创建sheet，则直接返回


            //写入Excel的行数
            int WriteRowCount = 0;



            //指明需要写入列名，则写入DataTable的列名,第一行写入列名
            if (IsWriteColumnName)
            {
                //sheet表创建新的一行,即第一行
                var ColumnNameRow = sheet.CreateRow(0); //0下标代表第一行
                //进行写入DataTable的列名
                for (int colunmNameIndex = 0; colunmNameIndex < sourceData.Columns.Count; colunmNameIndex++)
                {
                    ColumnNameRow.CreateCell(colunmNameIndex).SetCellValue(sourceData.Columns[colunmNameIndex].ColumnName);
                }
                WriteRowCount++;
            }


            //写入数据
            for (int row = 0; row < sourceData.Rows.Count; row++)
            {
                //sheet表创建新的一行
                var newRow = sheet.CreateRow(WriteRowCount);
                for (int column = 0; column < sourceData.Columns.Count; column++)
                {

                    newRow.CreateCell(column).SetCellValue(sourceData.Rows[row][column].ToString());

                }

                WriteRowCount++;  //写入下一行
            }


            //写入到excel中
            FileStream fs = new FileStream(TargetFileNamePath, FileMode.Open, FileAccess.Write);
            workbook.Write(fs);

            fs.Flush();
            fs.Close();
            return WriteRowCount;
        }

        

    }
}