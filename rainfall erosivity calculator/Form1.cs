using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Globalization;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace myCalculator
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// 使用NPOI读取Excel数据
        /// </summary>
        public class ImportCore
        {
            public IWorkbook workbook = null;  //新建IWorkbook对象  
            public string _filePath;
            public FileStream fileStream;


            #region 方法
            /// <summary>
            /// 进行计算
            /// </summary>
            public void Calculate()
            {
            string fileName = _filePath;
            fileStream = new FileStream(@fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本  
            {
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook 
                ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表
                IRow row;     //新建当前工作表行数据  
                IRow rowPre; //定义前一行
                IRow titleRow = sheet.GetRow(0);  //第一行
                ICell timeTitle = titleRow.CreateCell(titleRow.LastCellNum);
                timeTitle.SetCellValue("时刻数字计数法");
                ICell minutes = titleRow.CreateCell(titleRow.LastCellNum);
                minutes.SetCellValue("分钟");
                ICell difTitle = titleRow.CreateCell(titleRow.LastCellNum);
                difTitle.SetCellValue("时间间隔");
                ICell lineTitle = titleRow.CreateCell(titleRow.LastCellNum);
                lineTitle.SetCellValue("水位");
                ICell precipitationTitle = titleRow.CreateCell(titleRow.LastCellNum);
                precipitationTitle.SetCellValue("雨量");
                ICell intensityTitle = titleRow.CreateCell(titleRow.LastCellNum);
                intensityTitle.SetCellValue("雨强");
                ICell udTitle = titleRow.CreateCell(titleRow.LastCellNum);
                udTitle.SetCellValue("单位雨强");
                ICell tdTitle = titleRow.CreateCell(titleRow.LastCellNum);
                tdTitle.SetCellValue("时段雨强");
                ICell etTitle = titleRow.CreateCell(titleRow.LastCellNum);
                etTitle.SetCellValue("E总");
                ICell thirtyTitle = titleRow.CreateCell(titleRow.LastCellNum);
                thirtyTitle.SetCellValue("I30 cm/h");
                ICell RTitle = titleRow.CreateCell(titleRow.LastCellNum);
                RTitle.SetCellValue("R");
                int cloNum = titleRow.LastCellNum;
                int rowNum = sheet.LastRowNum;
                List<double> eList = new List<double>();  //储存时段雨强
                List<double> difList = new List<double>();  //储存简单时间间隔
                List<double> AlldifList = new List<double>();  //储存所有时间间隔
                List<double> AllrainList = new List<double>();  //储存所有雨量来判断
                List<double> AlleList = new List<double>();  //储存所有雨强
                List<double> countList = new List<double>();  //临时储存小于30的时间间隔来计算
                List<double> rainList = new List<double>();  //临时储存雨量来计算
                List<double> thirtyList = new List<double>();  //临时储存30分钟雨强来判断
                List<double> indexList = new List<double>();  //临时储存序号来判断
                List<double> sDList = new List<double>();  //临时储存聪明时间
                List<double> sRList = new List<double>();  //临时聪明雨量


                for (int i = 1; i <= sheet.LastRowNum; i++)  //对工作表每一行(第二行起) 
                {
                    rowPre = sheet.GetRow(i - 1); //rowpre读入第i-1行数据
                    row = sheet.GetRow(i);   //row读入第i行数据
                    if (row != null)
                    {
                        for (int j = 0; j < row.LastCellNum; j++) //对工作表每一列  
                        {
                            ICell cell = row.GetCell(j);  //获取每一个单元格  
                            ICell cellPre = rowPre.GetCell(j);  //获取上一个单元格
                            XSSFFormulaEvaluator a = new XSSFFormulaEvaluator(workbook);
                            cell = a.EvaluateInCell(cell);
                            if (j == 0) //对工作表第一列  
                            {
                                ICell st = row.CreateCell(2);
                                ICell mt = row.CreateCell(3);
                                int tt = (int)cell.NumericCellValue; //时间的整数部分
                                Double sTime = cell.NumericCellValue - tt; //时间的小数部分
                                st.SetCellValue(sTime);
                                mt.CellFormula = "TEXT(INDIRECT(ADDRESS(ROW(),COLUMN()-1)),\"[m]\")";//转换成分钟
                            }
                            ICell dif = row.CreateCell(4); //计算时间差
                            dif.CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-1))-INDIRECT(ADDRESS(ROW()-1,COLUMN()-1))";
                            XSSFFormulaEvaluator b = new XSSFFormulaEvaluator(workbook);
                            dif = b.EvaluateInCell(dif);
                            if (dif.CellType == CellType.Numeric)
                            {
                                if (dif.NumericCellValue <= 0)
                                {
                                    dif.SetCellValue("#VALUE!");
                                }
                            }
                            if (j == 1)
                            {
                                ICell waterLine = row.CreateCell(5); //计算水位
                                waterLine.CellFormula = "MID(INDIRECT(ADDRESS(ROW(),COLUMN()-4)),7,4)";
                                ICell precipitation = row.CreateCell(6); //计算雨量
                                precipitation.CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-1))-INDIRECT(ADDRESS(ROW()-1,COLUMN()-1))";
                                ICell intensity = row.CreateCell(7); //计算雨强
                                intensity.CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-1))/INDIRECT(ADDRESS(ROW(),COLUMN()-3))*6";
                                ICell unitDrop = row.CreateCell(8); //计算单位雨强
                                unitDrop.CellFormula = "210.3+89*LOG(INDIRECT(ADDRESS(ROW(),COLUMN()-1)))";
                                ICell totalDrop = row.CreateCell(9); //计算时段雨强
                                totalDrop.CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-3))/10*INDIRECT(ADDRESS(ROW(),COLUMN()-1))";
                            }

                            //最后三列
                            if (i > 1)
                            {
                                if (j == 9)
                                {
                                    String getType = cell.CellType.ToString();
                                    String preType = cellPre.CellType.ToString();
                                    ICell totalE = rowPre.CreateCell(10); //创建E总列，储存E总
                                    ICell thirtyE = rowPre.CreateCell(11);  //创建30分钟雨强列
                                    ICell last = rowPre.CreateCell(12);  //创建降雨侵蚀力列
                                    if (getType == "Numeric")
                                    {
                                        if (preType == "Numeric")
                                        {
                                            Double difValue = dif.NumericCellValue;

                                            //   smart
                                            Double rainValue = row.GetCell(6).NumericCellValue; //雨量
                                            Double eValue = row.GetCell(7).NumericCellValue;  //雨强
                                            AlleList.Add(eValue);   //所有雨强
                                            AlldifList.Add(difValue);  //所有时间间隔
                                            AllrainList.Add(rainValue);  //所有雨量
                                            //

                                            if (difValue <= 360)
                                            {
                                                Double nValue = cell.NumericCellValue;  //时段雨强
                                                eList.Add(nValue);  // 储存时段雨强用量计算E总

                                                // 以下计算30分钟雨强
                                                difList.Add(difValue);  //合格的时间间隔
                                                Double difSum = difList.Sum();   // 需要判断的时间间隔之和
                                                if ( difSum < 30 )  //时间间隔之和小于30时
                                                {
                                                    countList.Add(difValue);  //添加小于30，且总数小于30的时间
                                                    rainList.Add(rainValue);  //添加合格的雨量
                                                }
                                                else if ( difSum >= 30 )
                                                {
                                                    if (difList.Count > 1 && difSum - difValue < 30)
                                                    {
                                                        Double needTime = 30 - countList.Sum(); //还差多少时间
                                                        Double needRain = rainValue / difValue * needTime; //还差多少雨量
                                                        Double thirtyV = (rainList.Sum() + needRain) / 5;  //计算出的30分钟雨强
                                                        thirtyList.Add(thirtyV);  //不输出，只添加
                                                        difList.Clear();    //清空简单时间间隔list
                                                        countList.Clear();  //清空临时时间间隔list
                                                        rainList.Clear();   //清空临时储存雨量list
                                                    }
                                                    else
                                                    {
                                                        thirtyList.Add(eValue);  //如果时间间隔大于30，,直接加入30雨强list准备进行对比
                                                        difList.Clear();    //清空简单时间间隔list
                                                        countList.Clear();  //清空临时时间间隔list
                                                        rainList.Clear();   //清空临时储存雨量list
                                                    }
                                                }

                                                if (i == sheet.LastRowNum)  //如果是最后一行
                                                {
                                                    if (AlldifList.Min() < 30)  //仅判断含有dif小于30min的数据
                                                    {
                                                        for (int eindex = 0; eindex < AlleList.Count; eindex++)
                                                        {
                                                            indexList.Add(eindex);  //储存序号值
                                                        }
                                                        foreach (int T in indexList)  //循环每一个值附近的数
                                                        {
                                                            int p = 0;
                                                            int q = T - 1;
                                                            for (p = T; p < AlldifList.Count; p++)
                                                            {
                                                                if (T == 0 && sDList.Sum() < 30)
                                                                {
                                                                    sDList.Add(AlldifList[p]);
                                                                    sRList.Add(AllrainList[p]);
                                                                }
                                                                else if (T > 0 && sDList.Sum() < 30)
                                                                {
                                                                    if (q >= 0 && AlleList[q] > AlleList[p])
                                                                    {
                                                                        sDList.Add(AlldifList[q]);
                                                                        sRList.Add(AllrainList[q]);
                                                                        q = q - 1;  //向上走一个
                                                                        p = p - 1;  //以维持不变
                                                                    }
                                                                    else
                                                                    {
                                                                        sDList.Add(AlldifList[p]);
                                                                        sRList.Add(AllrainList[p]);
                                                                        //q不变
                                                                    }
                                                                }
                                                                else if (sDList.Sum() >= 30)
                                                                {
                                                                    Double beforeTime = sDList.Sum() - sDList[sDList.Count - 1];
                                                                    Double beforeRain = sRList.Sum() - sRList[sRList.Count - 1];
                                                                    Double needTime = 30 - beforeTime; //还差多少时间
                                                                    Double needRain = sRList[sRList.Count - 1] / sDList[sDList.Count - 1] * needTime; //还差多少雨量
                                                                    Double thirtyV = (beforeRain + needRain) / 5;  //计算出的30分钟雨强
                                                                    thirtyList.Add(thirtyV);
                                                                    sRList.Clear();
                                                                    sDList.Clear();
                                                                    break;
                                                                }
                                                            }

                                                            // 特殊情况判断
                                                            if (p == AlldifList.Count && sDList.Sum() < 30)
                                                            {
                                                                if (q < 0)
                                                                {
                                                                    Double thirtyV = sRList.Sum() / 5;
                                                                    thirtyList.Add(thirtyV);
                                                                    sRList.Clear();
                                                                    sDList.Clear();
                                                                }
                                                                else if (q >= 0)
                                                                {
                                                                    for (int m = q; m >= 0; m--)
                                                                    {
                                                                        if (sDList.Sum() < 30)
                                                                        {
                                                                            sDList.Add(AlldifList[q]);
                                                                            sRList.Add(AllrainList[q]);
                                                                        }
                                                                        else if (sDList.Sum() >= 30)
                                                                        {
                                                                            Double beforeTime = sDList.Sum() - sDList[sDList.Count - 1];
                                                                            Double beforeRain = sRList.Sum() - sRList[sRList.Count - 1];
                                                                            Double needTime = 30 - beforeTime; //还差多少时间
                                                                            Double needRain = sRList[sRList.Count - 1] / sDList[sDList.Count - 1] * needTime; //还差多少雨量
                                                                            Double thirtyV = (beforeRain + needRain) / 5;  //计算出的30分钟雨强
                                                                            thirtyList.Add(thirtyV);
                                                                            sRList.Clear();
                                                                            sDList.Clear();
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    row.CreateCell(10).SetCellValue(eList.Sum());  //计算且输出E总
                                                    row.CreateCell(11).SetCellValue(thirtyList.Max());  //输出30分钟雨强
                                                    row.CreateCell(12).CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-2))*INDIRECT(ADDRESS(ROW(),COLUMN()-1))/100";  //输出降雨侵蚀力
                                                    eList.Clear();  //清空E总list
                                                    sRList.Clear();
                                                    sDList.Clear();
                                                    indexList.Clear();
                                                    AlldifList.Clear();
                                                    AllrainList.Clear();
                                                    AlleList.Clear();
                                                    thirtyList.Clear();
                                                    difList.Clear();
                                                    countList.Clear();
                                                    rainList.Clear();
                                                }
                                            }
                                            else if (difValue > 360 && eList.Sum() != 0)   //中间出现大间隔时，输出结果，且清空各种
                                            {
                                                // 智能判断
                                                if (AlldifList.Min() < 30)  //仅判断dif小于30min的数据
                                                {
                                                    for (int eindex = 0; eindex < AlleList.Count; eindex++)
                                                    {
                                                        indexList.Add(eindex);  //储存序号值
                                                    }
                                                    foreach (int T in indexList)  //循环每一个值附近的数
                                                    {
                                                        int p = 0;
                                                        int q = T - 1;
                                                        for (p = T; p < AlldifList.Count; p++)
                                                        {
                                                            if (T == 0 && sDList.Sum() < 30)
                                                            {
                                                                sDList.Add(AlldifList[p]);
                                                                sRList.Add(AllrainList[p]);
                                                            }
                                                            else if (T > 0 && sDList.Sum() < 30)
                                                            {
                                                                if (q >= 0 && AlleList[q] > AlleList[p])
                                                                {
                                                                    sDList.Add(AlldifList[q]);
                                                                    sRList.Add(AllrainList[q]);
                                                                    q = q - 1;  //向上走一个
                                                                    p = p - 1;  //以维持不变
                                                                }
                                                                else
                                                                {
                                                                    sDList.Add(AlldifList[p]);
                                                                    sRList.Add(AllrainList[p]);
                                                                    //q不变
                                                                }
                                                            }
                                                            else if (sDList.Sum() >= 30)
                                                            {
                                                                Double beforeTime = sDList.Sum() - sDList[sDList.Count - 1];
                                                                Double beforeRain = sRList.Sum() - sRList[sRList.Count - 1];
                                                                Double needTime = 30 - beforeTime; //还差多少时间
                                                                Double needRain = sRList[sRList.Count - 1] / sDList[sDList.Count - 1] * needTime; //还差多少雨量
                                                                Double thirtyV = (beforeRain + needRain) / 5;  //计算出的30分钟雨强
                                                                thirtyList.Add(thirtyV);
                                                                sRList.Clear();
                                                                sDList.Clear();
                                                                break;
                                                            }
                                                        }

                                                        // 特殊情况判断
                                                        if (p == AlldifList.Count && sDList.Sum() < 30)
                                                        {
                                                            if (q < 0)
                                                            {
                                                                Double thirtyV = sRList.Sum() / 5;
                                                                thirtyList.Add(thirtyV);
                                                                sRList.Clear();
                                                                sDList.Clear();
                                                            }
                                                            else if (q >= 0)
                                                            {
                                                                for (int m = q; m >= 0; m--)
                                                                {
                                                                    if (sDList.Sum() < 30)
                                                                    {
                                                                        sDList.Add(AlldifList[q]);
                                                                        sRList.Add(AllrainList[q]);
                                                                    }
                                                                    else if (sDList.Sum() >= 30)
                                                                    {
                                                                        Double beforeTime = sDList.Sum() - sDList[sDList.Count - 1];
                                                                        Double beforeRain = sRList.Sum() - sRList[sRList.Count - 1];
                                                                        Double needTime = 30 - beforeTime; //还差多少时间
                                                                        Double needRain = sRList[sRList.Count - 1] / sDList[sDList.Count - 1] * needTime; //还差多少雨量
                                                                        Double thirtyV = (beforeRain + needRain) / 5;  //计算出的30分钟雨强
                                                                        thirtyList.Add(thirtyV);
                                                                        sRList.Clear();
                                                                        sDList.Clear();
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                totalE.SetCellValue(eList.Sum());  //计算且输出E总
                                                thirtyE.SetCellValue(thirtyList.Max());  //输出30分钟雨强
                                                last.CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-2))*INDIRECT(ADDRESS(ROW(),COLUMN()-1))/100";  //输出降雨侵蚀力
                                                eList.Clear();  //清空E总list
                                                sRList.Clear();
                                                sDList.Clear();
                                                indexList.Clear();
                                                AlldifList.Clear();
                                                AllrainList.Clear();
                                                AlleList.Clear();
                                                thirtyList.Clear();
                                                difList.Clear();
                                                countList.Clear();
                                                rainList.Clear();
                                            }
                                        }
                                    }
                                    else if (getType == "Error") //如果错误就在上一行输出，且清空
                                    {
                                        if (preType == "Numeric")
                                        {
                                            if (eList.Sum() != 0)
                                            {
                                                if (AlldifList.Min() < 30)  //仅判断dif小于30min的数据
                                                {
                                                    for (int eindex = 0; eindex < AlleList.Count; eindex++)
                                                    {
                                                        indexList.Add(eindex);  //储存序号值
                                                    }
                                                    foreach (int T in indexList)  //循环每一个值附近的数
                                                    {
                                                        int p = 0;
                                                        int q = T - 1;
                                                        for (p = T; p < AlldifList.Count; p++)
                                                        {
                                                            if (T == 0 && sDList.Sum() < 30)
                                                            {
                                                                sDList.Add(AlldifList[p]);
                                                                sRList.Add(AllrainList[p]);
                                                            }
                                                            else if (T > 0 && sDList.Sum() < 30)
                                                            {
                                                                if (q >= 0 && AlleList[q] > AlleList[p])
                                                                {
                                                                    sDList.Add(AlldifList[q]);
                                                                    sRList.Add(AllrainList[q]);
                                                                    q = q - 1;  //向上走一个
                                                                    p = p - 1;  //以维持不变
                                                                }
                                                                else
                                                                {
                                                                    sDList.Add(AlldifList[p]);
                                                                    sRList.Add(AllrainList[p]);
                                                                    //q不变
                                                                }
                                                            }
                                                            else if (sDList.Sum() >= 30)
                                                            {
                                                                Double beforeTime = sDList.Sum() - sDList[sDList.Count - 1];
                                                                Double beforeRain = sRList.Sum() - sRList[sRList.Count - 1];
                                                                Double needTime = 30 - beforeTime; //还差多少时间
                                                                Double needRain = sRList[sRList.Count - 1] / sDList[sDList.Count - 1] * needTime; //还差多少雨量
                                                                Double thirtyV = (beforeRain + needRain) / 5;  //计算出的30分钟雨强
                                                                thirtyList.Add(thirtyV);
                                                                sRList.Clear();
                                                                sDList.Clear();
                                                                break;
                                                            }
                                                        }
                                                        //特殊情况（遍历到底了或者整个降雨过程不满30min）
                                                        if (p == AlldifList.Count && sDList.Sum() < 30)
                                                        {
                                                            if (q < 0)
                                                            {
                                                                Double thirtyV = sRList.Sum() / 5;
                                                                thirtyList.Add(thirtyV);
                                                                sRList.Clear();
                                                                sDList.Clear();
                                                            }
                                                            else if (q >= 0)
                                                            {
                                                                for (int m = q; m >= 0; m--)
                                                                {
                                                                    if (sDList.Sum() < 30)
                                                                    {
                                                                        sDList.Add(AlldifList[q]);
                                                                        sRList.Add(AllrainList[q]);
                                                                    }
                                                                    else if (sDList.Sum() >= 30)
                                                                    {
                                                                        Double beforeTime = sDList.Sum() - sDList[sDList.Count - 1];
                                                                        Double beforeRain = sRList.Sum() - sRList[sRList.Count - 1];
                                                                        Double needTime = 30 - beforeTime; //还差多少时间
                                                                        Double needRain = sRList[sRList.Count - 1] / sDList[sDList.Count - 1] * needTime; //还差多少雨量
                                                                        Double thirtyV = (beforeRain + needRain) / 5;  //计算出的30分钟雨强
                                                                        thirtyList.Add(thirtyV);
                                                                        sRList.Clear();
                                                                        sDList.Clear();
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                if (thirtyList.Count != 0 && eList.Sum() != 0)
                                                {
                                                    totalE.SetCellValue(eList.Sum());  //输出E总
                                                    thirtyE.SetCellValue(thirtyList.Max());  //输出30分钟雨强
                                                    last.CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-2))*INDIRECT(ADDRESS(ROW(),COLUMN()-1))/100";  //输出降雨侵蚀力
                                                    eList.Clear();
                                                    sRList.Clear();
                                                    sDList.Clear();
                                                    indexList.Clear();
                                                    AlldifList.Clear();
                                                    AllrainList.Clear();
                                                    AlleList.Clear();
                                                    thirtyList.Clear();
                                                    difList.Clear();
                                                    countList.Clear();
                                                    rainList.Clear();
                                                }
                                            }
                                            else
                                            {
                                                eList.Clear();
                                                sRList.Clear();
                                                sDList.Clear();
                                                indexList.Clear();
                                                AlldifList.Clear();
                                                AllrainList.Clear();
                                                AlleList.Clear();
                                                thirtyList.Clear();
                                                difList.Clear();
                                                countList.Clear();
                                                rainList.Clear();
                                            }
                                        }
                                    }
                                    XSSFFormulaEvaluator z = new XSSFFormulaEvaluator(workbook);
                                    last = z.EvaluateInCell(last);
                                }
                            }
                        }
                    }
                }

                //设置单元格样式：
                ICellStyle style1 = workbook.CreateCellStyle();
                style1.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;  //垂直居中
                style1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;  //水平居中
                for (int i = 0; i <= sheet.LastRowNum; i++)  //设置表格样式 
                {
                    row = sheet.GetRow(i);
                    for (int j = 2; j < row.LastCellNum; j++) //对工作表每一列  
                    {
                        sheet.SetColumnWidth(0, 18 * 256);
                        sheet.DefaultColumnWidth = 16;
                        ICell cell = row.GetCell(j);  //获取每一个单元格
                        if (cell != null)
                        {
                            cell.CellStyle = style1;
                        }
                    }
                }
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本  
            {
                workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook
                ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表
                IRow row;     //新建当前工作表行数据  
                IRow rowPre; //定义前一行
                IRow titleRow = sheet.GetRow(0);  //第一行
                ICell timeTitle = titleRow.CreateCell(titleRow.LastCellNum);
                timeTitle.SetCellValue("时刻数字计数法");
                ICell minutes = titleRow.CreateCell(titleRow.LastCellNum);
                minutes.SetCellValue("分钟");
                ICell difTitle = titleRow.CreateCell(titleRow.LastCellNum);
                difTitle.SetCellValue("时间间隔");
                ICell lineTitle = titleRow.CreateCell(titleRow.LastCellNum);
                lineTitle.SetCellValue("水位");
                ICell precipitationTitle = titleRow.CreateCell(titleRow.LastCellNum);
                precipitationTitle.SetCellValue("雨量");
                ICell intensityTitle = titleRow.CreateCell(titleRow.LastCellNum);
                intensityTitle.SetCellValue("雨强");
                ICell udTitle = titleRow.CreateCell(titleRow.LastCellNum);
                udTitle.SetCellValue("单位雨强");
                ICell tdTitle = titleRow.CreateCell(titleRow.LastCellNum);
                tdTitle.SetCellValue("时段雨强");
                ICell etTitle = titleRow.CreateCell(titleRow.LastCellNum);
                etTitle.SetCellValue("E总");
                ICell thirtyTitle = titleRow.CreateCell(titleRow.LastCellNum);
                thirtyTitle.SetCellValue("I30 cm/h");
                ICell RTitle = titleRow.CreateCell(titleRow.LastCellNum);
                RTitle.SetCellValue("R");
                int cloNum = titleRow.LastCellNum;
                int rowNum = sheet.LastRowNum;
                List<double> eList = new List<double>();  //储存时段雨强
                List<double> difList = new List<double>();  //储存简单时间间隔
                List<double> AlldifList = new List<double>();  //储存所有时间间隔
                List<double> AllrainList = new List<double>();  //储存所有雨量来判断
                List<double> AlleList = new List<double>();  //储存所有雨强
                List<double> countList = new List<double>();  //临时储存小于30的时间间隔来计算
                List<double> rainList = new List<double>();  //临时储存雨量来计算
                List<double> thirtyList = new List<double>();  //临时储存30分钟雨强来判断
                List<double> indexList = new List<double>();  //临时储存序号来判断
                List<double> sDList = new List<double>();  //临时储存聪明时间
                List<double> sRList = new List<double>();  //临时聪明雨量

 
                for (int i = 1; i <= sheet.LastRowNum; i++)  //对工作表每一行(第二行起) 
                {
                    rowPre = sheet.GetRow(i - 1); //rowpre读入第i-1行数据
                    row = sheet.GetRow(i);   //row读入第i行数据
                    if (row != null)
                    {
                        for (int j = 0; j < row.LastCellNum; j++) //对工作表每一列  
                        {
                            ICell cell = row.GetCell(j);  //获取每一个单元格  
                            ICell cellPre = rowPre.GetCell(j);  //获取上一个单元格
                            HSSFFormulaEvaluator a = new HSSFFormulaEvaluator(workbook);
                            cell = a.EvaluateInCell(cell);
                            if (j == 0) //对工作表第一列  
                            {
                                ICell st = row.CreateCell(2);
                                ICell mt = row.CreateCell(3);
                                int tt = (int)cell.NumericCellValue; //时间的整数部分
                                Double sTime = cell.NumericCellValue - tt; //时间的小数部分
                                st.SetCellValue(sTime);
                                mt.CellFormula = "TEXT(INDIRECT(ADDRESS(ROW(),COLUMN()-1)),\"[m]\")";//转换成分钟
                                
                            }
                            ICell dif = row.CreateCell(4); //计算时间差
                            dif.CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-1))-INDIRECT(ADDRESS(ROW()-1,COLUMN()-1))";
                            HSSFFormulaEvaluator b = new HSSFFormulaEvaluator(workbook);
                            dif = b.EvaluateInCell(dif);
                            if (dif.CellType == CellType.Numeric)
                            {
                                if (dif.NumericCellValue <= 0)
                                {
                                    dif.SetCellValue("#VALUE!");
                                }
                            }

                            if (j == 1)
                            {
                                ICell waterLine = row.CreateCell(5); //计算水位
                                waterLine.CellFormula = "MID(INDIRECT(ADDRESS(ROW(),COLUMN()-4)),7,4)";
                                ICell precipitation = row.CreateCell(6); //计算雨量
                                precipitation.CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-1))-INDIRECT(ADDRESS(ROW()-1,COLUMN()-1))";
                                ICell intensity = row.CreateCell(7); //计算雨强
                                intensity.CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-1))/INDIRECT(ADDRESS(ROW(),COLUMN()-3))*6";
                                ICell unitDrop = row.CreateCell(8); //计算单位雨强
                                unitDrop.CellFormula = "210.3+89*LOG(INDIRECT(ADDRESS(ROW(),COLUMN()-1)))";
                                ICell totalDrop = row.CreateCell(9); //计算时段雨强
                                totalDrop.CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-3))/10*INDIRECT(ADDRESS(ROW(),COLUMN()-1))";
                            }

                            //最后三列
                            if (i > 1)
                            {
                                if (j == 9)
                                {
                                    String getType = cell.CellType.ToString();
                                    String preType = cellPre.CellType.ToString();
                                    ICell totalE = rowPre.CreateCell(10); //创建E总列，储存E总
                                    ICell thirtyE = rowPre.CreateCell(11);  //创建30分钟雨强列
                                    ICell last = rowPre.CreateCell(12);  //创建降雨侵蚀力列
                                    if (getType == "Numeric")
                                    {
                                        if (preType == "Numeric")
                                        {
                                            Double difValue = dif.NumericCellValue;

                                            //   smart
                                            Double rainValue = row.GetCell(6).NumericCellValue; //雨量
                                            Double eValue = row.GetCell(7).NumericCellValue;  //雨强
                                            AlleList.Add(eValue);   //所有雨强
                                            AlldifList.Add(difValue);  //所有时间间隔
                                            AllrainList.Add(rainValue);  //所有雨量
                                            //

                                            if (difValue <= 360)
                                            {
                                                Double nValue = cell.NumericCellValue;  //时段雨强
                                                eList.Add(nValue);  // 储存时段雨强用量计算E总

                                                // 以下计算30分钟雨强
                                                difList.Add(difValue);  //合格的时间间隔
                                                Double difSum = difList.Sum();   // 需要判断的时间间隔之和
                                                if (difSum < 30)  //时间间隔之和小于30时
                                                {
                                                    countList.Add(difValue);  //添加小于30，且总数小于30的时间
                                                    rainList.Add(rainValue);  //添加合格的雨量
                                                }
                                                else if (difSum >= 30)
                                                {
                                                    if (difList.Count > 1 && difSum - difValue < 30)
                                                    {
                                                        Double needTime = 30 - countList.Sum(); //还差多少时间
                                                        Double needRain = rainValue / difValue * needTime; //还差多少雨量
                                                        Double thirtyV = (rainList.Sum() + needRain) / 5;  //计算出的30分钟雨强
                                                        thirtyList.Add(thirtyV);  //不输出，只添加
                                                        difList.Clear();    //清空简单时间间隔list
                                                        countList.Clear();  //清空临时时间间隔list
                                                        rainList.Clear();   //清空临时储存雨量list
                                                    }
                                                    else
                                                    {
                                                        thirtyList.Add(eValue);  //如果时间间隔大于30，,直接加入30雨强list准备进行对比
                                                        difList.Clear();    //清空简单时间间隔list
                                                        countList.Clear();  //清空临时时间间隔list
                                                        rainList.Clear();   //清空临时储存雨量list
                                                    }
                                                }

                                                if (i == sheet.LastRowNum)  //如果是最后一行
                                                {
                                                    if (AlldifList.Min() < 30)  //仅判断含有dif小于30min的数据
                                                    {
                                                        for (int eindex = 0; eindex < AlleList.Count; eindex++)
                                                        {
                                                            indexList.Add(eindex);  //储存序号值
                                                        }
                                                        foreach (int T in indexList)  //循环每一个值附近的数
                                                        {
                                                            int p = 0;
                                                            int q = T - 1;
                                                            for (p = T; p < AlldifList.Count; p++)
                                                            {
                                                                if (T == 0 && sDList.Sum() < 30)
                                                                {
                                                                    sDList.Add(AlldifList[p]);
                                                                    sRList.Add(AllrainList[p]);
                                                                }
                                                                else if (T > 0 && sDList.Sum() < 30)
                                                                {
                                                                    if (q >= 0 && AlleList[q] > AlleList[p])
                                                                    {
                                                                        sDList.Add(AlldifList[q]);
                                                                        sRList.Add(AllrainList[q]);
                                                                        q = q - 1;  //向上走一个
                                                                        p = p - 1;  //以维持不变
                                                                    }
                                                                    else
                                                                    {
                                                                        sDList.Add(AlldifList[p]);
                                                                        sRList.Add(AllrainList[p]);
                                                                        //q不变
                                                                    }
                                                                }
                                                                else if (sDList.Sum() >= 30)
                                                                {
                                                                    Double beforeTime = sDList.Sum() - sDList[sDList.Count - 1];
                                                                    Double beforeRain = sRList.Sum() - sRList[sRList.Count - 1];
                                                                    Double needTime = 30 - beforeTime; //还差多少时间
                                                                    Double needRain = sRList[sRList.Count - 1] / sDList[sDList.Count - 1] * needTime; //还差多少雨量
                                                                    Double thirtyV = (beforeRain + needRain) / 5;  //计算出的30分钟雨强
                                                                    thirtyList.Add(thirtyV);
                                                                    sRList.Clear();
                                                                    sDList.Clear();
                                                                    break;
                                                                }
                                                            }

                                                            // 特殊情况判断
                                                            if (p == AlldifList.Count && sDList.Sum() < 30)
                                                            {
                                                                if (q < 0)
                                                                {
                                                                    Double thirtyV = sRList.Sum() / 5;
                                                                    thirtyList.Add(thirtyV);
                                                                    sRList.Clear();
                                                                    sDList.Clear();
                                                                }
                                                                else if (q >= 0)
                                                                {
                                                                    for (int m = q; m >= 0; m--)
                                                                    {
                                                                        if (sDList.Sum() < 30)
                                                                        {
                                                                            sDList.Add(AlldifList[q]);
                                                                            sRList.Add(AllrainList[q]);
                                                                        }
                                                                        else if (sDList.Sum() >= 30)
                                                                        {
                                                                            Double beforeTime = sDList.Sum() - sDList[sDList.Count - 1];
                                                                            Double beforeRain = sRList.Sum() - sRList[sRList.Count - 1];
                                                                            Double needTime = 30 - beforeTime; //还差多少时间
                                                                            Double needRain = sRList[sRList.Count - 1] / sDList[sDList.Count - 1] * needTime; //还差多少雨量
                                                                            Double thirtyV = (beforeRain + needRain) / 5;  //计算出的30分钟雨强
                                                                            thirtyList.Add(thirtyV);
                                                                            sRList.Clear();
                                                                            sDList.Clear();
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    row.CreateCell(10).SetCellValue(eList.Sum());  //计算且输出E总
                                                    row.CreateCell(11).SetCellValue(thirtyList.Max());  //输出30分钟雨强
                                                    row.CreateCell(12).CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-2))*INDIRECT(ADDRESS(ROW(),COLUMN()-1))/100";  //输出降雨侵蚀力
                                                    eList.Clear();  //清空E总list
                                                    sRList.Clear();
                                                    sDList.Clear();
                                                    indexList.Clear();
                                                    AlldifList.Clear();
                                                    AllrainList.Clear();
                                                    AlleList.Clear();
                                                    thirtyList.Clear();
                                                    difList.Clear();
                                                    countList.Clear();
                                                    rainList.Clear();
                                                }
                                            }
                                            else if (difValue > 360 && eList.Sum() != 0)   //中间出现大间隔时，输出结果，且清空各种
                                            {
                                                // 智能判断
                                                if (AlldifList.Min() < 30)  //仅判断dif小于30min的数据
                                                {
                                                    for (int eindex = 0; eindex < AlleList.Count; eindex++)
                                                    {
                                                        indexList.Add(eindex);  //储存序号值
                                                    }
                                                    foreach (int T in indexList)  //循环每一个值附近的数
                                                    {
                                                        int p = 0;
                                                        int q = T - 1;
                                                        for (p = T; p < AlldifList.Count; p++)
                                                        {
                                                            if (T == 0 && sDList.Sum() < 30)
                                                            {
                                                                sDList.Add(AlldifList[p]);
                                                                sRList.Add(AllrainList[p]);
                                                            }
                                                            else if (T > 0 && sDList.Sum() < 30)
                                                            {
                                                                if (q >= 0 && AlleList[q] > AlleList[p])
                                                                {
                                                                    sDList.Add(AlldifList[q]);
                                                                    sRList.Add(AllrainList[q]);
                                                                    q = q - 1;  //向上走一个
                                                                    p = p - 1;  //以维持不变
                                                                }
                                                                else
                                                                {
                                                                    sDList.Add(AlldifList[p]);
                                                                    sRList.Add(AllrainList[p]);
                                                                    //q不变
                                                                }
                                                            }
                                                            else if (sDList.Sum() >= 30)
                                                            {
                                                                Double beforeTime = sDList.Sum() - sDList[sDList.Count - 1];
                                                                Double beforeRain = sRList.Sum() - sRList[sRList.Count - 1];
                                                                Double needTime = 30 - beforeTime; //还差多少时间
                                                                Double needRain = sRList[sRList.Count - 1] / sDList[sDList.Count - 1] * needTime; //还差多少雨量
                                                                Double thirtyV = (beforeRain + needRain) / 5;  //计算出的30分钟雨强
                                                                thirtyList.Add(thirtyV);
                                                                sRList.Clear();
                                                                sDList.Clear();
                                                                break;
                                                            }
                                                        }
                                                        // 特殊情况判断
                                                        if (p == AlldifList.Count && sDList.Sum() < 30)
                                                        {
                                                            if (q < 0)
                                                            {
                                                                Double thirtyV = sRList.Sum() / 5;
                                                                thirtyList.Add(thirtyV);
                                                                sRList.Clear();
                                                                sDList.Clear();
                                                            }
                                                            else if (q >= 0)
                                                            {
                                                                for (int m = q; m >= 0; m--)
                                                                {
                                                                    if (sDList.Sum() < 30)
                                                                    {
                                                                        sDList.Add(AlldifList[q]);
                                                                        sRList.Add(AllrainList[q]);
                                                                    }
                                                                    else if (sDList.Sum() >= 30)
                                                                    {
                                                                        Double beforeTime = sDList.Sum() - sDList[sDList.Count - 1];
                                                                        Double beforeRain = sRList.Sum() - sRList[sRList.Count - 1];
                                                                        Double needTime = 30 - beforeTime; //还差多少时间
                                                                        Double needRain = sRList[sRList.Count - 1] / sDList[sDList.Count - 1] * needTime; //还差多少雨量
                                                                        Double thirtyV = (beforeRain + needRain) / 5;  //计算出的30分钟雨强
                                                                        thirtyList.Add(thirtyV);
                                                                        sRList.Clear();
                                                                        sDList.Clear();
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                totalE.SetCellValue(eList.Sum());  //计算且输出E总
                                                thirtyE.SetCellValue(thirtyList.Max());  //输出30分钟雨强
                                                last.CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-2))*INDIRECT(ADDRESS(ROW(),COLUMN()-1))/100";  //输出降雨侵蚀力
                                                eList.Clear();  //清空E总list
                                                sRList.Clear();
                                                sDList.Clear();
                                                indexList.Clear();
                                                AlldifList.Clear();
                                                AllrainList.Clear();
                                                AlleList.Clear();
                                                thirtyList.Clear();
                                                difList.Clear();
                                                countList.Clear();
                                                rainList.Clear();
                                            }
                                        }
                                    }
                                    else if (getType == "Error") //如果错误就在上一行输出，且清空
                                    {
                                        if (preType == "Numeric")
                                        {
                                            if (eList.Sum() != 0)
                                            {
                                                if (AlldifList.Min() < 30)  //仅判断dif小于30min的数据
                                                {
                                                    for (int eindex = 0; eindex < AlleList.Count; eindex++)
                                                    {
                                                        indexList.Add(eindex);  //储存序号值
                                                    }
                                                    foreach (int T in indexList)  //循环每一个值附近的数
                                                    {
                                                        int p = 0;
                                                        int q = T - 1;
                                                        for (p = T; p < AlldifList.Count; p++)
                                                        {
                                                            if (T == 0 && sDList.Sum() < 30)
                                                            {
                                                                sDList.Add(AlldifList[p]);
                                                                sRList.Add(AllrainList[p]);
                                                            }
                                                            else if (T > 0 && sDList.Sum() < 30)
                                                            {
                                                                if (q >= 0 && AlleList[q] > AlleList[p])
                                                                {
                                                                    sDList.Add(AlldifList[q]);
                                                                    sRList.Add(AllrainList[q]);
                                                                    q = q - 1;  //向上走一个
                                                                    p = p - 1;  //以维持不变
                                                                }
                                                                else
                                                                {
                                                                    sDList.Add(AlldifList[p]);
                                                                    sRList.Add(AllrainList[p]);
                                                                    //q不变
                                                                }
                                                            }
                                                            else if (sDList.Sum() >= 30)
                                                            {
                                                                Double beforeTime = sDList.Sum() - sDList[sDList.Count - 1];
                                                                Double beforeRain = sRList.Sum() - sRList[sRList.Count - 1];
                                                                Double needTime = 30 - beforeTime; //还差多少时间
                                                                Double needRain = sRList[sRList.Count - 1] / sDList[sDList.Count - 1] * needTime; //还差多少雨量
                                                                Double thirtyV = (beforeRain + needRain) / 5;  //计算出的30分钟雨强
                                                                thirtyList.Add(thirtyV);
                                                                sRList.Clear();
                                                                sDList.Clear();
                                                                break;
                                                            }
                                                        }
                                                        //特殊情况（遍历到底了或者整个降雨过程不满30min）
                                                        if (p == AlldifList.Count && sDList.Sum() < 30)
                                                        {
                                                            if (q < 0)
                                                            {
                                                                Double thirtyV = sRList.Sum() / 5;
                                                                thirtyList.Add(thirtyV);
                                                                sRList.Clear();
                                                                sDList.Clear();
                                                            }
                                                            else if (q >= 0)
                                                            {
                                                                for (int m = q; m >= 0; m--)
                                                                {
                                                                    if (sDList.Sum() < 30)
                                                                    {
                                                                        sDList.Add(AlldifList[q]);
                                                                        sRList.Add(AllrainList[q]);
                                                                    }
                                                                    else if (sDList.Sum() >= 30)
                                                                    {
                                                                        Double beforeTime = sDList.Sum() - sDList[sDList.Count - 1];
                                                                        Double beforeRain = sRList.Sum() - sRList[sRList.Count - 1];
                                                                        Double needTime = 30 - beforeTime; //还差多少时间
                                                                        Double needRain = sRList[sRList.Count - 1] / sDList[sDList.Count - 1] * needTime; //还差多少雨量
                                                                        Double thirtyV = (beforeRain + needRain) / 5;  //计算出的30分钟雨强
                                                                        thirtyList.Add(thirtyV);
                                                                        sRList.Clear();
                                                                        sDList.Clear();
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                if (thirtyList.Count != 0 && eList.Sum() != 0)
                                                {
                                                    totalE.SetCellValue(eList.Sum());  //输出E总
                                                    thirtyE.SetCellValue(thirtyList.Max());  //输出30分钟雨强
                                                    last.CellFormula = "INDIRECT(ADDRESS(ROW(),COLUMN()-2))*INDIRECT(ADDRESS(ROW(),COLUMN()-1))/100";  //输出降雨侵蚀力
                                                    eList.Clear();
                                                    sRList.Clear();
                                                    sDList.Clear();
                                                    indexList.Clear();
                                                    AlldifList.Clear();
                                                    AllrainList.Clear();
                                                    AlleList.Clear();
                                                    thirtyList.Clear();
                                                    difList.Clear();
                                                    countList.Clear();
                                                    rainList.Clear();
                                                }
                                            }
                                            else
                                            {
                                                eList.Clear();
                                                sRList.Clear();
                                                sDList.Clear();
                                                indexList.Clear();
                                                AlldifList.Clear();
                                                AllrainList.Clear();
                                                AlleList.Clear();
                                                thirtyList.Clear();
                                                difList.Clear();
                                                countList.Clear();
                                                rainList.Clear();
                                            }
                                        }
                                    }
                                    HSSFFormulaEvaluator z = new HSSFFormulaEvaluator(workbook);
                                    last = z.EvaluateInCell(last);
                                }
                            }
                        }
                    }
                }

                //设置单元格样式：
                ICellStyle style1 = workbook.CreateCellStyle();
                style1.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;  //垂直居中
                style1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;  //水平居中
                for (int i = 0; i <= sheet.LastRowNum; i++)  //设置表格样式 
                {
                    row = sheet.GetRow(i);
                    for (int j = 2; j < row.LastCellNum; j++) //对工作表每一列  
                    {
                        sheet.SetColumnWidth(0, 18 * 256);
                        sheet.DefaultColumnWidth = 16;
                        ICell cell = row.GetCell(j);  //获取每一个单元格
                        if (cell != null)
                        {
                            cell.CellStyle = style1;
                        }
                    }
                }
            }
        }
        #endregion

        }

        ImportCore get = new ImportCore();

        public Form1()
        {
            InitializeComponent();
        }
        

        private void button1_Click(object sender, EventArgs e)
        {
            var opdiag = new OpenFileDialog();
            opdiag.Filter = "(*.et;*.xls;*.xlsx)|*.et;*.xls;*.xlsx|all|*.*"; //筛选、设定文件显示类型

            if (opdiag.ShowDialog() == DialogResult.OK)
            {

                get._filePath = opdiag.FileName;
            }
        }

        Thread th;

        private void UpdateLabel(object str)
        {
            get.Calculate();
            if (label1.InvokeRequired)//不同线程为true，所以这里是true
            {
                BeginInvoke(new Action<string>(x => { label1.Text = x.ToString(); label1.ForeColor = Color.Green; label1.Location = new Point((this.Width - label1.Width - 12) / 2, (this.Height - label1.Height) / 2);button2.Enabled = true; }), str);
                th.Abort();
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (get._filePath != null)
            {
                label1.Text = "正在计算中,请勿退出";
                label1.ForeColor = Color.Red;
                label1.Location = new Point((this.Width - label1.Width - 12) / 2, (this.Height - label1.Height) / 2);
                button2.Enabled = false;
                th = new Thread(UpdateLabel);
                th.IsBackground = true;
                th.Start("计算完成");
            }
            else
            {
                label1.Text = "请选择文件";
                label1.ForeColor = Color.Red;
                label1.Location = new Point((this.Width - label1.Width - 12) / 2, (this.Height - label1.Height) / 2);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (th != null)
            {
                if (th.IsAlive)
                {
                    label1.Text = "计算中，请等待...";
                    label1.ForeColor = Color.Red;
                    label1.Location = new Point((this.Width - label1.Width - 12) / 2, (this.Height - label1.Height) / 2);
                }
                else if (!th.IsAlive)
                {
                    SaveFileDialog sa = new SaveFileDialog();
                    if (get._filePath.IndexOf(".xlsx") > 0)
                    {
                        sa.Filter = "Excel文件 (*.xlsx)|*.xlsx";
                    }
                    else if (get._filePath.IndexOf(".xls") > 0)
                    {
                        sa.Filter = "Excel文件 (*.xls)|*.xls";
                    }
                    sa.RestoreDirectory = true;
                    if (DialogResult.OK == sa.ShowDialog())
                    {   //保存事件   
                        FileStream file1 = new FileStream(sa.FileName, FileMode.Create);
                        if (get.workbook != null)
                        {
                            get.workbook.Write(file1);
                            get.fileStream.Close();
                            label1.Text = "生成成功";
                            label1.ForeColor = Color.Green;
                            label1.Location = new Point((this.Width - label1.Width - 12) / 2, (this.Height - label1.Height) / 2);
                        }
                    }
                }
            }
            else
            {
                label1.Text = "请选择文件";
                label1.ForeColor = Color.Red;
                label1.Location = new Point((this.Width - label1.Width-12) / 2, (this.Height - label1.Height) / 2);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

    }
}
