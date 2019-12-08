//PrintReport.cs,把mdb的数据导入到EXCEL中显示
//添加net引用：Microsoft.Office.Interop.Excel

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;//引用这个才能使用Missing字段
using System.Diagnostics;
using System.Windows;
//File
using System.IO;
//NPOI
using NPOI.HSSF.UserModel;
using Spire.Xls;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;

namespace Seekya
{
    class PrintReport
    {
        //public Excel.Application app;
        //public Excel.Workbooks wbs;
        //public Excel.Workbook wb;

        //直接打印
        public void ReportPrintDirect(string name, string gender, string age, string id, string instrumentType, string submitDoctor, string firstVisit, string hb, string hospital, string rbc, string co, string co2, string testDateLine, string department, string userDefine1, string userDefine2, string userDefine3, string userDefine4, string userDefine5, string userDefine6, string checkDoctor, string reportDoctor, string reportTime, string remark1, string remark2)
        {
            WriteCopyTemplatedirect(name, gender, age, id, instrumentType, submitDoctor, firstVisit, hb, hospital, rbc, co, co2, testDateLine, department, userDefine1, userDefine2, userDefine3, userDefine4, userDefine5, userDefine6, checkDoctor, reportDoctor, reportTime, remark1, remark2);
            //直接打印
            //Process.Start(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls");

        }
        //手动打印
        public void ReportPrintHand(string name, string gender, string age, string id, string instrumentType, string submitDoctor, string firstVisit, string hb, string hospital, string rbc, string co, string co2, string testDateLine, string department, string userDefine1, string userDefine2, string userDefine3, string userDefine4, string userDefine5, string userDefine6, string checkDoctor, string reportDoctor, string reportTime, string remark1, string remark2)
        {
            WriteCopyTemplatemanual(name, gender, age, id, instrumentType, submitDoctor, firstVisit, hb, hospital, rbc, co, co2, testDateLine, department, userDefine1, userDefine2, userDefine3, userDefine4, userDefine5, userDefine6, checkDoctor, reportDoctor, reportTime, remark1, remark2);
            //间接打印
            //Process.Start(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls");

        }

        //public void Open(string FileName)
        //{
        //    app = new Excel.Application();
        //    wbs = app.Workbooks;
        //    wb = wbs.Add(FileName);
        //    //wb = wbs.Open(FileName,  0, true, 5,"", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true,Type.Missing,Type.Missing);
        //    //wb = wbs.Open(FileName,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Excel.XlPlatform.xlWindows,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing);
        //}

        //往临时报告中写数据
        public void WriteCopyTemplatedirect(string name, string gender, string age, string id, string instrumentType, string submitDoctor, string firstVisit, string hb, string hospital, string rbc, string co, string co2, string testDateLine, string department, string userDefine1, string userDefine2, string userDefine3, string userDefine4, string userDefine5, string userDefine6, string checkDoctor, string reportDoctor, string reportTime, string remark1, string remark2)
        {

            ////模板文件  
            //string TempletFileName = null;//System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template.xls";
            //string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\print.txt";

            ////读打印模板名
            //try
            //{
            //    StreamReader sr = new StreamReader(pathString, Encoding.GetEncoding("gb2312"));

            //    sr.ReadLine();
            //    TempletFileName = sr.ReadLine();

            //    sr.Close();

            //}
            //catch (Exception ex)
            //{
            //    // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            //}

            //string str = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\Template\\template.xls";
            //Open(str);
            //Excel.Worksheet ws = (Excel.Worksheet)app.ActiveSheet;

            //DataTable dt = new DataTable();
            //dt.Columns.Add("name", typeof(string));
            //dt.Columns.Add("age", typeof(string));
            //dt.Columns.Add("zyh", typeof(string));
            //dt.Columns.Add("sex", typeof(string));
            //dt.Columns.Add("yqxh", typeof(string));
            //dt.Columns.Add("cbzd", typeof(string));
            //dt.Columns.Add("sjys", typeof(string));
            //dt.Columns.Add("hb", typeof(string));
            //dt.Columns.Add("yymc", typeof(string));
            //dt.Columns.Add("rbc", typeof(string));
            //dt.Columns.Add("CO", typeof(string));
            //dt.Columns.Add("eyht", typeof(string));
            //dt.Columns.Add("jyrq", typeof(string));
            //dt.Columns.Add("ksmc", typeof(string));
            //dt.Columns.Add("dyyi", typeof(string));
            //dt.Columns.Add("dyer", typeof(string));
            //dt.Columns.Add("dysan", typeof(string));
            //dt.Columns.Add("dysi", typeof(string));
            //dt.Columns.Add("dywu", typeof(string));
            //dt.Columns.Add("dyliu", typeof(string));
            //dt.Columns.Add("fhys", typeof(string));
            //dt.Columns.Add("bgys", typeof(string));
            //dt.Columns.Add("bgsj", typeof(string));
            //dt.Columns.Add("ldgd", typeof(string));
            //dt.Columns.Add("eyhtgd", typeof(string));
            //DataRow dr = dt.NewRow();
            //dr["name"] = name;
            //dr["age"] = age;
            //dr["zyh"] = id;
            //dr["sex"] = gender;
            //dr["yqxh"] = instrumentType;
            //dr["cbzd"] = firstVisit;
            //dr["sjys"] = submitDoctor;
            //dr["hb"] = hb;
            //dr["yymc"] = hospital;
            //dr["rbc"] = rbc;
            //dr["CO"] = co;
            //dr["eyht"] = co2;
            //dr["jyrq"] = testDateLine;
            //dr["ksmc"] = department;
            //dr["dyyi"] = userDefine1;
            //dr["dyer"] = userDefine2;
            //dr["dysan"] = userDefine3;
            //dr["dysi"] = userDefine4;
            //dr["dywu"] = userDefine5;
            //dr["dyliu"] = userDefine6;
            //dr["fhys"] = checkDoctor;
            //dr["bgys"] = reportDoctor;
            //dr["bgsj"] = reportTime;
            //dr["ldgd"] = remark1;
            //dr["eyhtgd"] = remark2;
            //dt.Rows.Add(dr);


            //int nameCellCount = app.ActiveWorkbook.Names.Count;//获得命名单元格的总数
            //int[] nameCellRow = new int[nameCellCount];//某个命名单元格的行
            //int[] nameCellColumn = new int[nameCellCount];//某个命名单元格的列
            //string[] nameCellName = new string[nameCellCount];//某个命名单元格的自定义名称，比如 工资

            //string strName;
            //string tmp;
            //int nameCellIdx = 0;
            //for (int i = 0; i < nameCellCount; i++)
            //{
            //    strName = app.ActiveWorkbook.Names.Item(i + 1).Name;
            //    app.Goto(strName);
            //    nameCellColumn[nameCellIdx] = app.ActiveCell.Column;
            //    nameCellRow[nameCellIdx] = app.ActiveCell.Row;
            //    nameCellName[nameCellIdx] = strName;
            //    nameCellIdx++;//真实的循环的命名单元格序号
            //}
            //for (int index = 0; index < nameCellCount; index++)
            //{
            //    tmp = dt.Rows[0][nameCellName[index]].ToString();
            //    ws.Cells[nameCellRow[index], nameCellColumn[index]] = tmp;
            //}
            //try
            //{
            //    string excelName = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\Template\\" + name + ".xls";
            //    wb.SaveAs(excelName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //}
            //catch (Exception eee)
            //{
            //    System.Windows.MessageBox.Show("ERROR26:" + eee.Message);
            //}
            //wb.Close(Type.Missing, Type.Missing, Type.Missing);
            //wbs.Close();
            //app.Quit();
            //wb = null;
            //wbs = null;
            //app = null;
            //GC.Collect();

            ////导出文件  
            ////string ReportFileName = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls";
            //FileStream file = new FileStream(TempletFileName, FileMode.Open, FileAccess.Read);
            //HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
            //HSSFSheet ws = hssfworkbook.GetSheet("Sheet1");
            ////添加或修改WorkSheet里的数据  
            ////System.Data.DataTable dt = new System.Data.DataTable();  
            ////dt = DbHelperMySQLnew.Query("select * from t_jb_info where id='" + id + "'").Tables[0];  
            //#region

            ////姓名
            //HSSFRow row = ws.GetRow(1);
            //HSSFCell cell = row.GetCell(19);
            //cell.SetCellValue(name);

            ////性别
            //row = ws.GetRow(2);
            //cell = row.GetCell(19);
            //cell.SetCellValue(gender);

            ////年龄
            //row = ws.GetRow(3);
            //cell = row.GetCell(19);
            //cell.SetCellValue(age);

            ////住院号
            //row = ws.GetRow(4);
            //cell = row.GetCell(19);
            //cell.SetCellValue(id);

            ////仪器型号
            //row = ws.GetRow(5);
            //cell = row.GetCell(19);
            //cell.SetCellValue(instrumentType);

            ////送检医生
            //row = ws.GetRow(6);
            //cell = row.GetCell(19);
            //cell.SetCellValue(submitDoctor);

            ////初步诊断
            //row = ws.GetRow(7);
            //cell = row.GetCell(19);
            //cell.SetCellValue(firstVisit);

            ////血红蛋白浓度
            //row = ws.GetRow(8);
            //cell = row.GetCell(19);
            //cell.SetCellValue(hb);

            ////医院名称
            //row = ws.GetRow(9);
            //cell = row.GetCell(19);
            //cell.SetCellValue(hospital);

            ////红细胞寿命
            //row = ws.GetRow(10);
            //cell = row.GetCell(19);
            //cell.SetCellValue(rbc);

            ////一氧化碳浓度
            //row = ws.GetRow(11);
            //cell = row.GetCell(19);
            //cell.SetCellValue(co);

            ////二氧化碳浓度
            //row = ws.GetRow(12);
            //cell = row.GetCell(19);
            //cell.SetCellValue(co2);

            ////检验日期
            //row = ws.GetRow(13);
            //cell = row.GetCell(19);
            //cell.SetCellValue(testDateLine);

            ////科室名称
            //row = ws.GetRow(14);
            //cell = row.GetCell(19);
            //cell.SetCellValue(department);

            ////定义1
            //row = ws.GetRow(15);
            //cell = row.GetCell(19);
            //cell.SetCellValue(userDefine1);

            ////定义2
            //row = ws.GetRow(16);
            //cell = row.GetCell(19);
            //cell.SetCellValue(userDefine2);

            ////定义3
            //row = ws.GetRow(17);
            //cell = row.GetCell(19);
            //cell.SetCellValue(userDefine3);

            ////定义4
            //row = ws.GetRow(18);
            //cell = row.GetCell(19);
            //cell.SetCellValue(userDefine4);

            ////定义5
            //row = ws.GetRow(19);
            //cell = row.GetCell(19);
            //cell.SetCellValue(userDefine5);

            ////定义6
            //row = ws.GetRow(20);
            //cell = row.GetCell(19);
            //cell.SetCellValue(userDefine6);

            ////复核医生
            //row = ws.GetRow(21);
            //cell = row.GetCell(19);
            //cell.SetCellValue(checkDoctor);

            ////报告医生
            //row = ws.GetRow(22);
            //cell = row.GetCell(19);
            //cell.SetCellValue(reportDoctor);

            ////报告时间
            //row = ws.GetRow(23);
            //cell = row.GetCell(19);
            //cell.SetCellValue(reportTime);

            ////零点过大
            //row = ws.GetRow(24);
            //cell = row.GetCell(19);
            //cell.SetCellValue(remark1);

            ////CO2过低
            //row = ws.GetRow(25);
            //cell = row.GetCell(19);
            //cell.SetCellValue(remark2);

            ////ws.GetRow(1).GetCell(1).SetCellValue("5");  
            //#endregion
            //ws.ForceFormulaRecalculation = true;

            //using (FileStream filess = File.OpenWrite(TempletFileName))
            //{
            //    hssfworkbook.Write(filess);
            //}

            //Process.Start(TempletFileName);
            string datetime1 = testDateLine.Substring(0, 4) + testDateLine.Substring(5, 2) + testDateLine.Substring(8, 2);
            string datetime2 = reportTime.Substring(0, 2) + reportTime.Substring(3, 2) + reportTime.Substring(6, 2);
            Workbook workbook = new Workbook();
            //workbook.LoadFromFile(@"E:\软件开发\【1】红细胞寿命测定仪上位机软件汇总-20181011更新\6.新版本源代码\红细胞寿命测定仪1.0版本-20190330\源代码 - Buffer版-升级测试版\Seekya\bin\Debug\Data\Template"+name+".xls");
            workbook.LoadFromFile(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\Template\\" + name +"("+datetime1+datetime2+")"+ ".xls");
            workbook.PrintDocument.PrintController = new StandardPrintController();
            workbook.PrintDocument.Print();


        }
        public void WriteCopyTemplatemanual(string name, string gender, string age, string id, string instrumentType, string submitDoctor, string firstVisit, string hb, string hospital, string rbc, string co, string co2, string testDateLine, string department, string userDefine1, string userDefine2, string userDefine3, string userDefine4, string userDefine5, string userDefine6, string checkDoctor, string reportDoctor, string reportTime, string remark1, string remark2)
        {

            ////模板文件  
            //string TempletFileName = null;//System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template.xls";
            //string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\print.txt";

            ////读打印模板名
            //try
            //{
            //    StreamReader sr = new StreamReader(pathString, Encoding.GetEncoding("gb2312"));

            //    sr.ReadLine();
            //    TempletFileName = sr.ReadLine();

            //    sr.Close();

            //}
            //catch (Exception ex)
            //{
            //    // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            //}

            //string str = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\Template\\template.xls";
            //Open(str);
            //Excel.Worksheet ws = (Excel.Worksheet)app.ActiveSheet;

            //DataTable dt = new DataTable();
            //dt.Columns.Add("name", typeof(string));
            //dt.Columns.Add("age", typeof(string));
            //dt.Columns.Add("zyh", typeof(string));
            //dt.Columns.Add("sex", typeof(string));
            //dt.Columns.Add("yqxh", typeof(string));
            //dt.Columns.Add("cbzd", typeof(string));
            //dt.Columns.Add("sjys", typeof(string));
            //dt.Columns.Add("hb", typeof(string));
            //dt.Columns.Add("yymc", typeof(string));
            //dt.Columns.Add("rbc", typeof(string));
            //dt.Columns.Add("CO", typeof(string));
            //dt.Columns.Add("eyht", typeof(string));
            //dt.Columns.Add("jyrq", typeof(string));
            //dt.Columns.Add("ksmc", typeof(string));
            //dt.Columns.Add("dyyi", typeof(string));
            //dt.Columns.Add("dyer", typeof(string));
            //dt.Columns.Add("dysan", typeof(string));
            //dt.Columns.Add("dysi", typeof(string));
            //dt.Columns.Add("dywu", typeof(string));
            //dt.Columns.Add("dyliu", typeof(string));
            //dt.Columns.Add("fhys", typeof(string));
            //dt.Columns.Add("bgys", typeof(string));
            //dt.Columns.Add("bgsj", typeof(string));
            //dt.Columns.Add("ldgd", typeof(string));
            //dt.Columns.Add("eyhtgd", typeof(string));
            //DataRow dr = dt.NewRow();
            //dr["name"] = name;
            //dr["age"] = age;
            //dr["zyh"] = id;
            //dr["sex"] = gender;
            //dr["yqxh"] = instrumentType;
            //dr["cbzd"] = firstVisit;
            //dr["sjys"] = submitDoctor;
            //dr["hb"] = hb;
            //dr["yymc"] = hospital;
            //dr["rbc"] = rbc;
            //dr["CO"] = co;
            //dr["eyht"] = co2;
            //dr["jyrq"] = testDateLine;
            //dr["ksmc"] = department;
            //dr["dyyi"] = userDefine1;
            //dr["dyer"] = userDefine2;
            //dr["dysan"] = userDefine3;
            //dr["dysi"] = userDefine4;
            //dr["dywu"] = userDefine5;
            //dr["dyliu"] = userDefine6;
            //dr["fhys"] = checkDoctor;
            //dr["bgys"] = reportDoctor;
            //dr["bgsj"] = reportTime;
            //dr["ldgd"] = remark1;
            //dr["eyhtgd"] = remark2;
            //dt.Rows.Add(dr);


            //int nameCellCount = app.ActiveWorkbook.Names.Count;//获得命名单元格的总数
            //int[] nameCellRow = new int[nameCellCount];//某个命名单元格的行
            //int[] nameCellColumn = new int[nameCellCount];//某个命名单元格的列
            //string[] nameCellName = new string[nameCellCount];//某个命名单元格的自定义名称，比如 工资

            //string strName;
            //string tmp;
            //int nameCellIdx = 0;
            //for (int i = 0; i < nameCellCount; i++)
            //{
            //    strName = app.ActiveWorkbook.Names.Item(i + 1).Name;
            //    app.Goto(strName);
            //    nameCellColumn[nameCellIdx] = app.ActiveCell.Column;
            //    nameCellRow[nameCellIdx] = app.ActiveCell.Row;
            //    nameCellName[nameCellIdx] = strName;
            //    nameCellIdx++;//真实的循环的命名单元格序号
            //}
            //for (int index = 0; index < nameCellCount; index++)
            //{
            //    tmp = dt.Rows[0][nameCellName[index]].ToString();
            //    ws.Cells[nameCellRow[index], nameCellColumn[index]] = tmp;
            //}
            //try
            //{
            //    string excelName = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\Template\\" + name + ".xls";
            //    wb.SaveAs(excelName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //}
            //catch (Exception eee)
            //{
            //    System.Windows.MessageBox.Show("ERROR27:" + eee.Message);
            //}
            //wb.Close(Type.Missing, Type.Missing, Type.Missing);
            //wbs.Close();
            //app.Quit();
            //wb = null;
            //wbs = null;
            ////app = null;
            //GC.Collect();
            //PublicMethod.Kill(app);
            ////导出文件  
            ////string ReportFileName = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls";
            //FileStream file = new FileStream(TempletFileName, FileMode.Open, FileAccess.Read);
            //HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
            //HSSFSheet ws = hssfworkbook.GetSheet("Sheet1");
            ////添加或修改WorkSheet里的数据  
            ////System.Data.DataTable dt = new System.Data.DataTable();  
            ////dt = DbHelperMySQLnew.Query("select * from t_jb_info where id='" + id + "'").Tables[0];  
            //#region

            ////姓名
            //HSSFRow row = ws.GetRow(1);
            //HSSFCell cell = row.GetCell(19);
            //cell.SetCellValue(name);

            ////性别
            //row = ws.GetRow(2);
            //cell = row.GetCell(19);
            //cell.SetCellValue(gender);

            ////年龄
            //row = ws.GetRow(3);
            //cell = row.GetCell(19);
            //cell.SetCellValue(age);

            ////住院号
            //row = ws.GetRow(4);
            //cell = row.GetCell(19);
            //cell.SetCellValue(id);

            ////仪器型号
            //row = ws.GetRow(5);
            //cell = row.GetCell(19);
            //cell.SetCellValue(instrumentType);

            ////送检医生
            //row = ws.GetRow(6);
            //cell = row.GetCell(19);
            //cell.SetCellValue(submitDoctor);

            ////初步诊断
            //row = ws.GetRow(7);
            //cell = row.GetCell(19);
            //cell.SetCellValue(firstVisit);

            ////血红蛋白浓度
            //row = ws.GetRow(8);
            //cell = row.GetCell(19);
            //cell.SetCellValue(hb);

            ////医院名称
            //row = ws.GetRow(9);
            //cell = row.GetCell(19);
            //cell.SetCellValue(hospital);

            ////红细胞寿命
            //row = ws.GetRow(10);
            //cell = row.GetCell(19);
            //cell.SetCellValue(rbc);

            ////一氧化碳浓度
            //row = ws.GetRow(11);
            //cell = row.GetCell(19);
            //cell.SetCellValue(co);

            ////二氧化碳浓度
            //row = ws.GetRow(12);
            //cell = row.GetCell(19);
            //cell.SetCellValue(co2);

            ////检验日期
            //row = ws.GetRow(13);
            //cell = row.GetCell(19);
            //cell.SetCellValue(testDateLine);

            ////科室名称
            //row = ws.GetRow(14);
            //cell = row.GetCell(19);
            //cell.SetCellValue(department);

            ////定义1
            //row = ws.GetRow(15);
            //cell = row.GetCell(19);
            //cell.SetCellValue(userDefine1);

            ////定义2
            //row = ws.GetRow(16);
            //cell = row.GetCell(19);
            //cell.SetCellValue(userDefine2);

            ////定义3
            //row = ws.GetRow(17);
            //cell = row.GetCell(19);
            //cell.SetCellValue(userDefine3);

            ////定义4
            //row = ws.GetRow(18);
            //cell = row.GetCell(19);
            //cell.SetCellValue(userDefine4);

            ////定义5
            //row = ws.GetRow(19);
            //cell = row.GetCell(19);
            //cell.SetCellValue(userDefine5);

            ////定义6
            //row = ws.GetRow(20);
            //cell = row.GetCell(19);
            //cell.SetCellValue(userDefine6);

            ////复核医生
            //row = ws.GetRow(21);
            //cell = row.GetCell(19);
            //cell.SetCellValue(checkDoctor);

            ////报告医生
            //row = ws.GetRow(22);
            //cell = row.GetCell(19);
            //cell.SetCellValue(reportDoctor);

            ////报告时间
            //row = ws.GetRow(23);
            //cell = row.GetCell(19);
            //cell.SetCellValue(reportTime);

            ////零点过大
            //row = ws.GetRow(24);
            //cell = row.GetCell(19);
            //cell.SetCellValue(remark1);

            ////CO2过低
            //row = ws.GetRow(25);
            //cell = row.GetCell(19);
            //cell.SetCellValue(remark2);

            ////ws.GetRow(1).GetCell(1).SetCellValue("5");  
            //#endregion
            //ws.ForceFormulaRecalculation = true;

            //using (FileStream filess = File.OpenWrite(TempletFileName))
            //{
            //    hssfworkbook.Write(filess);
            //}
            string datetime1 = testDateLine.Substring(0, 4) + testDateLine.Substring(5, 2) + testDateLine.Substring(8, 2);
            string datetime2 = reportTime.Substring(0, 2) + reportTime.Substring(3, 2) + reportTime.Substring(6, 2);
            string filename = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\Template\\" + name+"("+datetime1+datetime2+")" + ".xls";

            try
            {
                Process.Start(filename);
            }
            catch (Exception e)
            {
                System.Windows.MessageBox.Show("ERROR28:" + e.Message);
            }
        }

    }
    public class PublicMethod
    {
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
        {
            IntPtr t = new IntPtr(excel.Hwnd);//得到这个句柄，具体作用是得到这块内存入口 

            int k = 0;
            GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
            p.Kill();     //关闭进程k
        }
    }
}
