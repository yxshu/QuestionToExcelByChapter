using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Threading;


namespace QuestionToExcelByChapter
{
    //此类用于从onlinetext数据库中取得试题数据；
    //并且根据试题的章节保存在不同的EXCEL文件中
    //
    //
    class Program
    {

        static void Main(string[] args)
        {
            
            //sql语句，决定了我们要的数据
            string sqlstr = "select s.SubjectName,t.TextBookName,c2.ChapterName,c.ChapterName as nodename,QuestionTitle,AnswerA,AnswerB,AnswerC,AnswerD,CorrectAnswer,Explain,q.Remark from Question as q left join PaperCodes as p on q.PaperCodeId=p.PaperCodeId 	left join Subject as s on p.SubjectId=s.SubjectId left join TextBook as t on q.TextBookId=t.TextBookId	left join Chapter as c on q.ChapterId=c.ChapterId	left join Chapter as c2 on c.ChapterParentNo=c2.ChapterId";
            //数据库连接字符串
            String connsql = "server=.;database=OnLineTest;integrated security=SSPI"; // 数据库连接字符串,database设置为自己的数据库名，以Windows身份验证
            SqlConnection connection = null;
            //保存文件的路径
            string path = @"c:\MyDir";
            //根据path生成保存文件的文件夹
            try
            {
                if (Directory.Exists(path))
                {
                    Console.WriteLine(path + " 文件夹已经存在，不需要创建。");
                }
                else
                {
                    Directory.CreateDirectory(path);
                    Console.WriteLine(path + " 文件夹创建成功。");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(path + "文件夹创建失败，错误：" + ex.Message);
            }
            //从数据库中获取试题数据
            try
            {
                using (connection = new SqlConnection(connsql))
                {
                    SqlCommand command = new SqlCommand(sqlstr, connection);
                    DataSet dataset = new DataSet();
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    int resultrows = adapter.Fill(dataset);
                    Console.WriteLine(resultrows);
                    Console.WriteLine("================= 下面开始处理数据=================");
                    createExcle(dataset, PartitionDataset(dataset), path);
                    Console.WriteLine("================= 数据处理完成=================");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                connection.Close();
                Console.ReadLine();
            }
        }

        /// <summary>
        /// 根据list内的数组将dataset内的数据分别写入到excel中
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="list"></param>
        public static void createExcle(DataSet ds, List<int[]> list, string path)
        {

            for (int i = 0; i < list.Count; i++)
            {
                IWorkbook workbook = new HSSFWorkbook();//创建Workbook对象  
                ISheet sheet = workbook.CreateSheet("Sheet1");//创建工作表 

                //下面开始添加首行标题行
                IRow headerRow = sheet.CreateRow(0);//在工作表中添加首行  
                string[] headerRowName = new string[] { "题型", "试题内容", "选项A", "选项B", "选项C", "选项D", "选项E", "选项F", "标准答案", "答案解析", "难易程序" };
                ICellStyle style = workbook.CreateCellStyle();
                style.Alignment = HorizontalAlignment.Center;//设置单元格的样式：水平对齐居中
                IFont font = workbook.CreateFont();//新建一个字体样式对象
                font.Boldweight = short.MaxValue;//设置字体加粗样式
                style.SetFont(font);//使用SetFont方法将字体样式添加到单元格样式中
                for (int j = 0; j < headerRowName.Length; j++)
                {
                    ICell cell = headerRow.CreateCell(j);
                    cell.SetCellValue(headerRowName[j]);
                    cell.CellStyle = style;
                }


                //下面开始添加数据行
                int starrow = list[i][0];//开始行的行号
                int endrow = list[i][1];//结束行的行号
                string[] ArrayDifficult = { "难", "较难", "中等", "较易", "容易" };
                string[] ArrayAnswer = { "A", "B", "C", "D" };
                Random random = new Random();
                for (int k = starrow; k <= endrow; k++)
                {
                    Console.WriteLine("==========开始处理第 " + k + " 行数据。============");
                    int rownumber = sheet.LastRowNum;
                    IRow datarow = sheet.CreateRow(rownumber + 1);
                    DataRow dr = ds.Tables[0].Rows[k];
                    string QuestionTitle = (!string.IsNullOrEmpty(dr["Remark"].ToString().Trim()) ? dr["Remark"].ToString().Trim() + "---" : string.Empty) + dr["QuestionTitle"].ToString().Trim();
                    string a = dr["AnswerA"].ToString().Trim();
                    string b = dr["AnswerB"].ToString().Trim();
                    string c = dr["AnswerC"].ToString().Trim();
                    string d = dr["AnswerD"].ToString().Trim();
                    int answer = Int32.Parse(dr["CorrectAnswer"].ToString().Trim());
                    string explain = dr["Explain"].ToString().Trim();
                    //判断题
                    if (string.IsNullOrEmpty(a + b + c + d))
                    {
                        datarow.CreateCell(0).SetCellValue("判断");
                        datarow.CreateCell(2).SetCellValue(answer == 1 ? "正确" : "错误");
                    }
                    else
                    {//单选题 
                        datarow.CreateCell(0).SetCellValue("单选");
                        datarow.CreateCell(2).SetCellValue(a);
                    }
                    datarow.CreateCell(1).SetCellValue(QuestionTitle);//试题内容
                    datarow.CreateCell(3).SetCellValue(b);//选项B
                    datarow.CreateCell(4).SetCellValue(c);//选项C
                    datarow.CreateCell(5).SetCellValue(d);//选项D
                    datarow.CreateCell(6).SetCellValue(string.Empty);//选项E
                    datarow.CreateCell(7).SetCellValue(string.Empty);//选项F
                    datarow.CreateCell(8).SetCellValue(ArrayAnswer[answer - 1]);//标准答案
                    datarow.CreateCell(9).SetCellValue(explain);//答案解析
                    datarow.CreateCell(10).SetCellValue(ArrayDifficult[random.Next(5)]);//难易程度
                }
                for (int k = 0; k < headerRow.Cells.Count; k++)
                {
                    sheet.AutoSizeColumn(k);
                }
                //下面处理文件名
                string filename = path + "\\";
                DataRow Fdr = ds.Tables[0].Rows[starrow];
                filename += Fdr["TextBookName"].ToString().Trim()+"_";
                filename += Fdr["ChapterName"].ToString().Trim() + "_";
                filename += Fdr["nodename"].ToString().Trim().Replace(":", "");
                filename += ".xls";

                //下面开始写入文件
                using (FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.ReadWrite))
                {
                    workbook.Write(fs);
                    fs.Flush();
                    fs.Close();
                }
                Console.WriteLine("第" + (i+1) + " 个文件，写入成功，文件名：" + filename);
                //Thread.Sleep(1000*5);
            }
        }




        //将从数据库中取得的数据按照章节进行划分并保存在list中
        public static List<int[]> PartitionDataset(DataSet dataset)
        {
            int totalrows = dataset.Tables[0].Rows.Count;
            List<Int32[]> list = new List<int[]>();
            int star = 0, current = 0, end = 0;
            for (int i = 1; i < totalrows; i++)
            {
                current = i;
                DataRow currentrow = dataset.Tables[0].Rows[current];
                DataRow lastrow = dataset.Tables[0].Rows[current - 1];
                string currentSTR = currentrow[0].ToString().Trim() + currentrow[1].ToString().Trim() + currentrow[2].ToString().Trim() + currentrow[3].ToString().Trim();
                string lastSTR = lastrow[0].ToString().Trim() + lastrow[1].ToString().Trim() + lastrow[2].ToString().Trim() + lastrow[3].ToString().Trim();
                if (!string.Equals(currentSTR, lastSTR) || current == totalrows - 1)
                {
                    end = current == totalrows - 1 ? current : current - 1;
                    list.Add(new int[] { star, end });
                    star = current;
                }
            }
            return list;
        }
    }
}
