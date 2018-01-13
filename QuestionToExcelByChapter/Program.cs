using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.IO;

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
            string sqlstr = "select s.SubjectName,t.TextBookName,c2.ChapterName,c.ChapterName,QuestionTitle,AnswerA,AnswerB,AnswerC,AnswerD,CorrectAnswer,Explain,q.Remark from Question as q left join PaperCodes as p on q.PaperCodeId=p.PaperCodeId 	left join Subject as s on p.SubjectId=s.SubjectId left join TextBook as t on q.TextBookId=t.TextBookId	left join Chapter as c on q.ChapterId=c.ChapterId	left join Chapter as c2 on c.ChapterParentNo=c2.ChapterId";
            //数据库连接字符串
            String connsql = "server=.;database=OnLineTest;integrated security=SSPI"; // 数据库连接字符串,database设置为自己的数据库名，以Windows身份验证
            SqlConnection connection = null;
            //保存文件的路径
            string path = @"D:\MyDir";
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
                    Console.ReadLine();
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                connection.Close();
            }
        }
        //将从数据库中取得的数据保存到根据不同章节生成的EXCEL文件中
        private static bool HandlerDataset(DataSet dataset, Directory directory, out DirectoryInfo directoryinfo)
        {
            bool result = true;
            int totalrows, currentrow = 0, starrow = 0, endrow = 0;
            string filename;
            totalrows = dataset.Tables[0].Rows.Count;
            for (int i = 0; i < totalrows; i++)
            {
                DataRow  dr= dataset.Tables[0].Rows[starrow];
                for (int j = 0; j < 4; j++) {
                    filename += dr[j].ToString();
                }
            }
            return result;
        }
    }
}
