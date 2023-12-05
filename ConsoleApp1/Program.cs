using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Pdf;
using Spire.Pdf.General.Find;
using System;
using System.Collections;
using System.Drawing;
using System.Text;
using static System.Net.Mime.MediaTypeNames;
using System.IO;
using System.Web.UI.WebControls;
using System.Runtime.Remoting.Messaging;
using Vehicle.Helper;

namespace ConsoleApp1
{
    internal class Program
    {
        string[] file_name = new String[] { };
        PointF[] p = new PointF[] { };
        PointF[] p_2 = new PointF[] { };
        PointF[] p_3 = new PointF[] { };
        PointF w = new PointF { };
        int[] index = new int[] { };
        static  string time="";
        static string CX1 = "insert into AF_SBZH_LV (PERSON,TASK,TIME,IDD) VALUES ('";
        static string CX2 = "','";
        static string CX3 = ")";
        static string CX4 = "AF_SBZH_LV_X.NEXTVAL";
        static string[] name = {"李卫卫","毛宇坤","左永琪","范姜辉","周泽宇" };
        SQLFind SQLFind = new SQLFind();
        static void Main(string[] args)
        {
            Program a = new Program();
            a.Search_name();
            Console.ReadLine();
        }
        void Search_name()
        {
            ArrayList list = new ArrayList(file_name);
            String path = Environment.CurrentDirectory;
            Console.WriteLine(path);
            int tmp = path.LastIndexOf(@"\");
            Console.WriteLine(tmp);
            time = path.Substring(tmp+1,path.Length-1-tmp);
            Console.WriteLine(time);
            //第一种方法
            var files = Directory.GetFiles(path, "*.pdf");
            foreach (var file in files)
            {
                list.Add(file);
                Console.WriteLine(file.ToString());
            }
            file_name = (string[])list.ToArray(typeof(string));
            Search_page();
        }
        void Search_page()
        {
            ArrayList list = new ArrayList(index);

            for (int i = 0; i < file_name.Length; i++)
            {
                PdfDocument doc = new PdfDocument();
                doc.LoadFromFile(file_name[i].ToString());
                PdfTextFind[] results = null;
                int j = 0;
                foreach (PdfPageBase page in doc.Pages)
                {
                    results = page.FindText("设计变更管理附表", TextFindParameter.WholeWord).Finds;
                    j++;
                    foreach (PdfTextFind text in results)
                    {

                        list.Add(j);
                        break;
                        // point.Add(text.Position);
                    }

                };
                if (list.Count < i + 1)
                {
                    list.Add(0);
                }
                Console.WriteLine(list[i]);
            }
            index = (int[])list.ToArray(typeof(int));

            for (int i = 0; i < file_name.Length; i++)
            {
                Console.WriteLine("第" + (i + 1) + "个PDF" + file_name[i] + "设计变更管理附表在" + index[i] + "页");
                string tmp = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                Console.WriteLine("a" + (tmp.Substring(1, 3)));
                switch (tmp.Substring(1, 3)) {
                    case "ECN":
                        Cut_Rectangle(i);
                        break;
                    case "DZB":
                        Cut_Rectangle_1(i);
                        break;
                        default:
                        string tmp_9 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                        PdfDocument doc_9 = new PdfDocument();
                        doc_9.LoadFromFile(file_name[i].ToString());
                        doc_9.SaveToFile(Environment.CurrentDirectory + "\\非DZB和ECN" + tmp_9);
                        break;
                }
               /* if (tmp.Substring(1, 3).Equals("ECN"))
                {
                    Cut_Rectangle(i);
                }
                if (tmp.Substring(1, 3).Equals("DZB"))
                {
                    Cut_Rectangle_1(i);
                }*/

            }
            Console.WriteLine("完成，任意键退出");
        }
        void Find_switch(string numberString,int i) {
            switch (numberString)
            {

                case "1":
                    if (!Directory.Exists(Environment.CurrentDirectory + "\\李卫卫"))
                    {
                        Directory.CreateDirectory((Environment.CurrentDirectory + "\\李卫卫"));
                    }
                    string tmp = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                    PdfDocument doc_1 = new PdfDocument();
                    doc_1.LoadFromFile(file_name[i].ToString());
                    doc_1.SaveToFile(Environment.CurrentDirectory + "\\李卫卫" + tmp);
                    string url1 = CX1 + name[0] + CX2 + tmp + CX2 + time + "'," + CX4 + CX3;
                    SQLFind.SQLSearch1(url1);
                    break;
                case "2":
                    if (!Directory.Exists(Environment.CurrentDirectory + "\\李卫卫"))
                    {
                        Directory.CreateDirectory((Environment.CurrentDirectory + "\\李卫卫"));
                    }
                    string tmp_2 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                    PdfDocument doc_2 = new PdfDocument();
                    doc_2.LoadFromFile(file_name[i].ToString());
                    doc_2.SaveToFile(Environment.CurrentDirectory + "\\李卫卫" + tmp_2);
                    string url2 = CX1 + name[0] + CX2 + tmp_2 + CX2 + time + "'," + CX4 + CX3;
                    SQLFind.SQLSearch1(url2);
                    break;
                case "3":
                    if (!Directory.Exists(Environment.CurrentDirectory + "\\李卫卫"))
                    {
                        Directory.CreateDirectory((Environment.CurrentDirectory + "\\李卫卫"));
                    }
                    string tmp_3 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                    PdfDocument doc_3 = new PdfDocument();
                    doc_3.LoadFromFile(file_name[i].ToString());
                    doc_3.SaveToFile(Environment.CurrentDirectory + "\\李卫卫" + tmp_3);
                    string url3 = CX1 + name[0] + CX2 + tmp_3 + CX2 + time + "'," + CX4 + CX3;
                    SQLFind.SQLSearch1(url3);
                    break;
                case "4":
                    if (!Directory.Exists(Environment.CurrentDirectory + "\\毛宇坤"))
                    {
                        Directory.CreateDirectory((Environment.CurrentDirectory + "\\毛宇坤"));
                    }
                    string tmp_4 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                    PdfDocument doc_4 = new PdfDocument();
                    doc_4.LoadFromFile(file_name[i].ToString());
                    doc_4.SaveToFile(Environment.CurrentDirectory + "\\毛宇坤" + tmp_4);
                    string url4 = CX1 + name[1] + CX2 + tmp_4 + CX2 + time + "'," + CX4 + CX3;
                    SQLFind.SQLSearch1(url4);
                    break;
                case "5":
                    if (!Directory.Exists(Environment.CurrentDirectory + "\\毛宇坤"))
                    {
                        Directory.CreateDirectory((Environment.CurrentDirectory + "\\毛宇坤"));
                    }
                    string tmp_5 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                    PdfDocument doc_5 = new PdfDocument();
                    doc_5.LoadFromFile(file_name[i].ToString());
                    doc_5.SaveToFile(Environment.CurrentDirectory + "\\毛宇坤" + tmp_5);
                    string url5 = CX1 + name[1] + CX2 + tmp_5 + CX2 + time + "'," + CX4 + CX3;
                    Console.WriteLine(url5);
                    SQLFind.SQLSearch1(url5);
                    break;
                case "6":
                    if (!Directory.Exists(Environment.CurrentDirectory + "\\左永琪"))
                    {
                        Directory.CreateDirectory((Environment.CurrentDirectory + "\\左永琪"));
                    }
                    string tmp_6 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                    PdfDocument doc_6 = new PdfDocument();
                    doc_6.LoadFromFile(file_name[i].ToString());
                    doc_6.SaveToFile(Environment.CurrentDirectory + "\\左永琪" + tmp_6);
                    string url6 = CX1 + name[2] + CX2 + tmp_6 + CX2 + time + "'," + CX4 + CX3;
                    SQLFind.SQLSearch1(url6);
                    break;
                case "7":
                    if (!Directory.Exists(Environment.CurrentDirectory + "\\范姜辉"))
                    {
                        Directory.CreateDirectory((Environment.CurrentDirectory + "\\范姜辉"));
                    }
                    string tmp_7 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                    PdfDocument doc_7 = new PdfDocument();
                    doc_7.LoadFromFile(file_name[i].ToString());
                    doc_7.SaveToFile(Environment.CurrentDirectory + "\\范姜辉" + tmp_7);
                    string url7 = CX1 + name[3] + CX2 + tmp_7 + CX2 + time + "'," + CX4 + CX3;
                    SQLFind.SQLSearch1(url7);
                    break;
                case "8":
                    if (!Directory.Exists(Environment.CurrentDirectory + "\\周泽宇"))
                    {
                        Directory.CreateDirectory((Environment.CurrentDirectory + "\\周泽宇"));
                    }
                    string tmp_8 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                    PdfDocument doc_8 = new PdfDocument();
                    doc_8.LoadFromFile(file_name[i].ToString());
                    doc_8.SaveToFile(Environment.CurrentDirectory + "\\周泽宇" + tmp_8);
                    string url8 = CX1 + name[4] + CX2 + tmp_8 + CX2 + time + "'," + CX4 + CX3;
                    SQLFind.SQLSearch1(url8);
                    break;
                default:
                    Console.WriteLine(file_name[i] + "识别错误");
                    string tmp_9 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                    PdfDocument doc_9 = new PdfDocument();
                    doc_9.LoadFromFile(file_name[i].ToString());
                    doc_9.SaveToFile(Environment.CurrentDirectory + "\\识别错误" + tmp_9);
                    break;
            }

        }
        void Cut_Rectangle(int i)
        {
            ArrayList point = new ArrayList(p);
            ArrayList point_2 = new ArrayList(p_2);
            ArrayList point_3 = new ArrayList(p_3);
            //读取每个PDF

            PdfDocument doc = new PdfDocument();
            doc.LoadFromFile(file_name[i].ToString());
            //读取对应的页
            if (index[i] > 0)
            {
                PdfPageBase page = doc.Pages[index[i] - 1];
                PdfTextFind[] results_1 = null;
                PdfTextFind[] results_2 = null;
                PdfTextFind[] results_3 = null;
                //查找总装所在位置
                results_1 = page.FindText("总装", TextFindParameter.WholeWord).Finds;
                if (results_1.Length - 1 > 0)
                {
                    //Console.WriteLine("dijige" +(i+1)+ (results_1.Length-1));
                    foreach (PdfTextFind text_1 in results_1)
                    {
                        point.Add(text_1.Position);
                        break;
                    }
                    //查找C-所在位置
                    results_2 = page.FindText("所属C总成", TextFindParameter.WholeWord).Finds;
                    foreach (PdfTextFind text_2 in results_2)
                    {
                        point_2.Add(text_2.Position);
                        break;
                    }
                    results_3 = page.FindText("本", TextFindParameter.WholeWord).Finds;
                    foreach (PdfTextFind text_3 in results_3)
                    {
                        point_3.Add(text_3.Position);
                        break;
                    }
                    p = (PointF[])point.ToArray(typeof(PointF));
                    p_2 = (PointF[])point_2.ToArray(typeof(PointF));
                    p_3 = (PointF[])point_3.ToArray(typeof(PointF));
                    //截取位置
                    //string text = page.ExtractText(new RectangleF(p_3[0].X + 5, p[0].Y - 10, (p_2[0].X - p_3[0].X) + 5, 30));
                    string text = page.ExtractText(new RectangleF(p_3[0].X + 5, p[0].Y - 10, 60, 80));
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine(text);
                    // File.WriteAllText("Extract" + (i + 1) + ".txt", sb.ToString());
                    #region 单数值提取
                    //---------Enter your code here-----------
                    try
                    {
                        string numberString = System.Text.RegularExpressions.Regex.Replace(sb.ToString(), @"[^0-9]+", "").Substring(0, 1);

                        Console.WriteLine("第" + (i + 1) + "个PDF" + file_name[i] + "识别到" + numberString + "");
                        #endregion
                        Find_switch(numberString,i);


                    }
                    catch (Exception ex)
                    {
                        string tmp_9 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                        PdfDocument doc_9 = new PdfDocument();
                        doc_9.LoadFromFile(file_name[i].ToString());
                        doc_9.SaveToFile(Environment.CurrentDirectory + "\\识别错误" + tmp_9);
                        return; 
                    }
                }
                else
                {
                    Console.WriteLine(file_name[i] + "未找到总装零件");
                    string tmp_9 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                    PdfDocument doc_9 = new PdfDocument();
                    doc_9.LoadFromFile(file_name[i].ToString());
                    doc_9.SaveToFile(Environment.CurrentDirectory + "\\未找到总装零件" + tmp_9);
                }
            }
            else
            {
                Console.WriteLine(file_name[i] + "未找到设计变更附件");
                string tmp_9 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                PdfDocument doc_9 = new PdfDocument();
                doc_9.LoadFromFile(file_name[i].ToString());
                doc_9.SaveToFile(Environment.CurrentDirectory + "\\未找到设计变更附件" + tmp_9);
            }

        }

        void Cut_Rectangle_1(int i)
        {
            ArrayList point = new ArrayList(p);
            ArrayList point_2 = new ArrayList(p_2);
            ArrayList point_3 = new ArrayList(p_3);
            //读取每个PDF

            PdfDocument doc = new PdfDocument();
            doc.LoadFromFile(file_name[i].ToString());
            //读取对应的页
            if (index[i] > 0)
            {
                PdfPageBase page = doc.Pages[index[i] - 1];
                PdfTextFind[] results_1 = null;
                PdfTextFind[] results_2 = null;
                PdfTextFind[] results_3 = null;
                //查找总装所在位置
                results_1 = page.FindText("所属总成", TextFindParameter.WholeWord).Finds;
                if (results_1.Length - 1 > 0)
                {
                    //Console.WriteLine("dijige" +(i+1)+ (results_1.Length-1));
                    foreach (PdfTextFind text_1 in results_1)
                    {
                        point.Add(text_1.Position);
                        break;
                    }
                    //查找C-所在位置

                    p = (PointF[])point.ToArray(typeof(PointF));
                    //截取位置
                    string text = page.ExtractText(new RectangleF(p[0].X - 5, p[0].Y + 10, 30, 30));
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine(text);
                    // File.WriteAllText("Extract" + (i + 1) + ".txt", sb.ToString());
                    #region 单数值提取
                    //---------Enter your code here-----------
                    try
                    {
                        string numberString = System.Text.RegularExpressions.Regex.Replace(sb.ToString(), @"[^0-9]+", "").Substring(0, 1);
                    Console.WriteLine("第" + (i + 1) + "个PDF" + file_name[i] + "识别到" + numberString + "");
                        #endregion
                        Find_switch(numberString, i);
                    }
                    catch (Exception ex)
                    {
                        string tmp_9 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                        PdfDocument doc_9 = new PdfDocument();
                        doc_9.LoadFromFile(file_name[i].ToString());
                        doc_9.SaveToFile(Environment.CurrentDirectory + "\\识别错误" + tmp_9);
                        return; }
                }
                else
                {
                    Console.WriteLine(file_name[i] + "未找到总装零件");
                    string tmp_9 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                    PdfDocument doc_9 = new PdfDocument();
                    doc_9.LoadFromFile(file_name[i].ToString());
                    doc_9.SaveToFile(Environment.CurrentDirectory + "\\未找到总装零件" + tmp_9);
                }
            }
            else
            {
                Console.WriteLine(file_name[i] + "未找到设计变更附件");
                string tmp_9 = file_name[i].ToString().Replace(Environment.CurrentDirectory, "");
                PdfDocument doc_9 = new PdfDocument();
                doc_9.LoadFromFile(file_name[i].ToString());
                doc_9.SaveToFile(Environment.CurrentDirectory + "\\未找到设计变更附件" + tmp_9);
            }

        }

    }
}




