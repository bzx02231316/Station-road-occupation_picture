

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Drawing.Drawing2D;

namespace Station_road_occupation_picture
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            this.AutoScroll = true;//自动滚动
            AutoScrollMinSize = new Size(10000, 10000);//设置自动滚动的最小大小
            AutoScrollMargin = new Size(10, 10);//设置自动滚动边距的大小
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Console.ForegroundColor = ConsoleColor.Black;
            Console.BackgroundColor = ConsoleColor.White;
        }
        List<string[]> train_info = new List<string[]>();//存储各次列车股道占用信息
        static int train_num = 164;//京广高铁图面列车数
        static int lie_num = 7;//读取的excel表格的列数

        public void jgczff(string filePath)
        {
            // 根据文件路径获取Workbook对象
            IWorkbook wk = null;
            string extension = System.IO.Path.GetExtension(filePath);
            FileStream fs = File.OpenRead(filePath);
            if (extension.Equals(".xls"))
            {
                wk = new HSSFWorkbook(fs);
            }
            else
            {
                wk = new XSSFWorkbook(fs);
            }
            fs.Close();
            ISheet sheet = wk.GetSheetAt(3);// 获取Excel表格的第一个Sheet
            for (int i = 1; i < train_num; i++)
            {
                // Train train = new Train();
                IRow row = sheet.GetRow(i);// 获取第i行的数据
                string[] row_info = new string[lie_num];//存储每行信息的临时数组
                for (int j = 0; j < lie_num; j++)
                {
                    row_info[j] = row.GetCell(j).ToString();
                }
                train_info.Add(row_info);
            }
        }
        int a = 2;//时间分钟数展示在窗体上的系数
        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            jgczff("D:\\我是桌面\\所工作\\股道占用图片\\矩阵输出.xlsx");
            Graphics g = e.Graphics;
            g.TranslateTransform(this.AutoScrollPosition.X, this.AutoScrollPosition.Y);
            Pen purplepen = new Pen(Color.Purple, 2);
            Pen bluepen = new Pen(Color.Blue, 2);
            Pen redpen = new Pen(Color.Red, 2);
            Pen blackpen = new Pen(Color.Black, 2);
            Color mycolor = Color.FromArgb(3, 220, 101);
            Pen greenpen = new Pen(mycolor, 2);
            Pen yellowpen = new Pen(Color.Yellow, 2);
            Pen orangewpen = new Pen(Color.Orange, 2);
            Pen backpen = new Pen(Color.Purple, 200);
            Pen[] pen = new Pen[5] { redpen, blackpen, yellowpen, purplepen, bluepen };
            Dictionary<string, Pen> pencolor = new Dictionary<string, Pen> //列车方向与pen颜色匹配的字典
            {
                {"北京方向",orangewpen}, {"北京西二场方向",blackpen},{"动车段方向",bluepen}, {"京广方向",redpen},{"雄安方向",purplepen}
            }; //方向颜色字典

            Brush blackbrush = new SolidBrush(Color.Black);
            Brush backbrush = new SolidBrush(this.BackColor);
            Brush greenbrush = new SolidBrush(mycolor);
            Brush graybrush = new SolidBrush(Color.Gray);
            Brush redbrush = new SolidBrush(Color.Red);
            Brush bluebrush = new SolidBrush(Color.Blue);
            Brush yellowbrush = new SolidBrush(Color.Yellow);
            Brush orangebrush = new SolidBrush(Color.Orange);
            Brush Purplebrush = new SolidBrush(Color.Purple);
            Brush[] brush = new Brush[5] { redbrush, blackbrush, yellowbrush, Purplebrush, bluebrush };
            Dictionary<string, Brush> brushcolor = new Dictionary<string, Brush>//列车方向与笔刷颜色匹配的字典
            {
                {"北京方向",orangebrush}, {"北京西二场方向",blackbrush},{"动车段方向",bluebrush}, {"京广方向",redbrush},{"雄安方向",Purplebrush}
            };
            Font font = new Font("Times New Roman", 18, FontStyle.Italic);
            Font font1 = new Font("Times New Roman", 12, FontStyle.Regular);
            Font font4 = new Font("Times New Roman", 8, FontStyle.Regular);
            Font font6 = new Font("Times New Roman", 10, FontStyle.Regular);
            Font font3 = new Font("Times New Roman", 12, FontStyle.Regular);
            Font font2 = new Font("Times New Roman", 6, FontStyle.Bold);
            Font font5 = new Font("Times New Roman", 6, FontStyle.Bold);
            Font font7 = new Font("Times New Roman", 5, FontStyle.Regular);
            HatchBrush myHatchBrush = new HatchBrush(HatchStyle.BackwardDiagonal, Color.Gray, Color.White);
            HatchBrush myHatchBrush1 = new HatchBrush(HatchStyle.Percent25, this.BackColor, this.BackColor);
            var stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            int hour_num = 20;//列车运行图显示的小时个数
            int gudao_num = 8;//车站股道数目
            int width = hour_num * 60 * a;
            int height = 900;
            int startpointX = 100;
            int startpointY = 100;//开始位置在左上角
            // g.DrawRectangle(greenpen, 120, 100, width, height);
            List<string> time = new List<string>();
            for (int i = 4; i < 25; i++)//用于运行图小时标注的字符串
            {
                time.Add(i.ToString() + ":00");
            }
            int time_number = 0;
            //画竖线
            for (int i = startpointX; i <= startpointX + width; i += 10 * a)
            {
                if ((i - startpointX) / a % 60 != 30)
                {
                    if ((i - startpointX) / a % 60 == 0)//小时线
                    {
                        greenpen.Width = 5;//小时线宽度加粗
                        g.DrawString(time[time_number], font1, greenbrush, i, startpointY - 70, stringFormat);
                        g.DrawString(time[time_number], font1, greenbrush, i, height + 120, stringFormat);
                        time_number += 1;
                    }
                    greenpen.DashStyle = DashStyle.Solid;
                    g.DrawLine(greenpen, i, startpointY, i, startpointY + height);
                    greenpen.Width = 3;
                }
                else//半小时线——虚线线型
                {
                    greenpen.DashPattern = new float[] { 12, 4 };
                    greenpen.DashStyle = DashStyle.Custom;
                    g.DrawLine(greenpen, i, startpointY, i, startpointY + height);
                }
            }
            greenpen.DashStyle = DashStyle.Solid;
            int linY = startpointY;
            int[] tt_time = new int[] { 21, 9, 10, 14, 19, 12, 12, 15, 16, 13, 24, 17, 18, 22, 26, 7, 20, 22, 19, 25, 27, 19 };//1548
            string[] station = new string[] { "20G", "19G", "18G", "17G", "16G", "15G", "14G", "13G" };
            Dictionary<string, int[]> gudao_loction = new Dictionary<string, int[]> { };//车站各股道位置字典
            linY = startpointY;//车站中心线的临时纵坐标
            for (int i = 0; i < gudao_num; i++)//绘制横线
            {
                g.DrawLine(greenpen, startpointX, linY, startpointX + width, linY);//绘制横线
                g.DrawString(station[station.Length - 1 - i], font3, greenbrush, startpointX - 50, linY + 60, stringFormat);//标注股道名
                int[] loction = new int[] { startpointX - 60, linY + 60 };
                gudao_loction.Add(station[station.Length - 1 - i], loction);//股道的纵坐标
                linY += height / gudao_num;//每次累加股道宽度
            }
            linY = height + startpointY;
            g.DrawLine(greenpen, startpointX, linY, startpointX + width, linY);
            //绘制图片
            for (int i = 0; i < train_info.Count; i++)
            {
                int startloc = Convert.ToInt32(train_info[i][3]);
                int endloc = Convert.ToInt32(train_info[i][4]);
                pencolor[train_info[i][1]].Width = 14;
                g.DrawLine(pencolor[train_info[i][1]], startpointX + (startloc - (24 - hour_num) * 60) * a, gudao_loction[train_info[i][0]][1],
                    startpointX + (startloc + (endloc - startloc) / 2 - (24 - hour_num) * 60) * a, gudao_loction[train_info[i][0]][1]);//来向车次
                g.DrawString(train_info[i][5], font2, brushcolor[train_info[i][1]], startpointX + (startloc - (24 - hour_num) * 60) * a + 10, gudao_loction[train_info[i][0]][1] - 30, stringFormat);//标注来向车次
                pencolor[train_info[i][2]].Width = 14;
                g.DrawLine(pencolor[train_info[i][2]], startpointX + (startloc + (endloc - startloc) / 2 - (24 - hour_num) * 60) * a, gudao_loction[train_info[i][0]][1],
                   startpointX + (endloc - (24 - hour_num) * 60) * a, gudao_loction[train_info[i][0]][1]);//去向车次
                g.DrawString(train_info[i][6], font2, brushcolor[train_info[i][2]], startpointX + (startloc + (endloc - startloc) / 2 - (24 - hour_num) * 60) * a + 20,
                    gudao_loction[train_info[i][0]][1] + 15, stringFormat);//标注去向车次


            }
        }
    }
}