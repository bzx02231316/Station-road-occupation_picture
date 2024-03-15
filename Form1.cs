

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
            this.AutoScroll = true;//�Զ�����
            AutoScrollMinSize = new Size(10000, 10000);//�����Զ���������С��С
            AutoScrollMargin = new Size(10, 10);//�����Զ������߾�Ĵ�С
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Console.ForegroundColor = ConsoleColor.Black;
            Console.BackgroundColor = ConsoleColor.White;
        }
        List<string[]> train_info = new List<string[]>();//�洢�����г��ɵ�ռ����Ϣ
        static int train_num = 164;//�������ͼ���г���
        static int lie_num = 7;//��ȡ��excel��������

        public void jgczff(string filePath)
        {
            // �����ļ�·����ȡWorkbook����
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
            ISheet sheet = wk.GetSheetAt(3);// ��ȡExcel���ĵ�һ��Sheet
            for (int i = 1; i < train_num; i++)
            {
                // Train train = new Train();
                IRow row = sheet.GetRow(i);// ��ȡ��i�е�����
                string[] row_info = new string[lie_num];//�洢ÿ����Ϣ����ʱ����
                for (int j = 0; j < lie_num; j++)
                {
                    row_info[j] = row.GetCell(j).ToString();
                }
                train_info.Add(row_info);
            }
        }
        int a = 2;//ʱ�������չʾ�ڴ����ϵ�ϵ��
        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            jgczff("D:\\��������\\������\\�ɵ�ռ��ͼƬ\\�������.xlsx");
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
            Dictionary<string, Pen> pencolor = new Dictionary<string, Pen> //�г�������pen��ɫƥ����ֵ�
            {
                {"��������",orangewpen}, {"��������������",blackpen},{"�����η���",bluepen}, {"���㷽��",redpen},{"�۰�����",purplepen}
            }; //������ɫ�ֵ�

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
            Dictionary<string, Brush> brushcolor = new Dictionary<string, Brush>//�г��������ˢ��ɫƥ����ֵ�
            {
                {"��������",orangebrush}, {"��������������",blackbrush},{"�����η���",bluebrush}, {"���㷽��",redbrush},{"�۰�����",Purplebrush}
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
            int hour_num = 20;//�г�����ͼ��ʾ��Сʱ����
            int gudao_num = 8;//��վ�ɵ���Ŀ
            int width = hour_num * 60 * a;
            int height = 900;
            int startpointX = 100;
            int startpointY = 100;//��ʼλ�������Ͻ�
            // g.DrawRectangle(greenpen, 120, 100, width, height);
            List<string> time = new List<string>();
            for (int i = 4; i < 25; i++)//��������ͼСʱ��ע���ַ���
            {
                time.Add(i.ToString() + ":00");
            }
            int time_number = 0;
            //������
            for (int i = startpointX; i <= startpointX + width; i += 10 * a)
            {
                if ((i - startpointX) / a % 60 != 30)
                {
                    if ((i - startpointX) / a % 60 == 0)//Сʱ��
                    {
                        greenpen.Width = 5;//Сʱ�߿�ȼӴ�
                        g.DrawString(time[time_number], font1, greenbrush, i, startpointY - 70, stringFormat);
                        g.DrawString(time[time_number], font1, greenbrush, i, height + 120, stringFormat);
                        time_number += 1;
                    }
                    greenpen.DashStyle = DashStyle.Solid;
                    g.DrawLine(greenpen, i, startpointY, i, startpointY + height);
                    greenpen.Width = 3;
                }
                else//��Сʱ�ߡ�����������
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
            Dictionary<string, int[]> gudao_loction = new Dictionary<string, int[]> { };//��վ���ɵ�λ���ֵ�
            linY = startpointY;//��վ�����ߵ���ʱ������
            for (int i = 0; i < gudao_num; i++)//���ƺ���
            {
                g.DrawLine(greenpen, startpointX, linY, startpointX + width, linY);//���ƺ���
                g.DrawString(station[station.Length - 1 - i], font3, greenbrush, startpointX - 50, linY + 60, stringFormat);//��ע�ɵ���
                int[] loction = new int[] { startpointX - 60, linY + 60 };
                gudao_loction.Add(station[station.Length - 1 - i], loction);//�ɵ���������
                linY += height / gudao_num;//ÿ���ۼӹɵ����
            }
            linY = height + startpointY;
            g.DrawLine(greenpen, startpointX, linY, startpointX + width, linY);
            //����ͼƬ
            for (int i = 0; i < train_info.Count; i++)
            {
                int startloc = Convert.ToInt32(train_info[i][3]);
                int endloc = Convert.ToInt32(train_info[i][4]);
                pencolor[train_info[i][1]].Width = 14;
                g.DrawLine(pencolor[train_info[i][1]], startpointX + (startloc - (24 - hour_num) * 60) * a, gudao_loction[train_info[i][0]][1],
                    startpointX + (startloc + (endloc - startloc) / 2 - (24 - hour_num) * 60) * a, gudao_loction[train_info[i][0]][1]);//���򳵴�
                g.DrawString(train_info[i][5], font2, brushcolor[train_info[i][1]], startpointX + (startloc - (24 - hour_num) * 60) * a + 10, gudao_loction[train_info[i][0]][1] - 30, stringFormat);//��ע���򳵴�
                pencolor[train_info[i][2]].Width = 14;
                g.DrawLine(pencolor[train_info[i][2]], startpointX + (startloc + (endloc - startloc) / 2 - (24 - hour_num) * 60) * a, gudao_loction[train_info[i][0]][1],
                   startpointX + (endloc - (24 - hour_num) * 60) * a, gudao_loction[train_info[i][0]][1]);//ȥ�򳵴�
                g.DrawString(train_info[i][6], font2, brushcolor[train_info[i][2]], startpointX + (startloc + (endloc - startloc) / 2 - (24 - hour_num) * 60) * a + 20,
                    gudao_loction[train_info[i][0]][1] + 15, stringFormat);//��עȥ�򳵴�


            }
        }
    }
}