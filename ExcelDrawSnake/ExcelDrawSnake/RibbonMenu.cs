using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDrawSnake.Properties;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDrawSnake.DrawSnakeCS;
using ExcelDrawSnake.View;
using Microsoft.Office.Core;
using IRibbonControl = ExcelDna.Integration.CustomUI.IRibbonControl;
using IRibbonUI = ExcelDna.Integration.CustomUI.IRibbonUI;

namespace ExcelDrawSnake
{
    [ComVisible(true)]
    public class RibbonMenu : ExcelRibbon
    {
        /// <summary>
        /// Excel
        /// </summary>
        static readonly Excel.Application XlApp = (Excel.Application)ExcelDnaUtil.Application;
        /// <summary>
        /// 蛇
        /// </summary>
        private SnakeCoreControl _snake;
        /// <summary>
        /// 控制类
        /// </summary>
        private GameControl _gameControl;
        /// <summary>
        /// 游戏是否暂停
        /// </summary>
        private bool _presentBool;
        /// <summary> 是否处于主题跟随状态</summary>
        private bool _zhuTiGensuiBool;
        /// <summary> 蛇的初始移动速度 </summary>
        private int _snakeSpeed = 500;
        private readonly int sizeHeightX = 35; //活动范围的高就是行数
        private readonly int sizeWidthY = 35;  //活动范围的宽就是列数
        private readonly double rowHeightw = 11.25;  //小方块的高
        private readonly double columnWidth = 1.25;//小方块的宽
        

        /// <summary>蛇的颜色</summary>
        private Color _snakeColor = Color.Red;
        /// <summary>果实的颜色</summary>
        private  Color _randomPointColor = Color.Blue;

        private IRibbonUI _ribbonUi;

        private static Excel.Workbook _xlWorkbook ;
        private static Excel.Worksheet _xlWorksheet;
        public void RibbonMenu_Load(IRibbonUI control)
        {
            _ribbonUi = control;
        }

        #region 初始化游戏界面 设置游戏开始按钮是否可用

        //设置游戏开始按钮是否可用
        private bool _beginEnabled;
        public bool GetBeginEnabled(IRibbonControl control)
        {
            return _beginEnabled;
        }
        
        /// <summary>
        /// 初始化游戏界面
        /// </summary>
        /// <param name="control"></param>
        public void btnInit_Click(IRibbonControl control)
        {
            //int id = Thread.CurrentThread.ManagedThreadId;
            //MessageBox.Show(id.ToString());

            _xlWorkbook = XlApp.ActiveWorkbook;
            _xlWorksheet = _xlWorkbook.ActiveSheet;
            if (_snake != null)
            {
                _snake.SnakeTimer.Enabled = false;
                _snake = null;
            }
            ExcelSet.SetCellSize(_xlWorksheet, sizeHeightX, sizeWidthY, rowHeightw, columnWidth);
            
            _beginEnabled = true;               //激活开始btn
            _btnSnakeColorEnabled = true;       //激活蛇的颜色设置btn
            _btnRandomPointColorEnabled = true; //激活果实颜色设置btn
            _btnzhutigensui = true;             //激活主题跟随颜色设置btn
            _galThemeEnabled = true;            //激活主题设置gallery
            _ribbonUi.Invalidate(); //刷新显示
        }

        #endregion


        #region 游戏开始按钮

        //设置游戏控制按钮是否可用
        private bool _btnPresentEnabled;
        public bool GetbtnPresentEnabled(IRibbonControl control)
        {
            return _btnPresentEnabled;
        }
        
        private Bitmap _loginImage = Resources.logout;
        private string _loginLabel = "游戏开始";
        public Bitmap GetBeginImage(IRibbonControl control)
        {
            return _loginImage;
        }
        public string GetBeginLabel(IRibbonControl control)
        {
            return _loginLabel;
        }
        public void btnBegin_Click(IRibbonControl control)
        {
            if (_loginLabel == "游戏开始")
            {
                ExcelSet.SetCellSize(_xlWorksheet, sizeHeightX, sizeWidthY, rowHeightw, columnWidth);
                _snake = new SnakeCoreControl(sizeHeightX, sizeWidthY, _snakeSpeed);
                _snake.SnakePointListChange += Sna_SnakePointListChange;
                _snake.SnakeRandomPointChange += _snake_SnakeRandomPointChange;
                _snake.SnakeDie += _snake_SnakeDie;

                _presentBool = true;
                _snake.SnakeTimer.Enabled = _presentBool;

                ExcelDisplay.DisplayInit(_xlWorksheet, _snake.SnakePointList, _snakeColor);//画出初始化时的蛇
                ExcelDisplay.DislayRandomPoint(_xlWorksheet, _snake.SnakeRandomPoint,_randomPointColor);//画出果实

                _gameControl = GameControl.GetGameControl();
                _gameControl.KeyboardKeyChange += _gameControl_KeyboardKeyChange;
                
                _loginLabel = "游戏结束";
                _loginImage = Resources.login;
                _btnPresentEnabled = true;
            }
            else
            {
                try
                {
                    _gameControl.GameControlColsing();
                    _presentBool = false;
                    _snake.SnakeTimer.Enabled = _presentBool;
                    _snake = null;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
                _loginLabel = "游戏开始";
                _loginImage = Resources.logout;
                _btnPresentEnabled = false;
            }
            _ribbonUi.Invalidate(); //刷新显示
        }
        
        #endregion


        #region 游戏暂停按钮

        private Bitmap _presentImage = Resources.恢复数据;
        private string _presentLabel = "游戏暂停";
        public string GetPresentLabel(IRibbonControl control)
        {
            return _presentLabel;
        }
        public Bitmap GetPresentImage(IRibbonControl control)
        {
            return Resources.恢复数据;
        }
        public void btnPresent_Click(IRibbonControl control)
        {
            if (_presentLabel == "游戏暂停")
            {
                _presentLabel = "继续游戏";
                _presentBool = false;
            }
            else
            {
                _presentLabel = "游戏暂停";
                _presentBool = true;
            }
            _snake.SnakeTimer.Enabled = _presentBool;
            _ribbonUi.Invalidate();
        }

        #endregion


        #region 游戏难度

        public void btnDifficulty1_Click(IRibbonControl control)
        {
            _snakeSpeed = 500;
        }
        public void btnDifficulty2_Click(IRibbonControl control)
        {
            _snakeSpeed = 300;
        }
        public void btnDifficulty3_Click(IRibbonControl control)
        {
            _snakeSpeed = 100;
        }

        #endregion


        #region 主题And颜色设置

        #region 自定义蛇的颜色

        private bool _btnSnakeColorEnabled = false;
        private Bitmap _btnSnakeColorImage = new Bitmap(32, 32);
        public bool GetbtnSnakeColorEnabled(IRibbonControl control)
        {
            return _btnSnakeColorEnabled;
        }
        public Bitmap GetbtnSnakeColorImage(IRibbonControl control)
        {
            Graphics g = Graphics.FromImage(_btnSnakeColorImage);
            SolidBrush solidBrush = new SolidBrush(_snakeColor);
            g.FillRectangle(solidBrush, new Rectangle(0, 0, _btnSnakeColorImage.Width, _btnSnakeColorImage.Height));//这句实现填充矩形的功能
            return _btnSnakeColorImage;
        }
        /// <summary>
        /// 自定义蛇的颜色
        /// </summary>
        /// <param name="control"></param>
        public void btnSnakeColor_Click(IRibbonControl control)
        {
            Color col;
            ColorDialog colorDialog = new ColorDialog();

            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                col = colorDialog.Color;
                _snakeColor = col;
            }
            _ribbonUi.Invalidate();
        }

        #endregion


        #region 自定义果实的颜色

        private bool _btnRandomPointColorEnabled = false;
        private Bitmap _btnRandomPointColorImage = new Bitmap(32, 32);
        public bool GetbtnRandomPointColorEnabled(IRibbonControl control)
        {
            return _btnRandomPointColorEnabled;
        }
        public Bitmap GetbtnRandomPointColorImage(IRibbonControl control)
        {
            Graphics g = Graphics.FromImage(_btnRandomPointColorImage);
            SolidBrush solidBrush = new SolidBrush(_randomPointColor);
            g.FillRectangle(solidBrush, new Rectangle(0, 0, _btnRandomPointColorImage.Width, _btnRandomPointColorImage.Height));//这句实现填充矩形的功能
            return _btnRandomPointColorImage;
        }
        /// <summary>
        /// 自定义果实的颜色
        /// </summary>
        /// <param name="control"></param>
        public void btnRandomPointColor_Click(IRibbonControl control)
        {
            Color col;
            ColorDialog colorDialog = new ColorDialog();

            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                col = colorDialog.Color;
                _randomPointColor = col;
                ExcelDisplay.DislayRandomPoint(_xlWorksheet, _snake.SnakeRandomPoint, _randomPointColor);
            }
            _ribbonUi.Invalidate();
        }

        #endregion

        #region 主题跟随
        private bool _btnzhutigensui = false;
        public bool Getzhutigensui(IRibbonControl control)
        {
            return _btnzhutigensui;
        }
        public void btnThemeGensui_Click(IRibbonControl control)
        {
            if (_zhuTiGensuiBool)
            {
                _zhuTiGensuiBool = false;

                _btnSnakeColorEnabled = true; //激活蛇的颜色设置btn
                _btnRandomPointColorEnabled = true; //激活果实颜色设置btn
                _galThemeEnabled = true; //激活主题设置gallery
            }
            else
            {
                _zhuTiGensuiBool = true;

                _btnSnakeColorEnabled = false;
                _btnRandomPointColorEnabled = false;
                _galThemeEnabled = false;
            }
            _ribbonUi.Invalidate();
        }

        #endregion

        #region 主题

        private bool _galThemeEnabled = false;

        public bool GetThemeEnabled(IRibbonControl control)
        {
            return _galThemeEnabled;
        }

        public void btnZhuti1_Click(IRibbonControl control)
        {
            //InsertPicture("K8", _xlWorksheet, @"D:\Myproject\ExcelProject\ExcelDrawSnake\ExcelDrawSnake\aaa.png");
            Excel.Range rng = _xlWorksheet.Range["K8", "K8"];
            InsertPicture(rng, _xlWorksheet, @"D:\Myproject\ExcelProject\ExcelDrawSnake\ExcelDrawSnake\aaa.png");
        }

        public void btnZhuti2_Click(IRibbonControl control)
        {

        }

        #endregion


        #endregion


        #region 游戏成绩

        private string _scorest = "0";

        public string GetSorceLabel(IRibbonControl control)
        {
            return _scorest;
        }

        #endregion
        

        #region 游戏控制

        private void _gameControl_KeyboardKeyChange(object sender, EventArgs e)
        {
            switch (_gameControl.KeyboardKey)
            {
                case Keys.Up:
                    _snake.CurrentMoveDirection = MoveDirectionEnum.Up;
                    break;
                case Keys.Down:
                    _snake.CurrentMoveDirection = MoveDirectionEnum.Down;
                    break;
                case Keys.Left:
                    _snake.CurrentMoveDirection = MoveDirectionEnum.Left;
                    break;
                case Keys.Right:
                    _snake.CurrentMoveDirection = MoveDirectionEnum.Right;
                    break;
            }
        }

        #endregion


        #region 游戏显示

        private void Sna_SnakePointListChange(object sender, EventArgs e)
        {
            ExcelDisplay.Display(_xlWorksheet, _snake.SnakePointList, _snakeColor);
        }

        private void _snake_SnakeRandomPointChange(object sender, EventArgs e)
        {
            if(_zhuTiGensuiBool)
            {
                _snakeColor = _randomPointColor;
                _randomPointColor = GetRandomColor();
            }
            ExcelDisplay.DislayRandomPoint(_xlWorksheet, _snake.SnakeRandomPoint, _randomPointColor);
            _scorest = $"{_snake.SnakeLength}";
            _ribbonUi.Invalidate();
        }

        #endregion


        #region 游戏结束

        delegate void DelegateThreadFunction();
        private void _snake_SnakeDie(object sender, EventArgs e)
        {
            //int id = Thread.CurrentThread.ManagedThreadId;
            //MessageBox.Show(id.ToString());

            MessageBox.Show($"游戏结束:\n    成绩：{_snake.SnakeLength}");

            //GameOver gameOver = new GameOver();
            //gameOver.ShowDialog();
            _loginLabel = "游戏开始";
            _gameControl.GameControlColsing();
            _presentBool = false;
            _snake.SnakeTimer.Enabled = _presentBool;
            _snake = null;

            _ribbonUi.Invalidate();
        }
        
        #endregion





        /// <summary>
        /// 获取随机生成的ARGB颜色
        /// </summary>
        /// <returns></returns>
        private Color GetRandomColor()
        {
            Random ra = new Random(unchecked((int)DateTime.Now.Ticks));

            var a = ra.Next(0, 255);//A
            var r = ra.Next(0, 255);//R
            var g = ra.Next(0, 255);//G
            var b = ra.Next(0, 255);//B

            Color color = Color.FromArgb(a, r, g, b);
            return color;
        }
        
        #region 插入图片

        /// 将图片插入到指定的单元格位置，并设置图片的宽度和高度。
        /// 注意：图片必须是绝对物理路径
        /// </summary>
        /// <param name="RangeName">单元格名称，例如：B4</param>
        /// <param name="PicturePath">要插入图片的绝对路径。</param>
        public void InsertPicture(string RangeName, Excel._Worksheet sheet, string PicturePath)
        {
            Excel.Range rng = (Excel.Range) sheet.get_Range(RangeName, Type.Missing);
            rng.Select();
            float PicLeft, PicTop, PicWidth, PicHeight; //距离左边距离，顶部距离，图片宽度、高度
            PicTop = Convert.ToSingle(rng.Top);
            PicWidth = Convert.ToSingle(rng.MergeArea.Width);
            PicHeight = Convert.ToSingle(rng.Height);
            PicWidth = Convert.ToSingle(rng.Width);
            PicLeft = Convert.ToSingle(rng.Left); //+ (Convert.ToSingle(rng.MergeArea.Width) - PicWidth) / 2;
            try
            {
                Excel.Pictures pics = (Excel.Pictures) sheet.Pictures(Type.Missing);
                pics.Insert(PicturePath, Type.Missing);
                pics.Left = (double) rng.Left;
                pics.Top = (double) rng.Top;
                //pics.Width = (double)rng.Width;
                //pics.Height = (double)rng.Height;
                pics.Width = 72;
                pics.Height = 15;
            }
            catch
            {
            }
            sheet.Shapes.AddPicture(PicturePath, MsoTriState.msoFalse,
                MsoTriState.msoTrue, PicLeft, PicTop, PicWidth, PicHeight);
        }

        /*
         * 如果是要在某个区域插入，改区域没有命名的话，直接传入选中区域
         * Cell1 = SourceSheet.Cells[第几行, 第几列];
         * Cell2 = SourceSheet.Cells[Row , Column];
         * SourceRange = SourceSheet.get_Range(Cell1, Cell2);
         * 然后把上面的这句去掉 Excel.Range rng = (Excel.Range)sheet.get_Range(RangeName, Type.Missing);
         * 把rng换成SourceRange
         * */

        /// 将图片插入到指定的单元格位置，并设置图片的宽度和高度。
        /// 注意：图片必须是绝对物理路径
        /// </summary>
        /// <param name="rng">Excel单元格选中的区域</param>
        /// <param name="PicturePath">要插入图片的绝对路径。</param>
        public void InsertPicture(Excel.Range rng, Excel._Worksheet sheet, string PicturePath)
        {
            rng.Select();
            float PicLeft, PicTop, PicWidth, PicHeight;
            try
            {
                PicLeft = Convert.ToSingle(rng.Left);
                PicTop = Convert.ToSingle(rng.Top);
                PicWidth = Convert.ToSingle(rng.Width);
                PicHeight = Convert.ToSingle(rng.Height);

                //参数含义：
                //图片路径
                //是否链接到文件
                //图片插入时是否随文档一起保存
                //图片在文档中的坐标位置 坐标
                //图片显示的宽度和高度
                sheet.Shapes.AddPicture(PicturePath, MsoTriState.msoFalse,
                    MsoTriState.msoTrue, PicLeft, PicTop, 11, 11);
            }
            catch (Exception ex)
            {
                MessageBox.Show("错误：" + ex.Message);
            }
        }

        #endregion

    }
}
