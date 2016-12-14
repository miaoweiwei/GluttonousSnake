using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Timers;
using System.Windows.Forms;
using Timer = System.Timers.Timer;

namespace ExcelDrawSnake.DrawSnakeCS
{
    /// <summary>
    /// 贪吃蛇核心控制
    /// 活动区域
    /// </summary>
    public class SnakeCoreControl
    {
        /// <summary>
        /// 活动区域的高X
        /// </summary>
        private readonly int _heightX;
        /// <summary>
        /// 活动区域的宽Y
        /// </summary>
        private readonly int _widthY;

        private bool _snakeRunState=true;
        /// <summary>
        /// 蛇死亡时触发
        /// </summary>
        public event EventHandler SnakeDie;
        /// <summary>
        /// 蛇的运行状态 True表示正在运行 False表示死亡
        /// </summary>
        public bool SnakeRunState
        {
            get { return _snakeRunState; }
            set
            {
                _snakeRunState = value;
                if ((SnakeDie != null) && (!value))
                {
                    SnakeDie(this, new EventArgs());
                }
            }
        }

        private List<Point> _snakePointList;
        /// <summary>
        /// 蛇的坐标点的集合发生改变时触发
        /// </summary>
        public event EventHandler SnakePointListChange;
        /// <summary>
        /// 蛇的坐标点的集合
        /// </summary>
        public List<Point> SnakePointList
        {
            get { return _snakePointList; }
            set
            {
                _snakePointList = value;
                if (SnakePointListChange!=null)
                {
                    SnakePointListChange(this,new EventArgs());
                }
            }
        }

        /// <summary>
        /// 蛇身长度
        /// </summary>
        public int SnakeLength { get; private set; }
        /// <summary>
        /// 当前移动方向
        /// </summary>
        public MoveDirectionEnum CurrentMoveDirection { get; set; }
        /// <summary>
        /// 上一次移动的方向
        /// </summary>
        public MoveDirectionEnum LastMoveDirection { get; set; }

        /// <summary>
        /// 用于设置移动速度
        /// </summary>
        public Timer SnakeTimer { get; set; }
        /// <summary>
        /// 当前速度（时间间隔）
        /// </summary>
        public int CurrentSpeed { get; set; }

        public event EventHandler SnakeRandomPointChange;
        private Point _snakeRandomPoint;
        /// <summary>
        /// 随机生成点 果实
        /// </summary>
        public Point SnakeRandomPoint {
            get { return _snakeRandomPoint; }
            private set
            {
                _snakeRandomPoint = value;
                if (SnakeRandomPointChange!=null)
                {
                    SnakeRandomPointChange(this, new EventArgs());
                }
            }
        }

        /// <summary>
        /// 贪吃蛇核心控制初始化
        /// 活动区域
        /// 移动速度时间间隔
        /// 第一行的前三个点为蛇的初始化身体
        /// </summary>
        /// <param name="heightX">活动区域的高</param>
        /// <param name="widthY">活动区域的宽</param>
        /// <param name="snakeSpeed">初始化移动速度时间间隔(ms)</param>
        public SnakeCoreControl(int heightX, int widthY, int snakeSpeed)
        {
            if ((heightX < 4) || (widthY < 4) ||(snakeSpeed<=60))
            {
                return;
            }
            LastMoveDirection = MoveDirectionEnum.Right;
            CurrentMoveDirection = MoveDirectionEnum.Right;
            _heightX = heightX;
            _widthY = widthY;
            CurrentSpeed = snakeSpeed;
            SnakePointList = new List<Point>
            {
                new Point(0, 0),
                new Point(0, 1),
                new Point(0, 2)
            };
            SnakeRandomPoint = GetRandomPoint(0, heightX, widthY);
            SnakeLength = SnakePointList.Count;
            SnakeTimer = new Timer {Interval = CurrentSpeed};
            SnakeTimer.Elapsed += SnakeTimer_Elapsed;
        }
        
        /// <summary>
        /// 控制蛇的移动
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SnakeTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            SnakeTimer.Enabled = false;
            if (SnakePointList.Count> 3)
            {
                SnakePointList.RemoveAt(0);
            }

            var pointtemp = SnakePointList[SnakePointList.Count - 1];
            switch (CurrentMoveDirection)
            {
                case MoveDirectionEnum.Up:
                    if (LastMoveDirection == MoveDirectionEnum.Down)
                    {
                        //若原来是向下走 那就继续向下
                        pointtemp.X = pointtemp.X + 1;
                        CurrentMoveDirection = LastMoveDirection;
                    }
                    else
                    {
                        LastMoveDirection = CurrentMoveDirection;
                        pointtemp.X = pointtemp.X - 1;
                    }
                    break;
                case MoveDirectionEnum.Down:
                    if (LastMoveDirection == MoveDirectionEnum.Up)
                    {
                        //若原来是向上走 那就继续向上
                        pointtemp.X = pointtemp.X - 1;
                        CurrentMoveDirection = LastMoveDirection;
                    }
                    else
                    {
                        LastMoveDirection = CurrentMoveDirection;
                        pointtemp.X = pointtemp.X + 1;
                    }
                    break;
                case MoveDirectionEnum.Left:
                    if (LastMoveDirection == MoveDirectionEnum.Right)
                    {
                        //若原来是向右走 那就继续向右
                        pointtemp.Y = pointtemp.Y + 1;
                        CurrentMoveDirection = LastMoveDirection;
                    }
                    else
                    {
                        LastMoveDirection = CurrentMoveDirection;
                        pointtemp.Y = pointtemp.Y - 1;
                    }
                    break;
                case MoveDirectionEnum.Right:
                    if (LastMoveDirection == MoveDirectionEnum.Left)
                    {
                        //若原来是向左走 那就继续向左
                        pointtemp.Y = pointtemp.Y - 1;
                        CurrentMoveDirection = LastMoveDirection;
                    }
                    else
                    {
                        LastMoveDirection = CurrentMoveDirection;
                        pointtemp.Y = pointtemp.Y + 1;
                    }
                    break;
            }
            if ((pointtemp.X > _heightX-1) || (pointtemp.Y > _widthY-1)||(pointtemp.X<0)||(pointtemp.Y<0))
            {
                SnakeRunState = false;
                return;
            }
            List<Point> pointTempList = SnakePointList;

            if (pointtemp == SnakeRandomPoint)
            {
                pointTempList.Insert(0, new Point(0, 0));
                SnakeRandomPoint = GetRandomPoint(0, _heightX, _widthY);
                SnakeLength = _snakePointList.Count;

                if (CurrentSpeed > 60)
                {//移动加速
                    CurrentSpeed -= 20;
                    SnakeTimer.Interval = CurrentSpeed;
                }
            }

            pointTempList.Add(pointtemp);
            SnakePointList = pointTempList;
            
            SnakeTimer.Enabled = true;
        }

        /// <summary>
        /// 生成与蛇的坐标点的集合不重叠的随机点
        /// </summary>
        /// <param name="minValue">随机点坐标XY的最小值</param>
        /// <param name="xMaxValue">随机点坐标X的最大值</param>
        /// <param name="yMaxValue">随机点坐标XY的最大值</param>
        /// <returns></returns>
        public Point GetRandomPoint( int minValue, int xMaxValue,int yMaxValue)
        {
            Random ra = new Random(unchecked((int)DateTime.Now.Ticks));
            Point point = new Point();
            while (true)
            {
                point.X = ra.Next(minValue, xMaxValue);//随机取数
                point.Y = ra.Next(minValue, yMaxValue);//随机取数
                if (!SnakePointList.Any(snakePoint => (snakePoint.X == point.X) && (snakePoint.Y == point.Y)))
                {
                    break;
                }
            }
            return point;
        }
    }

    /// <summary>
    /// 移动方向
    /// </summary>
    public enum MoveDirectionEnum
    {
        /// <summary>
        /// 上
        /// </summary>
        Up = 1,
        /// <summary>
        /// 下
        /// </summary>
        Down = 2,
        /// <summary>
        /// 左
        /// </summary>
        Left = 3,
        /// <summary>
        /// 右
        /// </summary>
        Right = 4
    }
}
