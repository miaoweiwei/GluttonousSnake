using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelDrawSnake.DrawSnakeCS
{
    public class GameControl
    {
        /// <summary>
        /// 键盘Hook管理
        /// </summary>
        private readonly KeyboardHookLib _keyboardHook = null;
        
        #region 单例

        private static volatile GameControl _gameControl = null;
        private static readonly object LockHelper = new object();

        public static GameControl GetGameControl()
        {
            if (_gameControl == null)
            {
                lock (LockHelper)
                {
                    if (_gameControl == null)
                        _gameControl = new GameControl();
                }
            }
            return _gameControl;
        }

        #endregion


        private Keys _key;
        /// <summary>
        /// 键盘按键改变时（上下左右）
        /// </summary>
        public event EventHandler KeyboardKeyChange;
        /// <summary>
        /// 键盘按键上下左右
        /// </summary>
        public Keys KeyboardKey {
            get { return _key; }
            set
            {
                _key = value;
                if (KeyboardKeyChange != null)
                {
                    KeyboardKeyChange(this, new EventArgs());
                }
            }
        }

        private GameControl()
        {
            //安装勾子
            _keyboardHook = new KeyboardHookLib();
            _keyboardHook.InstallHook(this.OnKeyPress);
        }

        /// <summary>
        /// 关闭控制类
        /// </summary>
        public void GameControlColsing()
        {
            //取消勾子
            if (_keyboardHook != null) _keyboardHook.UninstallHook();
            _gameControl = null;
        }

        public void OnKeyPress(KeyboardHookLib.HookStruct hookStruct, out bool handle)
        {
            handle = false; //预设不拦截任何键Z
            Keys key = (Keys)hookStruct.vkCode;
            switch (key)
            {
                case Keys.Up:
                    KeyboardKey = Keys.Up;
                    break;
                case Keys.Down:
                    KeyboardKey = Keys.Down;
                    break;
                case Keys.Left:
                    KeyboardKey = Keys.Left;
                    break;
                case Keys.Right:
                    KeyboardKey = Keys.Right;
                    break;
            }
        }
    }
}
