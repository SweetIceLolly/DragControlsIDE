#include "Controls.h"

int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	/* 您可以在创建窗体前执行您的代码，但是直到
	窗体关闭前在该语句块之后的代码不会被执行 */

	if (!MainWindow.Create(hInstance))
	{
		MessageBox(0, "创建窗体失败！", "创建窗体失败", 48);
		UnregisterClass(MainWindow.ClassName, hInstance);
		return 0;
	}

	/* 直到窗体关闭后此处的代码才会被执行 */

	return 0;
}

MyIpAddress a;
//此过程负责创建所有的控件
void CreateAllControls()
{
	a.hMenu = 1;
	a.Visible = true;
	a.Enabled = true;
	a.Left = 10;
	a.Top = 10;
	a.Height = 50;
	a.Width = 200;
	a.Create();
}

/*===================================================================================*/
/* 进入VB代码模式！ */

/*
以下是窗体的各种事件
【请注意】窗体所有的事件不得删除，若需要删除，
则需要把Controls.h头文件里相关的代码一并删改。
*/

//窗体加载事件
void Form_Load()
{
	
}

//窗体获得焦点事件
void Form_Activate()
{
	
}

//窗体失去焦点事件
void Form_Deactivate()
{
	
}

/*
键盘按下按键事件

KeyCode: 按下的按键的Ascii码
Shift: 系统功能键的按键状态，可以是Controls.h中定义的 *_KEY 的常数值
IsLongPress: 是否长按当前按键
*/
void Form_KeyDown(int KeyCode, int Shift, bool IsLongPress)
{
	
}

/*
键盘松开按键事件

KeyCode: 松开的按键的Ascii码
Shift: 系统功能键的按键状态，可以是Controls.h中定义的 *_KEY 的常数值
*/
void Form_KeyUp(int KeyCode, int Shift)
{
	
}

/*
窗体鼠标移动事件

Button: 鼠标按键的状态，可以是Controls.h中定义的 *_BUTTON 的常数值
Shift: 系统功能键的按键状态，可以是Controls.h中定义的 *_KEY 的常数值
X, Y: 鼠标所在的坐标
*/
void Form_MouseMove(int Button, int Shift, int X, int Y)
{
	
}

//鼠标按下按键事件，参数请见Form_MouseMove()
void Form_MouseDown(int Button, int Shift, int X, int Y)
{
	
}

//鼠标按键松开事件，参数请见Form_MouseMove()
void Form_MouseUp(int Button, int Shift, int X, int Y)
{
	
}

//窗体点击事件
void Form_Click()
{
	
}

/*
窗体双击事件

Button: 双击事件被触发时鼠标按键的状态，可以是Controls.h中定义的 WHEEL_*_BUTTON 的常数值
Shift: 双击事件被触发时系统功能键的按下状态（不支持Alt键），可以是Controls.h中定义的WHEEL_*_KEY 的常数值
X, Y: 双击事件被触发时鼠标在窗体中的坐标
*/
void Form_DoubleClick(int Button, int Shift, int X, int Y)
{
	
}

//鼠标离开事件
void Form_MouseLeave()
{
	
}

//窗体重绘事件
void Form_Paint()
{

}

/*
窗体开始改变大小事件

SizeDirection: 可以是 WMSZ_* 的常数值，为窗体从不同位置更改大小的方向的数值
【注意】返回一个非零值意味着窗体将取消更改大小
*/
int Form_BeginResize(int SizeDirection)
{
	return 0;
}

/*
窗体改变大小事件

SizeMode: 可以是 SIZE_* 的常数值 
NewWidth, NewHeight: 新的窗体大小，为窗体更改大小的方式的数值
*/
void Form_Resize(int SizeMode, int NewWidth, int NewHeight)
{
	
}

//窗体改变大小结束
void Form_FinishResizing()
{
	
}

/*
窗体开始移动事件
【注意】返回一个非零值意味着窗体将取消移动
*/
int Form_BeginMove()
{
	return 0;
}

/*
窗体移动事件

X, Y: 窗体的新坐标
*/
void Form_Move(int X, int Y)
{
	
}

//窗体移动结束
void Form_FinishMoving()
{

}

/*
窗体将要关闭事件
【注意】返回一个非零值意味着窗体将取消关闭
*/
int Form_QueryUnload()
{
	return 0;
}

//窗体关闭事件
void Form_Unload()
{
	
}

/*
鼠标滚轮事件

WheelDelta: 滚轮增加/减少的数值（当该参数为正数时为向上滚动，为负数则为向下滚动）
Button: 滚轮事件被触发时鼠标按键的状态，可以是Controls.h中定义的 WHEEL_*_BUTTON 的常数值
Shift: 滚轮事件被触发时系统功能键的按下状态（不支持Alt键），可以是Controls.h中定义的WHEEL_*_KEY 的常数值
X, Y: 滚轮事件被触发时鼠标在窗体中的坐标
*/
void Form_MouseWheel(int WheelDelta, int Button, int Shift, int X, int Y)
{
	
}

/*窗体所有事件结束*/
/*===================================================================================*/
/*以下为所有控件的事件*/
