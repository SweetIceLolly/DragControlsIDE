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

//此过程负责创建所有的控件
void CreateAllControls()
{
【CreateAllControlsCodeHere】
}

/*===================================================================================*/
/* 进入VB代码模式！ */

/*
以下是窗体的各种事件
【请注意】窗体所有的事件不得删除，若需要删除，
则需要把Controls.h头文件里相关的代码一并删改。
*/

【WindowCodeHere】
/*窗体所有事件结束*/
/*===================================================================================*/
/*以下为所有控件的事件*/
【AllControlsCodeHere】