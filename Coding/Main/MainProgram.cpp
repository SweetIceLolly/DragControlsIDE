#include "Controls.h"

int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	/* �������ڴ�������ǰִ�����Ĵ��룬����ֱ��
	����ر�ǰ�ڸ�����֮��Ĵ��벻�ᱻִ�� */

	if (!MainWindow.Create(hInstance))
	{
		MessageBox(0, "��������ʧ�ܣ�", "��������ʧ��", 48);
		UnregisterClass(MainWindow.ClassName, hInstance);
		return 0;
	}

	/* ֱ������رպ�˴��Ĵ���Żᱻִ�� */

	return 0;
}

//�˹��̸��𴴽����еĿؼ�
void CreateAllControls()
{
��CreateAllControlsCodeHere��
}

/*===================================================================================*/
/* ����VB����ģʽ�� */

/*
�����Ǵ���ĸ����¼�
����ע�⡿�������е��¼�����ɾ��������Ҫɾ����
����Ҫ��Controls.hͷ�ļ�����صĴ���һ��ɾ�ġ�
*/

��WindowCodeHere��
/*���������¼�����*/
/*===================================================================================*/
/*����Ϊ���пؼ����¼�*/
��AllControlsCodeHere��