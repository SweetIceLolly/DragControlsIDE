#include <windows.h>
#include <stdlib.h>

#define Me								MainWindow
#define GCLP_HBRBACKGROUND				(-10)
#define CLEARTYPE_QUALITY				5
#define SetClassLongPtr					SetClassLongA
#define GET_KEYSTATE_WPARAM(wParam)     (LOWORD(wParam))
#define GET_X_LPARAM(lp)                ((int)(short)LOWORD(lp))
#define GET_Y_LPARAM(lp)                ((int)(short)HIWORD(lp))
#define MAKEIPRANGE(low, high)			((LPARAM)(WORD)(((BYTE)(high) << 8) + (BYTE)(low)))
#define MAKEIPADDRESS(b1,b2,b3,b4)		((LPARAM)(((DWORD)(b1)<<24)+((DWORD)(b2)<<16)+((DWORD)(b3)<<8)+((DWORD)(b4))))

#define SS_EDITCONTROL					0x00002000L
#define CB_SETMINVISIBLE				0x1700 + 1
#define CB_GETMINVISIBLE				0x1700 + 2
#define UDS_HORZ						0x0040
#define UDM_SETRANGE32					(WM_USER + 111)
#define UDM_GETRANGE32					(WM_USER + 112)
#define UDM_SETACCEL					(WM_USER + 107)
#define UDM_GETACCEL					(WM_USER + 108)
#define UDM_SETPOS32					(WM_USER + 113)
#define UDM_GETPOS32					(WM_USER + 114)
#define PBS_SMOOTH              		0x01
#define PBS_VERTICAL            		0x04
#define PBS_MARQUEE             		0x08
#define PBM_GETRANGE            		(WM_USER + 7)
#define PBM_SETRANGE32          		(WM_USER + 6)
#define PBM_GETBARCOLOR         		(WM_USER + 15)
#define PBM_SETBARCOLOR         		(WM_USER + 9)
#define PBM_GETBKCOLOR          		(WM_USER + 14)
#define PBM_SETBKCOLOR					0x2000 + 1
#define PBM_DELTAPOS            		(WM_USER + 3)
#define PBM_GETPOS              		(WM_USER + 8)
#define PBM_SETPOS              		(WM_USER + 2)
#define TBS_AUTOTICKS					0x0001
#define TBS_VERT						0x0002
#define TBS_DOWNISLEFT					0x0400
#define TBS_NOTHUMB						0x0080
#define TBS_TOOLTIPS					0x0100
#define TBS_LEFT						0x0004
#define TBS_RIGHT						0x0000
#define TBS_BOTH						0x0008
#define TBS_NOTICKS						0x0010
#define TBM_SETTIPSIDE					(WM_USER + 31)
#define TBTS_TOP						0
#define TBTS_LEFT						1
#define TBTS_BOTTOM						2
#define TBTS_RIGHT						3
#define TBM_SETTICFREQ					(WM_USER + 20)
#define TBM_SETRANGEMAX					(WM_USER + 8)
#define TBM_SETRANGEMIN					(WM_USER + 7)
#define TBM_SETRANGE					(WM_USER + 6)
#define TBM_GETRANGEMAX					(WM_USER + 2)
#define TBM_GETRANGEMIN					(WM_USER + 1)
#define TBM_SETLINESIZE					(WM_USER + 23)
#define TBM_SETPAGESIZE					(WM_USER + 21)
#define TBM_SETPOS						(WM_USER + 5)
#define TBM_GETPOS						(WM_USER)
#define HKM_SETHOTKEY					(WM_USER + 1)
#define HKM_GETHOTKEY					(WM_USER + 2)
#define LVN_FIRST						(0U - 100U)
#define NM_FIRST						(0U-  0U)
#define LVN_BEGINLABELEDIT				(LVN_FIRST - 5)
#define LVN_ENDLABELEDIT				(LVN_FIRST - 6)
#define LVN_COLUMNCLICK					(LVN_FIRST - 8)
#define LVN_ITEMCHANGED					(LVN_FIRST - 1)
#define NM_SETFOCUS						(NM_FIRST - 7)
#define NM_KILLFOCUS					(NM_FIRST - 8)
#define LVCFMT_LEFT						0x0000
#define LVS_LIST						0x0003
#define LVS_ICON						0x0000
#define LVS_REPORT						0x0001
#define LVS_SMALLICON					0x0002
#define LVS_SORTASCENDING				0x0010
#define LVS_SORTDESCENDING				0x0020
#define LVS_ALIGNLEFT					0x0800
#define LVS_AUTOARRANGE					0x0100
#define LVS_EDITLABELS					0x0200
#define LVIS_SELECTED					0x0002
#define LVS_SINGLESEL					0x0004
#define LVCF_WIDTH						0x0002
#define LVCF_TEXT						0x0004
#define LVCF_FMT						0x0001
#define LVM_FIRST						0x1000
#define LVM_INSERTCOLUMN				(LVM_FIRST + 27)
#define LVIF_TEXT						0x00000001
#define LVM_INSERTITEM					(LVM_FIRST + 7)
#define LVM_SETITEMTEXT					(LVM_FIRST + 46)
#define LVM_GETITEM						(LVM_FIRST + 5)
#define LVM_GETITEMCOUNT				(LVM_FIRST + 4)
#define LVM_GETCOLUMN					(LVM_FIRST + 25)
#define LVM_SETCOLUMN					(LVM_FIRST + 26)
#define LVM_SETCOLUMNWIDTH				(LVM_FIRST + 30)
#define LVM_EDITLABEL					(LVM_FIRST + 23)
#define LVM_CANCELEDITLABEL				(LVM_FIRST + 179)
#define LVM_DELETEALLITEMS				(LVM_FIRST + 9)
#define LVM_DELETECOLUMN				(LVM_FIRST + 28)
#define LVM_DELETEITEM					(LVM_FIRST + 8)
#define LVM_ENSUREVISIBLE				(LVM_FIRST + 19)
#define LVM_FINDITEM 					(LVM_FIRST + 13)
#define LVM_SETBKCOLOR					(LVM_FIRST + 1)
#define LVM_GETBKCOLOR					(LVM_FIRST + 0)
#define LVM_SETTEXTBKCOLOR				(LVM_FIRST + 38)
#define LVM_GETTEXTBKCOLOR				(LVM_FIRST + 37)
#define LVM_GETTEXTCOLOR				(LVM_FIRST + 35)
#define LVM_SETTEXTCOLOR				(LVM_FIRST + 36)
#define LVM_SCROLL						(LVM_FIRST + 20)
#define LVM_SETITEMPOSITION32			(LVM_FIRST + 49)
#define LVM_GETITEMPOSITION				(LVM_FIRST + 16)
#define LVM_GETSELECTEDCOLUMN			(LVM_FIRST + 174)
#define LVM_GETTOPINDEX					(LVM_FIRST + 39)
#define LVM_SETVIEW						(LVM_FIRST + 142)
#define LVM_GETNEXTITEM					(LVM_FIRST + 12)
#define LV_VIEW_ICON					0x0000
#define LV_VIEW_DETAILS					0x0001
#define LV_VIEW_SMALLICON				0x0002
#define LV_VIEW_LIST					0x0003
#define LVFI_PARTIAL					0x0008
#define LVFI_STRING						0x0002
#define LVFI_NEARESTXY					0x0040
#define LVNI_SELECTED					0x0002
#define NM_CLICK						(NM_FIRST - 2)
#define NM_RCLICK						(NM_FIRST - 5)
#define TVN_FIRST						(0U - 400U)
#define TVN_BEGINLABELEDIT				(TVN_FIRST - 10)
#define TVN_ENDLABELEDIT				(TVN_FIRST - 11)
#define TVN_ITEMEXPANDING				(TVN_FIRST - 5)
#define TVN_SELCHANGING					(TVN_FIRST-1)
#define TVN_SELCHANGED					(TVN_FIRST-2)
#define TVS_HASBUTTONS					0x0001
#define TVS_HASLINES					0x0002
#define TVS_LINESATROOT					0x0004
#define TVS_EDITLABELS					0x0008
#define TVS_NOSCROLL					0x2000
#define TVS_NOHSCROLL					0x8000
#define TVS_SHOWSELALWAYS				0x0020
#define TVS_TRACKSELECT					0x0200
#define TVS_CHECKBOXES					0x0100
#define TVI_LAST						((HTREEITEM)(ULONG_PTR)-0x0FFFE)
#define TVIF_TEXT						0x0001
#define TV_FIRST						0x1100
#define TVM_INSERTITEM					(TV_FIRST + 0)
#define TVM_DELETEITEM					(TV_FIRST + 1)
#define TVM_ENSUREVISIBLE				(TV_FIRST + 20)
#define TVM_EXPAND						(TV_FIRST + 2)
#define TVM_EDITLABEL					(TV_FIRST + 14)
#define TVM_ENDEDITLABELNOW				(TV_FIRST + 22)
#define TVM_SETITEMHEIGHT				(TV_FIRST + 27)
#define TVM_GETITEMHEIGHT				(TV_FIRST + 28)
#define TVM_SETBKCOLOR					(TV_FIRST + 29)
#define TVM_SETTEXTCOLOR				(TV_FIRST + 30)
#define TVM_GETBKCOLOR					(TV_FIRST + 31)
#define TVM_GETTEXTCOLOR				(TV_FIRST + 32)
#define TVM_SETLINECOLOR				(TV_FIRST + 40)
#define TVM_GETLINECOLOR				(TV_FIRST + 41)
#define TVM_GETCOUNT					(TV_FIRST + 5)
#define TVM_GETINDENT					(TV_FIRST + 6)
#define TVM_SETINDENT					(TV_FIRST + 7)
#define TVM_GETITEM						(TV_FIRST + 12)
#define TVM_SETITEM						(TV_FIRST + 13)
#define TVM_GETNEXTITEM					(TV_FIRST + 10)
#define TVM_SELECTITEM					(TV_FIRST + 11)
#define TVGN_PARENT						0x0003
#define TVGN_CARET						0x0009
#define TVGN_LASTVISIBLE				0x000A
#define TVGN_NEXTVISIBLE				0x0006
#define TVM_GETVISIBLECOUNT				(TV_FIRST + 16)
#define TCS_BOTTOM						0x0002
#define TCS_FLATBUTTONS					0x0008
#define TCS_FIXEDWIDTH					0x0400
#define TCS_FOCUSONBUTTONDOWN			0x1000
#define TCS_FORCELABELLEFT				0x0020
#define TCS_HOTTRACK					0x0040
#define TCS_MULTILINE					0x0200
#define TCS_SCROLLOPPOSITE				0x0001
#define TCS_VERTICAL					0x0080
#define TCM_FIRST						0x1300
#define TCM_DELETEALLITEMS				(TCM_FIRST + 9)
#define TCM_DELETEITEM					(TCM_FIRST + 8)
#define TCM_DESELECTALL					(TCM_FIRST + 50)
#define TCM_GETCURSEL					(TCM_FIRST + 11)
#define TCIF_TEXT						0x0001
#define TCM_GETITEM						(TCM_FIRST + 5)
#define TCM_GETITEMCOUNT				(TCM_FIRST + 4)
#define TCM_GETROWCOUNT					(TCM_FIRST + 44)
#define TCM_HIGHLIGHTITEM				(TCM_FIRST + 51)
#define TCHT_ONITEMICON					0x0002
#define TCHT_ONITEMLABEL				0x0004
#define TCHT_ONITEM						(TCHT_ONITEMICON | TCHT_ONITEMLABEL)
#define TCM_HITTEST						(TCM_FIRST + 13)
#define TCM_INSERTITEM					(TCM_FIRST + 7)
#define TCM_SETCURSEL					(TCM_FIRST + 12)
#define TCM_SETITEM						(TCM_FIRST + 6)
#define TCM_SETITEMSIZE					(TCM_FIRST + 41)
#define TCM_SETCURFOCUS					(TCM_FIRST + 48)
#define TCM_SETMINTABWIDTH				(TCM_FIRST + 49)
#define TCN_FIRST						(0U - 550U)
#define TCN_SELCHANGE					(TCN_FIRST - 1)
#define TCN_SELCHANGING					(TCN_FIRST - 2)
#define ACS_AUTOPLAY					0x0004
#define ACS_CENTER						0x0001
#define ACS_TRANSPARENT					0x0002
#define ACM_PLAY						(WM_USER + 101)
#define ACM_STOP						(WM_USER + 102)
#define ACM_ISPLAYING					(WM_USER + 104)
#define ACM_OPEN 						(WM_USER + 100)
#define ACN_START						1
#define ACN_STOP						2
#define ES_SUNKEN						0x00004000
#define ES_DISABLENOSCROLL				0x00002000
#define ES_NOIME						0x00080000
#define ES_SELECTIONBAR 				0x01000000
#define EM_AUTOURLDETECT				(WM_USER + 91)
#define	AURL_ENABLEURL					1
#define EM_CANPASTE 					(WM_USER + 50)
#define EM_REDO 						(WM_USER + 84)
#define EM_CANREDO						(WM_USER + 85)
#define EM_EXGETSEL 					(WM_USER + 52)
#define EM_EXSETSEL 					(WM_USER + 55)
#define EM_EXLIMITTEXT					(WM_USER + 53)
#define EM_FINDTEXTEX					(WM_USER + 79)
#define CFM_BOLD						0x00000001
#define CFM_ITALIC						0x00000002
#define CFM_UNDERLINE					0x00000004
#define CFM_STRIKEOUT					0x00000008
#define CFE_PROTECTED					0x00000010
#define CFM_LINK						0x00000020
#define CFM_COLOR						0x40000000
#define CFM_SIZE						0x80000000
#define CFM_FACE						0x20000000
#define CFM_OFFSET						0x10000000
#define CFM_CHARSET 					0x08000000
#define CFM_EFFECTS						(CFM_BOLD | CFM_ITALIC | CFM_UNDERLINE | CFM_COLOR | \
											CFM_STRIKEOUT | CFE_PROTECTED | CFM_LINK)
#define CFM_ALL							(CFM_EFFECTS | CFM_SIZE | CFM_FACE | CFM_OFFSET | CFM_CHARSET)
#define EM_GETCHARFORMAT				(WM_USER + 58)
#define SCF_SELECTION					0x0001
#define EM_SETCHARFORMAT				(WM_USER + 68)
#define DTN_FIRST2						(0U-753U)
#define DTN_DATETIMECHANGE				(DTN_FIRST2 - 6)
#define DTS_LONGDATEFORMAT				0x0004
#define DTS_RIGHTALIGN					0x0020
#define DTS_SHOWNONE					0x0002
#define DTS_TIMEFORMAT					0x0009
#define DTS_UPDOWN						0x0001
#define DTM_FIRST						0x1000
#define DTM_GETRANGE					(DTM_FIRST + 3)
#define GDTR_MIN						0x0001
#define GDTR_MAX						0x0002
#define DTM_SETRANGE					(DTM_FIRST + 4)
#define DTM_GETSYSTEMTIME				(DTM_FIRST + 1)
#define DTM_SETSYSTEMTIME				(DTM_FIRST + 2)
#define GDT_VALID						0
#define GDT_NONE						1
#define DTM_SETFORMAT 					(DTM_FIRST + 5)
#define MCS_MULTISELECT					0x0002
#define MCS_WEEKNUMBERS					0x0004
#define MCS_NOTODAYCIRCLE				0x0008
#define MCS_NOTODAY						0x0010
#define MCM_FIRST						0x1000
#define MCM_GETCURSEL					(MCM_FIRST + 1)
#define MCM_SETCURSEL					(MCM_FIRST + 2)
#define MCM_SETFIRSTDAYOFWEEK			(MCM_FIRST + 15)
#define MCM_GETFIRSTDAYOFWEEK			(MCM_FIRST + 16)
#define MCM_GETRANGE					(MCM_FIRST + 17)
#define MCM_SETRANGE					(MCM_FIRST + 18)
#define MCM_GETMAXSELCOUNT				(MCM_FIRST + 3)
#define MCM_SETMAXSELCOUNT				(MCM_FIRST + 4)
#define MCM_GETSELRANGE					(MCM_FIRST + 5)
#define MCM_SETSELRANGE					(MCM_FIRST + 6)
#define MCM_SETTODAY					(MCM_FIRST + 12)
#define MCM_GETTODAY					(MCM_FIRST + 13)
#define MCN_FIRST						(0U - 746U)
#define MCN_SELCHANGE					(MCN_FIRST - 3)
#define IPM_CLEARADDRESS				(WM_USER + 100)
#define IPM_SETADDRESS					(WM_USER + 101)
#define IPM_SETRANGE					(WM_USER + 103)
#define IPM_SETFOCUS					(WM_USER + 104)
#define IPM_ISBLANK						(WM_USER + 105)
#define IPN_FIELDCHANGED				(0U - 860U)

struct _TREEITEM;
typedef struct _TREEITEM *HTREEITEM;

typedef struct _UDACCEL {
	UINT nSec;
	UINT nInc;
} UDACCEL, *LPUDACCEL;

typedef struct
{
   int iLow;
   int iHigh;
} PBRANGE, *PPBRANGE;

typedef struct tagNMLISTVIEW
{
	NMHDR   hdr;
	int     iItem;
	int     iSubItem;
	UINT    uNewState;
	UINT    uOldState;
	UINT    uChanged;
	POINT   ptAction;
	LPARAM  lParam;
} NMLISTVIEW, *LPNMLISTVIEW;

typedef struct tagLVITEMA
{
	UINT mask;
	int iItem;
	int iSubItem;
	UINT state;
	UINT stateMask;
	LPSTR pszText;
	int cchTextMax;
	int iImage;
	LPARAM lParam;
	int iIndent;
} LVITEM, *LPLVITEM;

typedef struct tagLVDISPINFO {
	NMHDR hdr;
	LVITEM item;
} NMLVDISPINFO, *LPNMLVDISPINFO;

typedef struct tagLVCOLUMNA
{
	UINT mask;
	int fmt;
	int cx;
	LPSTR pszText;
	int cchTextMax;
	int iSubItem;
	int iImage;
	int iOrder;
} LVCOLUMN, *LPLVCOLUMN;

typedef struct tagLVFINDINFOA
{
	UINT flags;
	LPCSTR psz;
	LPARAM lParam;
	POINT pt;
	UINT vkDirection;
} LVFINDINFO, *LPFINDINFO;

typedef struct tagTVITEMA {
	UINT      mask;
	HTREEITEM hItem;
	UINT      state;
	UINT      stateMask;
	LPSTR     pszText;
	int       cchTextMax;
	int       iImage;
	int       iSelectedImage;
	int       cChildren;
	LPARAM    lParam;
} TVITEM, *LPTVITEM;

typedef struct tagTVDISPINFOA {
	NMHDR hdr;
	TVITEM item;
} NMTVDISPINFO, *LPNMTVDISPINFO;

typedef struct tagNMTREEVIEWA {
	NMHDR       hdr;
	UINT        action;
	TVITEM    itemOld;
	TVITEM    itemNew;
	POINT       ptDrag;
} NMTREEVIEW, *LPNMTREEVIEW;

typedef struct tagTVITEMEXA {
	UINT      mask;
	HTREEITEM hItem;
	UINT      state;
	UINT      stateMask;
	LPSTR     pszText;
	int       cchTextMax;
	int       iImage;
	int       iSelectedImage;
	int       cChildren;
	LPARAM    lParam;
	int       iIntegral;
} TVITEMEXA, *LPTVITEMEXA;

typedef struct tagTVINSERTSTRUCTA {
	HTREEITEM hParent;
	HTREEITEM hInsertAfter;
	union
	{
		TVITEMEXA itemex;
		TVITEM  item;
	} DUMMYUNIONNAME;
} TVINSERTSTRUCT, *LPTVINSERTSTRUCT;

typedef struct tagTCITEMA
{
	UINT mask;
	DWORD dwState;
	DWORD dwStateMask;
	LPSTR pszText;
	int cchTextMax;
	int iImage;

	LPARAM lParam;
} TCITEM, *LPTCITEM;

typedef struct tagTCHITTESTINFO
{
	POINT pt;
	UINT flags;
} TCHITTESTINFO, *LPTCHITTESTINFO;

typedef struct _charformat
{
	UINT		cbSize;
	DWORD		dwMask;
	DWORD		dwEffects;
	LONG		yHeight;
	LONG		yOffset;
	COLORREF	crTextColor;
	BYTE		bCharSet;
	BYTE		bPitchAndFamily;
	char		szFaceName[LF_FACESIZE];
} CHARFORMAT;

typedef struct _charrange
{
	LONG	cpMin;
	LONG	cpMax;
} CHARRANGE;

typedef struct _findtextexa
{
	CHARRANGE chrg;
	LPCSTR	  lpstrText;
	CHARRANGE chrgText;
} FINDTEXTEX;

typedef struct tagNMDATETIMECHANGE
{
	NMHDR       nmhdr;
	DWORD       dwFlags;
	SYSTEMTIME  st;
} NMDATETIMECHANGE, *LPNMDATETIMECHANGE;

typedef struct tagNMSELCHANGE
{
	NMHDR           nmhdr;
	SYSTEMTIME      stSelStart;
	SYSTEMTIME      stSelEnd;
} NMSELCHANGE, *LPNMSELCHANGE;

typedef LONG(NTAPI *NtSuspendProcess)(IN HANDLE ProcessHandle);		//挂起进程API

const UINT MY_DEBUGGER_BREAKPOINT = 0x8888;		//调试器断点命中消息
const UINT MY_DEBUGGER_MEMDATA = 0x9999;		//内存读取消息
const HWND DEBUGGER_HWND = (HWND)【DebuggerHwnd】;

//滚动条移动速度常数
const int  HS_LargeChange[【NumberOfHS】] = { 【ArrayOfHSLarge】 };
const int  HS_SmallChange[【NumberOfHS】] = { 【ArrayOfHSSmall】 };
const int  VS_LargeChange[【NumberOfVS】] = { 【ArrayOfVSLarge】 };
const int  VS_SmallChange[【NumberOfVS】] = { 【ArrayOfVSSmall】 };

/* 左键按下判定变量 */
POINT CurUpPos;									//鼠标松开时在屏幕的坐标
RECT  CurrentWindowRect;						//当前窗体的位置和大小
/* 鼠标位置追踪 */
TRACKMOUSEEVENT tme;
/* 窗体移动中标记 */
bool bIsMoving = false;							//是否正在移动窗体
/* 窗体更改大小标记 */
bool bIsResizing = false;						//是否正在更改大小
/* 所有控件原先的WndProc地址 */
WNDPROC PrevStaticProc;							//STATIC
WNDPROC PrevEditProc;							//EDIT
WNDPROC PrevButtonProc;							//BUTTON
WNDPROC PrevComboProc;							//COMBOBOX
WNDPROC PrevListProc;							//LISTBOX
WNDPROC PrevScrollBarProc;						//SCROLL
WNDPROC PrevUpDownProc;							//UpDown
WNDPROC PrevProgressBarProc;					//ProgressBar
WNDPROC PrevSliderProc;							//Slider
WNDPROC PrevHotkeyProc;							//Hotkey
WNDPROC PrevListViewProc;						//ListView
WNDPROC PrevTreeViewProc;						//TreeView
WNDPROC PrevTabProc;							//Tab
WNDPROC PrevRichEditProc;						//RichEdit
WNDPROC PrevTimePickerProc;						//TimePicker
WNDPROC PrevMonthCalendarProc;					//MonthCalendar

/*============================================================================*/
/* 定义常数 */
/* 键盘功能键 */
const int SHIFT_KEY = 0x1;						//Shift键
const int CTRL_KEY = 0x2;						//Ctrl键
const int ALT_KEY = 0x4;						//Alt键

/* 鼠标按键 */
const int LEFT_BUTTON = 0x1;					//左键
const int RIGHT_BUTTON = 0x2;					//右键
const int MIDDLE_BUTTON = 0x4;					//中键

/* 滚轮按键 */
const int WHEEL_LEFT_BUTTON = 0x1;				//左键
const int WHEEL_RIGHT_BUTTON = 0x2;				//右键
const int WHEEL_MIDDLE_BUTTON = 0x4;			//中键
const int WHEEL_SHIFT_KEY = 4;					//Shift键
const int WHEEL_CTRL_KEY = 8;					//Ctrl键

//============================================================================
/* 定义窗体所有事件的函数原型 */
void Form_Load();								//窗体加载
void Form_Activate();							//窗体获得焦点
void Form_Deactivate();							//窗体失去焦点
void Form_KeyDown(int, int, bool);				//窗体按下按键
void Form_KeyUp(int, int);						//窗体松开按键
void Form_MouseDown(int, int, int, int);		//鼠标按键按下
void Form_MouseMove(int, int, int, int);		//鼠标移动
void Form_MouseUp(int, int, int, int);			//鼠标按键松开
void Form_MouseWheel(int, int, int, int, int);	//鼠标滚轮
void Form_Click();								//鼠标点击
void Form_DoubleClick(int, int, int, int);		//鼠标双击
void Form_MouseLeave();							//鼠标离开
void Form_Paint();								//窗体重绘
int Form_BeginResize(int);						//窗体开始改变大小
void Form_Resize(int, int, int);				//窗体改变大小
void Form_FinishResizing();						//窗体改变大小结束
int Form_BeginMove();							//窗体开始移动
void Form_Move(int, int);						//窗体移动
void Form_FinishMoving();						//窗体移动结束
int Form_QueryUnload();							//窗体将要关闭
void Form_Unload();								//窗体关闭
/*---------------------------------------------------------------------*/
void CreateAllControls();						//创建所有的控件过程
void GetCurrentRect();							//获取当前窗体的位置和大小
HWND GetCurrentHwnd();							//获取当前窗体的句柄
HINSTANCE GetCurrentHinstance();				//获取当前窗体的实例句柄
void SetWindowLongEx(HWND, DWORD, DWORD);		//设置窗体样式
void OrCalc(bool, long*, long);					//或运算过程（用于计算窗体样式）
void Breakpoint(long);							//断点命中过程
void WatchBreakpoint(int, void*, SIZE_T);		//监视点命中过程
void SuspendProcess();							//挂起进程过程
void ReregisterClass(LPCSTR, LPCSTR,
	WNDPROC*, WNDPROC);							//重新注册类过程
/*---------------------------------------------------------------------*/
/* 以下是所有控件事件的函数原型 */
【AllEventsDefHere】

//============================================================================
/* 计算当前的Shift值 */
int GetShiftValue()
{
	int ShiftValue = 0;
	if (GetAsyncKeyState(VK_SHIFT))		ShiftValue |= SHIFT_KEY;		//Shift键
	if (GetAsyncKeyState(VK_CONTROL))	ShiftValue |= CTRL_KEY;			//Ctrl键
	if (GetAsyncKeyState(VK_MENU))		ShiftValue |= ALT_KEY;			//Alt键
	return ShiftValue;
}

//============================================================================
/* 窗体主消息处理 */
LRESULT CALLBACK WndProc(HWND hWnd,	UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int TargetHmenu = 0;										//在某些事件中获得的控件的hMenu
	int CurrentPos = 0;											//滚动条事件中用来记录滑块位置的变量

	switch (uMsg)
	{
	case WM_MOUSEMOVE:											//鼠标移动
		TrackMouseEvent(&tme);										//继续追踪鼠标移动
		Form_MouseMove(wParam & ~(MK_CONTROL | MK_SHIFT),			//取得鼠标按键状态
			wParam & ~(wParam & ~(MK_CONTROL | MK_SHIFT)),			//取得系统功能键按键状态
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//取得鼠标坐标
		break;

	case WM_MOUSELEAVE:											//鼠标离开窗体
		Form_MouseLeave();
		TrackMouseEvent(NULL);										//停止追踪鼠标移动
		break;

	case WM_LBUTTONDOWN:										//左键按下
		SetCapture(hWnd);											//设置窗体鼠标捕获
		Form_MouseDown(1, GetShiftValue(),							//触发窗体左键按下消息
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//取得鼠标坐标
		break;

	case WM_LBUTTONUP:											//左键松开
		GetCursorPos(&CurUpPos);									//获取鼠标坐标
		ReleaseCapture();											//停止窗体鼠标捕获
		GetWindowRect(hWnd, &CurrentWindowRect);					//获得窗体坐标
		Form_MouseUp(1, GetShiftValue(),							//触发窗体左键松开消息
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//取得鼠标坐标
		/* 如果鼠标保持在窗体范围内的话就触发点击事件 */
		if ((CurUpPos.x > CurrentWindowRect.left) && (CurUpPos.x < CurrentWindowRect.right) &&
			(CurUpPos.y > CurrentWindowRect.top) && (CurUpPos.y < CurrentWindowRect.bottom))
		{
			Form_Click();
		}
		break;

	case WM_RBUTTONDOWN:										//右键按下
		SetCapture(hWnd);											//设置窗体鼠标捕获
		Form_MouseDown(2, GetShiftValue(), 							//触发窗体右键按下消息
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//取得鼠标坐标
		break;

	case WM_RBUTTONUP:											//右键松开
		GetCursorPos(&CurUpPos);									//获取鼠标坐标
		ReleaseCapture();											//停止窗体鼠标捕获
		GetWindowRect(hWnd, &CurrentWindowRect);					//获得窗体坐标
		Form_MouseUp(2, GetShiftValue(), 							//触发窗体右键松开消息
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//取得鼠标坐标
		/* 如果鼠标保持在窗体范围内的话就触发点击事件 */
		if ((CurUpPos.x > CurrentWindowRect.left) && (CurUpPos.x < CurrentWindowRect.right) &&
			(CurUpPos.y > CurrentWindowRect.top) && (CurUpPos.y < CurrentWindowRect.bottom))
		{
			Form_Click();
		}
		break;

	case WM_MBUTTONDOWN:										//中键按下
		SetCapture(hWnd);											//设置窗体鼠标捕获
		Form_MouseDown(4, GetShiftValue(), 							//触发窗体中键按下消息
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//取得鼠标坐标
		break;

	case WM_MBUTTONUP:											//中键松开
		GetCursorPos(&CurUpPos);									//获取鼠标坐标
		ReleaseCapture();											//停止窗体鼠标捕获
		GetWindowRect(hWnd, &CurrentWindowRect);					//获得窗体坐标
		Form_MouseUp(4, GetShiftValue(), 							//触发窗体中键松开消息
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//取得鼠标坐标
		/* 如果鼠标保持在窗体范围内的话就触发点击事件 */
		if ((CurUpPos.x > CurrentWindowRect.left) && (CurUpPos.x < CurrentWindowRect.right) &&
			(CurUpPos.y > CurrentWindowRect.top) && (CurUpPos.y < CurrentWindowRect.bottom))
		{
			Form_Click();
		}
		break;

	case WM_MOUSEWHEEL:											//鼠标滚轮
		//触发窗体滚轮事件
		Form_MouseWheel(GET_WHEEL_DELTA_WPARAM(wParam),								//大于0是向上滚动 向下滚动反之
			GET_KEYSTATE_WPARAM(wParam) & ~(MK_CONTROL | MK_SHIFT),					//获取鼠标按键状态
			GET_KEYSTATE_WPARAM(wParam) & ~(wParam & ~(MK_CONTROL | MK_SHIFT)),		//获取系统功能键状态
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));							//获取鼠标坐标
		break;

	case WM_LBUTTONDBLCLK:
		Form_DoubleClick(1,	GET_KEYSTATE_WPARAM(wParam) & ~(wParam & ~(MK_CONTROL | MK_SHIFT)),		//获取系统功能键状态
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));											//获取鼠标坐标
		break;

	case WM_RBUTTONDBLCLK:
		Form_DoubleClick(2, GET_KEYSTATE_WPARAM(wParam) & ~(wParam & ~(MK_CONTROL | MK_SHIFT)),		//获取系统功能键状态
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));											//获取鼠标坐标
		break;

	case WM_MBUTTONDBLCLK:
		Form_DoubleClick(4, GET_KEYSTATE_WPARAM(wParam) & ~(wParam & ~(MK_CONTROL | MK_SHIFT)),		//获取系统功能键状态
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));											//获取鼠标坐标
		break;

	case WM_SETFOCUS:											//窗体获得焦点
		Form_Activate();
		break;

	case WM_KILLFOCUS:											//窗体失去焦点
		if (bIsMoving)												//如果状态指窗体正在移动 但是 窗体失去焦点
		{															//说明窗体已经停止移动，触发窗体移动事件
			Form_FinishMoving();
		}
		Form_Deactivate();											//触发窗体失去焦点事件
		break;

	case WM_KEYDOWN:											//键盘按下按键
		Form_KeyDown(wParam, GetShiftValue(),						//获取Ascii码和系统功能键状态 
			(bool)((lParam >> 30) != 0));							//将lParam移30位以得到是否长按数据
		break;

	case WM_KEYUP:												//键盘松开按键
		Form_KeyUp(wParam, GetShiftValue());
		break;

	case WM_ERASEBKGND:											//窗体重绘
		Form_Paint();
		break;

	case WM_SIZING:												//窗体开始改变大小
		bIsResizing = true;											//记录为窗体正在更改大小
		if (Form_BeginResize(wParam) != 0)							//如果函数传回非0值则取消更改大小
		{
			bIsResizing = false;										//记录为窗体未在更改大小
			ReleaseCapture();											//释放窗体鼠标捕获，取消更改大小
		}
		break;

	case WM_SIZE:												//窗体改变大小
		GetCurrentRect();											//记录当前窗体位置和大小
		Form_Resize(wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));	//触发窗体移动事件
		break;

	case WM_SYSCOMMAND:
		if ((wParam == SC_MOVE) || (wParam == SC_MOVE + 2))		//拦截到窗体开始移动消息
		{
			bIsMoving = true;										//记录为窗体正在移动
			if (Form_BeginMove() != 0)								//如果函数传回非0值则取消移动
			{
				bIsMoving = false;
				return 0;
			}
			break;													//否则窗体正常移动
		}
		break;														//对于其它事件就放行
		
	case WM_COMMAND:											//控件事件
		switch (LOWORD(wParam))										//获取控件的hMenu
		{
		【ControlEventsHere】
		}
		break;

	case WM_NOTIFY:												//控件事件
		switch ((*(NMHDR*)lParam).idFrom)							//获取控件的hMenu
		{
		【ControlNotifyCodeHere】
		}
		break;

	case WM_HSCROLL:											//滚动条消息
	case WM_VSCROLL:
		TargetHmenu = (int)GetMenu((HWND)lParam);					//获取控件hMenu
		if (TargetHmenu)											//hMenu若不为零
		{
			switch (LOWORD(wParam))										//对不同的滚动方式进行处理
			{
			case SB_THUMBPOSITION:										//用户拖动滑块
			case SB_THUMBTRACK:
				CurrentPos = HIWORD(wParam);								//获取滑块位置
				break;

			case SB_PAGELEFT:											//向左快速移动
				CurrentPos = GetScrollPos((HWND)lParam, SB_CTL);			//获取滑块位置
				CurrentPos = CurrentPos - ((uMsg == WM_HSCROLL) ?
					HS_LargeChange[TargetHmenu - 1] :
					VS_LargeChange[TargetHmenu - 1]);
				break;

			case SB_PAGERIGHT:											//向右快速移动
				CurrentPos = GetScrollPos((HWND)lParam, SB_CTL);			//获取滑块位置
				CurrentPos = CurrentPos + ((uMsg == WM_HSCROLL) ?
					HS_LargeChange[TargetHmenu - 1] :
					VS_LargeChange[TargetHmenu - 1]);
				break;

			case SB_LINELEFT:											//向左慢速移动
				CurrentPos = GetScrollPos((HWND)lParam, SB_CTL);			//获取滑块位置
				CurrentPos = CurrentPos - ((uMsg == WM_HSCROLL) ?
					HS_SmallChange[TargetHmenu - 1] :
					VS_SmallChange[TargetHmenu - 1]);
				break;

			case SB_LINERIGHT:											//向右慢速移动
				CurrentPos = GetScrollPos((HWND)lParam, SB_CTL);			//获取滑块位置
				CurrentPos = CurrentPos + ((uMsg == WM_HSCROLL) ?
					HS_SmallChange[TargetHmenu - 1] :
					VS_SmallChange[TargetHmenu - 1]);
				break;

			case SB_ENDSCROLL:											//停止拖动
				CurrentPos = GetScrollPos((HWND)lParam, SB_CTL);			//获取滑块位置
				break;
			}
			SetScrollPos((HWND)lParam, SB_CTL, CurrentPos, TRUE);		//更新滑块位置
		}
		break;
		
	case WM_MOVE:												//窗体移动
		GetCurrentRect();											//记录当前窗体位置和大小
		Form_Move(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));		//触发窗体移动事件
		break;
	
	case WM_CAPTURECHANGED:										//窗体鼠标捕获消息
		if (bIsMoving && (GetAsyncKeyState(VK_LBUTTON) == 0))		//如果状态指窗体正在移动 但是 鼠标左键已经松开
		{															//说明窗体已经停止移动，触发窗体移动事件
			bIsMoving = false;
			Form_FinishMoving();
		}
		break;

	case WM_NCHITTEST:											//捕获鼠标在窗体中所有位置的事件
		if (bIsMoving && (GetCapture() != GetCurrentHwnd()))		//如果状态指窗体正在移动 但是 窗体却不获得鼠标捕获
		{															//说明窗体已经停止移动，触发窗体移动事件
			bIsMoving = false;
			Form_FinishMoving();
		}
		break;

	case WM_EXITSIZEMOVE:										//窗体退出调整大小或者退出移动状态
		if (bIsResizing)											//如果窗体正在更改大小
		{
			bIsResizing = false;
			Form_FinishResizing();										//触发窗体改变大小结束事件
		}
		break;

	case WM_CLOSE:												//窗体将要关闭
		if (Form_QueryUnload() != 0)								//窗体取消关闭
		{
			return 0;													//如果函数返回非0值就取消窗体关闭
		}
		break;													//否则窗体继续关闭

	case WM_DESTROY:											//窗体关闭
		Form_Unload();
		return 0;
	}

	return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* STATIC控件消息处理 */
LRESULT CALLBACK StaticProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)								//附加值2
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);					//获取当前控件的标识符
	switch (CurrCtlHmenu)
	{
	【StaticProcCode】
	}

	return CallWindowProc(PrevStaticProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* EDIT控件消息处理 */
LRESULT CALLBACK EditProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);					//获取当前控件的标识符
	switch (CurrCtlHmenu)
	{
	【EditProcCode】
	}

	return CallWindowProc(PrevEditProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* BUTTON控件消息处理 */
LRESULT CALLBACK ButtonProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);					//获取当前控件的标识符
	switch (CurrCtlHmenu)
	{
	【ButtonProcCode】
	}

	return CallWindowProc(PrevButtonProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* COMBOBOX控件消息处理 */
LRESULT CALLBACK ComboProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	【ComboProcCode】
	}

	return CallWindowProc(PrevComboProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* LISTBOX控件消息处理 */
LRESULT CALLBACK ListProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	【ListProcCode】
	}

	return CallWindowProc(PrevListProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* SCROLLBAR控件消息处理 */
LRESULT CALLBACK ScrollBarProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	SCROLLINFO sInfo;

	switch (CurrCtlHmenu)
	{
	【ScrollBarProcCode】
	}

	return CallWindowProc(PrevScrollBarProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* UpDown控件消息处理 */
LRESULT CALLBACK UpDownProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	SCROLLINFO si;

	switch (CurrCtlHmenu)
	{
	【UpDownProcCode】
	}

	return CallWindowProc(PrevUpDownProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* ProgressBar控件消息处理 */
LRESULT CALLBACK ProgressBarProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	【ProgressBarProcCode】
	}
	return CallWindowProc(PrevProgressBarProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* Slider控件消息处理 */
LRESULT CALLBACK SliderProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	【SliderProcCode】
	}
	return CallWindowProc(PrevSliderProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* Hotkey控件消息处理 */
LRESULT CALLBACK HotkeyProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	【HotkeyProcCode】
	}
	return CallWindowProc(PrevHotkeyProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* ListView控件消息处理 */
LRESULT CALLBACK ListViewProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	【ListViewProcCode】
	}
	return CallWindowProc(PrevListViewProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* TreeView控件消息处理 */
LRESULT CALLBACK TreeViewProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	【TreeViewProcCode】
	}
	return CallWindowProc(PrevTreeViewProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* Tab控件消息处理 */
LRESULT CALLBACK TabProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	【TabProcCode】
	}
	return CallWindowProc(PrevTabProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* RichEdit控件消息处理 */
LRESULT CALLBACK RichEditProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	【RichEditProcCode】
	}
	return CallWindowProc(PrevRichEditProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* TimePicker控件消息处理 */
LRESULT CALLBACK TimePickerProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);					//获取当前控件的标识符
	switch (CurrCtlHmenu)
	{
	【TimePickerProcCode】
	}
	return CallWindowProc(PrevTimePickerProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* MonthCalendar控件消息处理 */
LRESULT CALLBACK MonthCalendarProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);					//获取当前控件的标识符
	switch (CurrCtlHmenu)
	{
	【MonthCalendarProcCode】
	}
	return CallWindowProc(PrevMonthCalendarProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* 窗体计时器消息处理 */
void CALLBACK TimerProc(HWND hWnd, UINT uMsg, UINT iTimerID, DWORD dwTime)
{
	switch (iTimerID)
	{
【AllTimerIDHere】
	}
}

//============================================================================
/* 定义各种控件通用的控件类（包含每个控件相同的属性和过程），
之后的控件类便可以继承这个类的属性和过程 */
class MyControls
{
public:
	int			hMenu;							//控件的hMenu（唯一标识符）
	LONG		Left;							//水平位置
	LONG		Top;							//垂直位置
	LONG		Width;							//窗体宽度
	LONG		Height;							//窗体高度
	HDC			hDC;							//设备上下文句柄
	HWND		CurrentHwnd;					//当前控件的句柄
	bool		Visible;						//是否可视
	bool		Enabled;						//是否激活

	//删除控件
	void Unload()
	{
		DestroyWindow(CurrentHwnd);
	}

	/*
	更改控件位置
	NewLeft: 新的X坐标
	NewTop: 新的Y坐标
	NewWidth: 新的宽度
	NewHeight: 新的高度
	*/
	void Move(int NewLeft, int NewTop, int NewWidth, int NewHeight)
	{
		SetWindowPos(CurrentHwnd, 0, NewLeft, NewTop, NewWidth, NewHeight, 0);
		Left = NewLeft;
		Top = NewTop;
		Width = NewWidth;
		Height = NewHeight;
	}

	/*
	设置控件是否可视
	bVisible: 控件是否可视
	*/
	void SetVisible(bool bVisible)
	{
		ShowWindow(CurrentHwnd, bVisible ? SW_SHOW : SW_HIDE);
		Visible = bVisible;
	}

	/*
	设置控件是否禁用
	IsEnabled: 是否启用窗体
	*/
	void SetEnabled(bool IsEnabled)
	{
		EnableWindow(CurrentHwnd, IsEnabled);
		Enabled = IsEnabled;
	}
};

//============================================================================
/* 定义按钮控件、组框、单选框、多选框通用的按钮控件类（包含它们共有的属性和过程），
它们的控件类便可以继承这个类的属性和过程 */
class ButtonPublicClass : public MyControls
{
public:
	char*		Text;							//文本
	DWORD		TextPos;						//文本位置

	/*
	设置按钮的文本
	NewText: 新的文本
	*/
	void SetText(char* NewText)
	{
		Text = NewText;
		SetWindowText(CurrentHwnd, NewText);
	}

	/*
	以数字设置按钮的文本
	NewText: 有效的数值表达式
	*/
	void SetText(int NewText)
	{
		char Buffer[255];
		itoa(NewText, Buffer, 10);
		SetWindowText(CurrentHwnd, Buffer);
	}

	/* 模拟鼠标点击的操作 */
	void Click()
	{
		SendMessage(CurrentHwnd, BM_CLICK, 0, 0);
	}

	/* 获取文本 */
	char* GetText()
	{
		char* tmp = new char[255];
		GetWindowText(CurrentHwnd, tmp, 255);
		return tmp;
	}
};

//============================================================================
class MyTimer
{
public:
	UINT_PTR	ID;								//计时器ID
	UINT		Interval;						//计时间隔
	bool		Enabled;						//计时器是否被激活

	//创建计时器
	bool Create()
	{
		Enabled = (bool)(SetTimer(GetCurrentHwnd(), ID, Interval, TimerProc) != 0);
		return Enabled;
	}
	
	//删除计时器
	void Kill()
	{
		KillTimer(GetCurrentHwnd(), ID);
		Enabled = false;
	}
};

//============================================================================
class MyImage : public MyControls
{
public:
	//创建图片控件
	HWND Create()
	{
		CurrentHwnd = CreateWindowEx(0, "MyStatic", "", WS_VISIBLE | WS_CHILD | SS_BLACKFRAME | SS_NOTIFY,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		SetVisible(Visible);
		SetEnabled(Enabled);
		return CurrentHwnd;
	}
};

//============================================================================
class MyLabel : public MyControls
{
public:
	int			TextPos;						//文本位置
	char*		Caption;						//标签内容
	bool		BlackBorder;					//是否有黑色边框
	bool		BlackFilled;					//是否以黑色填充
	bool		AutoNextLine;					//是否自动换行
	bool		AutoEllipsis;					//是否自动添加省略号

	//创建标签控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD | SS_NOTIFY;		//初始化控件样式

		//计算控件的样式
		OrCalc(BlackBorder, &lStyle, WS_BORDER);				//黑色边框
		OrCalc(BlackFilled, &lStyle, SS_BLACKRECT);				//黑色填充
		switch (TextPos)										//文本位置
		{
		case 1:														//中
			lStyle |= SS_CENTER;
			break;

		case 2:														//右
			lStyle |= SS_RIGHT;
			break;
		}
		OrCalc(AutoNextLine, &lStyle, SS_EDITCONTROL);			//自动换行
		OrCalc(AutoEllipsis, &lStyle, SS_ENDELLIPSIS);			//自动添加省略号

		//创建控件
		CurrentHwnd = CreateWindowEx(0, "MyStatic", Caption, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);
		return CurrentHwnd;
	}

	/*
	设置标签标题
	NewCaption: 新的标题字符串
	*/
	void SetCaption(char* NewCaption)
	{
		SetWindowText(CurrentHwnd, NewCaption);			//更改标签标题
		Caption = NewCaption;
	}

	/*
	用数值来设置标签标题。
	NewCaption: 有效的数值表达式
	*/
	void SetCaption(int NewCaption)
	{
		char tmp[255];
		itoa(NewCaption, tmp, 10);
		SetCaption(tmp);
	}

	/*
	设置标签文本位置
	NewPos: 标签的文本位置。
	*/
	void SetTextPos(int NewPos)
	{
		switch (NewPos)									//根据不同设置值更改标签文本位置
		{
		case 0:												//左
			SetWindowLongEx(CurrentHwnd, SS_LEFT, SS_CENTER | SS_RIGHT);
			break;

		case 1:												//中
			SetWindowLongEx(CurrentHwnd, SS_CENTER, SS_LEFT | SS_RIGHT);
			break;

		case 2:												//右
			SetWindowLongEx(CurrentHwnd, SS_RIGHT, SS_LEFT | SS_CENTER);
			break;

		}
		TextPos = NewPos;
	}

	/*
	设置标签控件是否自动换行
	bAutoNextLine: 是否自动换行
	*/
	void SetAutoNextLine(bool bAutoNextLine)
	{
		bAutoNextLine ?
			SetWindowLongEx(CurrentHwnd, SS_EDITCONTROL, 0):
			SetWindowLongEx(CurrentHwnd, 0, SS_EDITCONTROL);
		AutoNextLine = bAutoNextLine;
	}

	/*
	设置标签控件是否自动添加省略号
	bAutoEllipsis: 是否自动添加省略号
	*/
	void SetAutoEllipsis(bool bAutoEllipsis)
	{
		bAutoEllipsis ?
			SetWindowLongEx(CurrentHwnd, SS_ENDELLIPSIS, 0):
			SetWindowLongEx(CurrentHwnd, 0, SS_ENDELLIPSIS);
		AutoEllipsis = bAutoEllipsis;
	}

	/*
	设置标签控件是否以黑色填充
	bBlackFilled: 是否以黑色填充
	*/
	void SetBlackFilled(bool bBlackFilled)
	{
		bBlackFilled ?
			SetWindowLongEx(CurrentHwnd, SS_BLACKRECT, 0):
			SetWindowLongEx(CurrentHwnd, 0, SS_BLACKRECT);
		BlackFilled = bBlackFilled;
	}
};

//============================================================================
class MyEdit : public MyControls
{
public:
	char*		Text;							//文本
	bool		AutoHScroll;					//自动水平滚动
	bool		AutoVScroll;					//自动垂直滚动
	int			TextPos;						//文本位置
	bool		ForceLowercase;					//强制小写
	bool		ForceUppercase;					//强制大写
	bool		ForceNumber;					//强制数字
	bool		IsPassword;						//密码文本
	char		PasswordChar;					//密码字符
	bool		ReadOnly;						//文本只读
	bool		BlackBorder;					//黑色边框
	bool		ClientEdgeBorder;				//立体边框
	bool		Multiline;						//多行文本
	int			ScrollBars;						//滚动条

	//创建文本控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;					//初始化控件样式

		//计算控件的样式
		OrCalc(AutoHScroll, &lStyle, ES_AUTOHSCROLL);			//自动水平滚动
		OrCalc(AutoVScroll, &lStyle, ES_AUTOVSCROLL);			//自动垂直滚动
		switch (TextPos)										//文本位置
		{
		case 1:														//中
			lStyle |= ES_CENTER;
			break;

		case 2:														//右
			lStyle |= ES_RIGHT;
			break;
		}
		OrCalc(ForceLowercase, &lStyle, ES_LOWERCASE);			//强制小写
		OrCalc(ForceUppercase, &lStyle, ES_UPPERCASE);			//强制大写
		OrCalc(ForceNumber, &lStyle, ES_NUMBER);				//强制数字
		OrCalc(IsPassword, &lStyle, ES_PASSWORD);				//密码文本
		OrCalc(ReadOnly, &lStyle, ES_READONLY);					//文本只读
		OrCalc(BlackBorder, &lStyle, WS_BORDER);				//黑色边框
		OrCalc(Multiline, &lStyle, ES_MULTILINE);				//多行文本
		switch (ScrollBars)										//滚动条
		{
		case 1:														//水平
			lStyle |= WS_HSCROLL;
			break;

		case 2:														//垂直
			lStyle |= WS_VSCROLL;
			break;

		case 3:														//两个都有
			lStyle |= WS_HSCROLL | WS_VSCROLL;
			break;
		}

		//创建控件
		CurrentHwnd = CreateWindowEx(ClientEdgeBorder ? WS_EX_CLIENTEDGE : 0, "MyEdit", Text, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		//设置密码字符
		if (IsPassword)		SendMessage(CurrentHwnd, EM_SETPASSWORDCHAR, (WPARAM)PasswordChar, 0);

		return CurrentHwnd;
	}

	/*
	设置文本框文本
	NewText: 新的文本
	*/
	void SetText(char* NewText)
	{
		Text = NewText;
		SetWindowText(CurrentHwnd, NewText);
	}

	/*
	以数字设置文本框文本
	NewText: 有效的数字表达式
	*/
	void SetText(int NewText)
	{
		char Buffer[255];
		itoa(NewText, Buffer, 10);
		SetWindowText(CurrentHwnd, Buffer);
	}

	/*
	设置文本框是否强制要求小写
	bLowerCase: 是否强制要求小写
	*/
	void SetForceLowerCase(bool bLowerCase)
	{
		bLowerCase ?
			SetWindowLongEx(CurrentHwnd, ES_LOWERCASE, 0) :
			SetWindowLongEx(CurrentHwnd, 0, ES_LOWERCASE);
		ForceLowercase = bLowerCase;
	}

	/*
	设置文本框是否强制要求大写
	bUpperCase: 是否强制要求大写
	*/
	void SetForceUpperCase(bool bUpperCase)
	{
		bUpperCase ?
			SetWindowLongEx(CurrentHwnd, ES_UPPERCASE, 0) :
			SetWindowLongEx(CurrentHwnd, 0, ES_UPPERCASE);
		ForceUppercase = bUpperCase;
	}

	/*
	设置文本框是否强制要求数字
	bNumber: 是否强制要求数字
	*/
	void SetForceNumber(bool bNumber)
	{
		bNumber ?
			SetWindowLongEx(CurrentHwnd, ES_NUMBER, 0) :
			SetWindowLongEx(CurrentHwnd, 0, ES_NUMBER);
		ForceNumber = bNumber;
	}

	/*
	设置文本框是否文本只读
	bReadOnly: 是否文本只读
	*/
	void SetReadOnly(bool bReadOnly)
	{
		SendMessage(CurrentHwnd, EM_SETREADONLY, bReadOnly, 0);
		ReadOnly = bReadOnly;
	}

	/* 清空文本框的撤销缓存 */
	void EmptyUndoBuffer()
	{
		SendMessage(CurrentHwnd, EM_EMPTYUNDOBUFFER, 0, 0);
	}

	/* 获取文本框能看到的第一行 */
	long GetFirstVisibleLine()
	{
		return (long)SendMessage(CurrentHwnd, EM_GETFIRSTVISIBLELINE, 0, 0);
	}

	/* 获取文本限制范围 */
	long GetLimitText()
	{
		return (long)SendMessage(CurrentHwnd, EM_GETLIMITTEXT, 0, 0);
	}

	/*
	设置文本限制范围
	NewLimit: 新的文本输入字数限制。如果为0则无限制。
	*/
	void SetLimitText(long NewLimit)
	{
		SendMessage(CurrentHwnd, EM_SETLIMITTEXT, NewLimit, 0);
	}

	/* 获取文本的行数 */
	long GetLineCount()
	{
		return (long)SendMessage(CurrentHwnd, EM_GETLINECOUNT, 0, 0);
	}

	/*
	设置文本框的左边距
	NewMargin: 新的左边距
	*/
	void SetLeftMargin(int NewMargin)
	{
		SendMessage(CurrentHwnd, EM_SETMARGINS, EC_LEFTMARGIN, MAKELPARAM(NewMargin, 0));
	}

	/*
	设置文本框的右边距
	NewMargin: 新的右边距
	*/
	void SetRightMargin(int NewMargin)
	{
		SendMessage(CurrentHwnd, EM_SETMARGINS, EC_RIGHTMARGIN, MAKELPARAM(0, NewMargin));
	}

	/* 获取文本框的左边距 */
	int GetLeftMargin()
	{
		return (int)(LOWORD(SendMessage(CurrentHwnd, EM_GETMARGINS, 0, 0)));
	}

	/* 获取文本框的右边距 */
	int GetRightMargin()
	{
		return (int)(HIWORD(SendMessage(CurrentHwnd, EM_GETMARGINS, 0, 0)));
	}


	/*
	设置文本框的密码文本
	NewPasswordChar: 新的密码字符
	*/
	void SetPasswordChar(char NewPasswordChar)
	{
		SendMessage(CurrentHwnd, EM_SETPASSWORDCHAR, (WPARAM)NewPasswordChar, 0);
		PasswordChar = NewPasswordChar;
	}

	/*
	设置文本框的密码文本
	NewPasswordChar: 新的密码字符的Ascii码
	*/
	void SetPasswordChar(int NewPasswordChar)
	{
		SendMessage(CurrentHwnd, EM_SETPASSWORDCHAR, (WPARAM)NewPasswordChar, 0);
		PasswordChar = (char)NewPasswordChar;
	}

	/* 获取文本框的密码文本 */
	char GetPasswordChar()
	{
		return (char)SendMessage(CurrentHwnd, EM_GETPASSWORDCHAR, 0, 0);
	}

	/* 获取文本框的选取文本长度 */
	DWORD GetSelLength()
	{
		DWORD StartPos = 0, EndPos = 0;
		SendMessage(CurrentHwnd, EM_GETSEL, (WPARAM)&StartPos, (LPARAM)&EndPos);
		return (EndPos - StartPos);
	}

	/* 获取文本框的选取文本开头 */
	DWORD GetSelStart()
	{
		DWORD StartPos = 0, EndPos = 0;
		SendMessage(CurrentHwnd, EM_GETSEL, (WPARAM)&StartPos, (LPARAM)&EndPos);
		return StartPos;
	}

	/*
	设置文本框的选取文本长度
	如果StartPos为0，EndPos为-1，则文本全选；
	如果StartPos为-1，则文本取消选择；
	如果EndPos为-1，则文本选取范围是StartPos到文本末尾。
	*/
	void SetSel(DWORD StartPos, DWORD EndPos)
	{
		SendMessage(CurrentHwnd, EM_SETSEL, (WPARAM)StartPos, (LPARAM)EndPos);
	}

	/*
	设置文本框的光标位置
	StartPos: 新的光标位置
	如果StartPos为-1，则文本取消选择。
	*/
	void SetSelStart(DWORD StartPos)
	{
		SendMessage(CurrentHwnd, EM_SETSEL, (WPARAM)StartPos, (LPARAM)StartPos);
	}

	/*
	设置文本框的选取范围
	SelLength: 选取范围的长度
	如果SelLength是-1则文本框的选取范围是当前光标所在位置到文本末尾。
	*/
	void SetSelLength(DWORD SelLength)
	{
		DWORD CurrPos = GetSelStart();
		SendMessage(CurrentHwnd, EM_SETSEL, (WPARAM)CurrPos, (LPARAM)(CurrPos + SelLength));
	}

	/*
	获取指定行的长度
	lnNumber: 行号。如果为-1则返回未选取的文本长度。
	*/
	long GetLineLength(long lnNumber)
	{
		return (long)SendMessage(CurrentHwnd, EM_LINELENGTH, lnNumber, 0);
	}

	/*
	水平滚动指定数量的列数
	CharCount: 水平滚动的列数
	如果文本框是单行文本框则返回false，如果是多行文本框则返回true
	*/
	bool HScroll(long CharCount)
	{
		return (bool)SendMessage(CurrentHwnd, EM_LINESCROLL, (WPARAM)CharCount, 0);
	}

	/*
	垂直滚动指定数量的行数
	LineCount: 垂直滚动的行数
	如果文本框是单行文本框则返回false，如果是多行文本框则返回true
	*/
	bool VScroll(long LineCount)
	{
		return (bool)SendMessage(CurrentHwnd, EM_LINESCROLL, 0, (LPARAM)LineCount);
	}

	/*
	把当前选择的文本替换成指定的字符串
	bCanUndo: 此操作是否可撤销
	NewText: 新的文本
	*/
	void SetSelText(bool bCanUndo, char* NewText)
	{
		SendMessage(CurrentHwnd, EM_REPLACESEL, (WPARAM)bCanUndo, (LPARAM)NewText);
	}

	/* 获取当前光标所在的行 */
	long GetCurrLine()
	{
		DWORD CurrSel = SendMessage(CurrentHwnd, EM_GETSEL, 0, 0);
		return (long)SendMessage(CurrentHwnd, EM_LINEFROMCHAR, (WPARAM)(CurrSel / 65536), 0);
	}

	/* 获取当前光标所在的列 */
	long GetCurrCol()
	{
		DWORD CurrSel = SendMessage(CurrentHwnd, EM_GETSEL, 0, 0);
		DWORD CurrLnSel = SendMessage(CurrentHwnd, EM_LINEINDEX, (WPARAM)-1, 0);
		return ((CurrSel / 65536) - CurrLnSel + 1);
	}

	/* 复制操作 */
	void Copy()
	{
		SendMessage(CurrentHwnd, WM_COPY, 0, 0);
	}

	/* 剪切操作 */
	void Cut()
	{
		SendMessage(CurrentHwnd, WM_CUT, 0, 0);
	}

	/* 撤销操作 */
	void Undo()
	{
		SendMessage(CurrentHwnd, EM_UNDO, 0, 0);
	}

	/* 清除选择的文本 */
	void Clear()
	{
		SendMessage(CurrentHwnd, WM_CLEAR, 0, 0);
	}

	/* 粘贴操作 */
	void Paste()
	{
		SendMessage(CurrentHwnd, WM_PASTE, 0, 0);
	}

	/* 获取文本长度 */
	long GetTextLength()
	{
		return (long)SendMessage(CurrentHwnd, WM_GETTEXTLENGTH, 0, 0);
	}

	/* 滚动到光标 */
	void ScrollToCaret()
	{
		SendMessage(CurrentHwnd, EM_SCROLLCARET, 0, 0);
	}

	/* 获取文本 */
	char* GetText()
	{
		int Length = GetWindowTextLength(CurrentHwnd) + 1;
		char* tmp = new char[Length];
		GetWindowText(CurrentHwnd, tmp, Length);
		return tmp;
	}

	/* 使文本框获取焦点 */
	void SetFocus()
	{
		::SetFocus(CurrentHwnd);
	}
};

//============================================================================
class MyFrame : public ButtonPublicClass
{
public:
	//创建组框控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD | BS_GROUPBOX;					//初始化控件样式
		lStyle |= TextPos;													//添加文本位置到样式里

		//创建控件
		CurrentHwnd = CreateWindowEx(0, "MyButton", Text, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}
};

//============================================================================
class MyButton : public ButtonPublicClass
{
public:
	bool		ClientEdgeBorder;				//立体边框
	bool		Flat;							//扁平
	bool		BlackBorder;					//黑色边框

	//创建当前的按钮控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;								//初始化控件样式
		LONG ExStyle = 0;													//扩展样式

		//计算控件的样式
		lStyle |= TextPos;													//添加文本位置到样式里
		OrCalc(ClientEdgeBorder, &ExStyle, WS_EX_CLIENTEDGE);				//立体边框
		OrCalc(Flat, &lStyle, BS_FLAT);										//扁平
		OrCalc(BlackBorder, &lStyle, WS_BORDER);							//黑色边框

		//创建控件
		CurrentHwnd = CreateWindowEx(ExStyle, "MyButton", Text, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}
};

//============================================================================
class MyCheckBox : public ButtonPublicClass
{
public:
	bool		ClientEdgeBorder;				//立体边框
	bool		Flat;							//扁平
	bool		BlackBorder;					//黑色边框
	bool		PushLike;						//按钮形式

	//创建多选框控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX;				//初始化控件样式
		LONG ExStyle = 0;													//扩展样式

		//计算控件的样式
		lStyle |= TextPos;													//添加文本位置到样式里
		OrCalc(ClientEdgeBorder, &ExStyle, WS_EX_CLIENTEDGE);				//立体边框
		OrCalc(Flat, &lStyle, BS_FLAT);										//扁平
		OrCalc(BlackBorder, &lStyle, WS_BORDER);							//黑色边框
		OrCalc(PushLike, &lStyle, BS_PUSHLIKE);								//按钮样式

		//创建控件
		CurrentHwnd = CreateWindowEx(ExStyle, "MyButton", Text, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/* 设置复选框的勾选状态 */
	void SetChecked(bool bChecked)
	{
		SendMessage(CurrentHwnd, BM_SETCHECK, bChecked ? BST_CHECKED : BST_UNCHECKED, 0);
	}

	/* 获取复选框的勾选状态 */
	bool GetChecked()
	{
		return (SendMessage(CurrentHwnd, BM_GETCHECK, 0, 0) == BST_CHECKED);
	}
};

//============================================================================
class MyOption : public MyCheckBox
{
public:
	//创建单选框控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD | BS_AUTORADIOBUTTON;			//初始化控件样式
		LONG ExStyle = 0;													//扩展样式

		//计算控件的样式
		lStyle |= TextPos;													//添加文本位置到样式里
		OrCalc(ClientEdgeBorder, &ExStyle, WS_EX_CLIENTEDGE);				//立体边框
		OrCalc(Flat, &lStyle, BS_FLAT);										//扁平
		OrCalc(BlackBorder, &lStyle, WS_BORDER);							//黑色边框
		OrCalc(PushLike, &lStyle, BS_PUSHLIKE);								//按钮样式

		//创建控件
		CurrentHwnd = CreateWindowEx(ExStyle, "MyButton", Text, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}
};

//============================================================================
class MyCombo : public MyControls
{
public:
	DWORD		VerticalScrollBar;				//垂直滚动条
	bool		AutoHscroll;					//自动水平滚动
	bool		ForceLowerCase;					//强制小写
	bool		ForceUppercase;					//强制大写
	bool		DropDownStyle;					//列表样式
	bool		AutoSort;						//自动排列

	//创建组合框控件
	HWND Create()
	{
		LONG lStyle = WS_CHILD | CBS_HASSTRINGS | WS_VISIBLE;				//初始化控件样式

		lStyle |= VerticalScrollBar;
		OrCalc(AutoHscroll, &lStyle, CBS_AUTOHSCROLL);
		OrCalc(ForceLowerCase, &lStyle, CBS_LOWERCASE);
		OrCalc(ForceUppercase, &lStyle, CBS_UPPERCASE);
		OrCalc(!DropDownStyle, &lStyle, CBS_DROPDOWN);
		OrCalc(AutoSort, &lStyle, CBS_SORT);

		CurrentHwnd = CreateWindowEx(0, "MyCombo", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	添加指定的文本到组合框中的指定位置
	strAdd: 需要添加的文本
	ListIndex: 文本添加到的位置。如果是-1 则添加到组合框的末尾
	如果添加成功则返回对应的列表序号，如果失败则返回CB_ERR (-1)
	*/
	int AddItem(char* strAdd, int ListIndex = -1)
	{
		return (int)SendMessage(CurrentHwnd, CB_INSERTSTRING, (WPARAM)ListIndex, (LPARAM)strAdd);
	}

	/*
	添加指定的数值到组合框中的指定位置
	strAdd: 需要添加的数值
	ListIndex: 文本添加到的位置。如果是-1 则添加到组合框的末尾
	如果添加成功则返回对应的列表序号，如果失败则返回CB_ERR (-1)
	*/
	int AddItem(int strAdd, int ListIndex = -1)
	{
		char Buffer[255];
		itoa(strAdd, Buffer, 10);
		return (int)SendMessage(CurrentHwnd, CB_INSERTSTRING, (WPARAM)ListIndex, (LPARAM)Buffer);
	}

	/*
	删除指定序号的列表项
	ListIndex: 指定的列表项序号
	如果执行成功则返回列表中剩余的项目数，如果失败则返回CB_ERR (-1)
	*/
	int RemoveItem(int ListIndex)
	{
		return (int)SendMessage(CurrentHwnd, CB_DELETESTRING, (WPARAM)ListIndex, 0);
	}

	/* 清空所有列表项 */
	void Clear()
	{
		SendMessage(CurrentHwnd, CB_RESETCONTENT, 0, 0);
	}

	/*
	在组合框中列出指定目录的文件
	FilePath: 文件路径
	FileType: 筛选的文件类型，为DDL_*的常数
	如果执行成功则返回最后一个添加的列表项的序号，如果失败则返回CB_ERR (-1)，如果列表空间不足则返回CB_ERRSPACE (-2)
	*/
	int ListDirFiles(LPCTSTR FilePath, DWORD FileType = 0)
	{
		return (int)SendMessage(CurrentHwnd, CB_DIR, (WPARAM)FileType, (LPARAM)FilePath);
	}

	/*
	在组合框中查找有指定文本的列表项
	StrFind: 需要查找的文本
	StartIndex: 从哪个列表项开始查找。如果为-1则从头开始查找
	FullMatch: 是否需要全字匹配。如果不是，则检测文本的开头是否包含查找的文本
	如果执行成功则返回找到的列表项的序号，如果失败则返回CB_ERR (-1)
	*/
	int FindItem(char* strFind, int StartIndex = -1, bool FullMatch = true)
	{
		return (int)SendMessage(CurrentHwnd, FullMatch ? CB_FINDSTRINGEXACT : CB_FINDSTRING,
			(WPARAM)StartIndex, (LPARAM)strFind);
	}

	/* 获取组合框中的项目数量 */
	int GetListCount()
	{
		return (int)SendMessage(CurrentHwnd, CB_GETCOUNT, 0, 0);
	}

	/*
	获取组合框中当前选取的列表项
	返回值为当前选取的列表项序号。如果没有列表项被选取则返回CB_ERR (-1)
	*/
	int GetSelItem()
	{
		return (int)SendMessage(CurrentHwnd, CB_GETCURSEL, 0, 0);
	}

	/*
	获取组合框的列表的下拉状态
	如果下拉列表可视则返回true
	*/
	bool GetDroppedState()
	{
		return (bool)SendMessage(CurrentHwnd, CB_GETDROPPEDSTATE, 0, 0);
	}

	/* 获取组合框的下拉列表的宽度 */
	int GetDroppedWidth()
	{
		return (int)SendMessage(CurrentHwnd, CB_GETDROPPEDWIDTH, 0, 0);
	}

	/*
	获取组合框中的文本框的选取范围
	outSelStart: 用来存放得到的文本选取区域的开头
	outSelEnd: 用来存放得到的文本选取区域的结尾
	*/
	void GetEditSel(int* outSelStart, int* outSelEnd)
	{
		SendMessage(CurrentHwnd, CB_GETEDITSEL, (WPARAM)outSelStart, (LPARAM)outSelEnd);
	}

	/* 获取组合框中其中一个列表项的高度 */
	int GetItemHeight()
	{
		return (int)SendMessage(CurrentHwnd, CB_GETITEMHEIGHT, 0, 0);
	}

	/*
	获取组合框中指定列表项的文本
	ListIndex: 指定的列表项序号
	*/
	char* GetListText(int ListIndex)
	{
		int Length = (int)SendMessage(CurrentHwnd, CB_GETLBTEXTLEN, ListIndex, 0);
		char* tmp = new char[Length];
		if (SendMessage(CurrentHwnd, CB_GETLBTEXT, ListIndex, (LPARAM)tmp) != CB_ERR)
		{
			return tmp;
		}
		else
		{
			return "";
		}
	}

	/* 获取文本 */
	char* GetText()
	{
		int Length = GetWindowTextLength(CurrentHwnd) + 1;
		char* tmp = new char[Length];
		GetWindowText(CurrentHwnd, tmp, Length);
		return tmp;
	}

	/*
	设置文本
	NewText: 新的文本
	*/
	void SetText(char* NewText)
	{
		SetWindowText(CurrentHwnd, NewText);
	}

	/*
	以数字设置文本
	NewText: 有效的数字表达式
	*/
	void SetText(int NewText)
	{
		char Buffer[255];
		itoa(NewText, Buffer, 10);
		SetWindowText(CurrentHwnd, Buffer);
	}

	/* 获取组合框的下拉列表中最少可以看见的列表数 */
	int GetMinimumVisibleItems()
	{
		return (int)SendMessage(CurrentHwnd, CB_GETMINVISIBLE, 0, 0);
	}

	/*
	获取组合框可视的第一个列表项的序号
	如果函数执行成功则返回组合框的下拉列表中的第一个可视的列表项序号，如果失败则返回CB_ERR (-1)
	*/
	int GetFirstVisibleItem()
	{
		return (int)SendMessage(CurrentHwnd, CB_GETTOPINDEX, 0, 0);
	}

	/*
	设置文本框最长的文本长度
	LimitLength: 限制的长度，如果为0则为系统默认
	*/
	void SetLimitLength(int LimitLength)
	{
		SendMessage(CurrentHwnd, CB_LIMITTEXT, (WPARAM)LimitLength, 0);
	}

	/*
	设置组合框选择的列表项
	ListIndex: 列表项的序号，如果为-1则不选择任何列表项
	如果执行成功则返回选取的列表项的序号，如果执行失败则返回CB_ERR (-1)
	*/
	int SetSelItem(int ListIndex)
	{
		return (int)SendMessage(CurrentHwnd, CB_SETCURSEL, ListIndex, 0);
	}

	/*
	设置组合框选择的列表项
	ListText: 需要被选择的列表项的文本，系统会检测文本的开头是否包含查找的文本
	StartItem: 开始搜索的列表项。如果是-1则从头开始查找
	如果执行成功则返回选取的列表项的序号，如果执行失败则返回CB_ERR (-1)
	*/
	int SetSelItem(char* ListText, int StartItem = -1)
	{
		return (int)SendMessage(CurrentHwnd, CB_SELECTSTRING, (WPARAM)StartItem, (LPARAM)ListText);
	}

	/*
	设置组合框的下拉列表的宽度
	NewWidth: 下拉列表新的宽度
	如果执行成功则返回下拉列表新的宽度，如果执行失败则返回CB_ERR (-1)
	*/
	int SetDroppedWidth(int NewWidth)
	{
		return (int)SendMessage(CurrentHwnd, CB_SETDROPPEDWIDTH, (WPARAM)NewWidth, 0);
	}

	/*
	设置组合框中文本框的选取范围
	SelStart: 文本选取的开头。如果为-1则取消选取文本
	SelEnd: 文本选取的结尾。如果为-1则文本选取的范围为SelStart到末尾
	如果执行成功则返回TRUE，如果执行失败则返回CB_ERR (-1)
	*/
	bool SetSelText(int SelStart, int SelEnd)
	{
		return (SendMessage(CurrentHwnd, CB_SETEDITSEL, 0, MAKELPARAM(SelStart, SelEnd)) == TRUE);
	}

	/*
	设置组合框的下拉列表中每一个列表项的高度
	NewHeight: 每个列表项新的高度
	如果执行成功则返回列表项新的高度，如果执行失败则返回CB_ERR (-1)
	*/
	int SetItemHeight(int NewHeight)
	{
		return (int)SendMessage(CurrentHwnd, CB_SETITEMHEIGHT, 0, NewHeight);
	}

	/*
	设置组合框的下拉列表每次至少显示的列表项数
	ItemCount: 至少显示的列表项数
	如果执行成功则返回true
	*/
	bool SetMinimumVisibleItems(int ItemCount)
	{
		return (SendMessage(CurrentHwnd, CB_SETMINVISIBLE, (WPARAM)ItemCount, 0) == TRUE);
	}

	/*
	使指定的列表项能够被显示到组合框的下拉列表中
	ListIndex: 需要显示在列表框中的项目
	如果成功则返回true
	*/
	bool ScrollToItem(int ListIndex)
	{
		return (SendMessage(CurrentHwnd, CB_SETTOPINDEX, (WPARAM)ListIndex, 0) == 0);
	}

	/*
	显示或者隐藏组合框的下拉列表
	bShow: 是否显示组合框的下拉列表
	*/
	void ShowDropDownList(bool bShow)
	{
		SendMessage(CurrentHwnd, CB_SHOWDROPDOWN, (WPARAM)bShow, 0);
	}
};

//============================================================================
class MyListBox : public MyControls
{
public:
	DWORD		VerticalScrollBar;				//垂直滚动条
	bool		MultiSelect;					//允许多选
	bool		MultiColumn;					//允许多列
	bool		ClientEdgeBorder;				//立体边框
	bool		BlackBorder;					//黑色边框
	bool		AutoSort;						//自动排列

	//创建列表框控件
	HWND Create()
	{
		LONG lStyle = WS_CHILD | LBS_HASSTRINGS | LBS_NOINTEGRALHEIGHT | LBS_NOTIFY;		//初始化控件样式

		lStyle |= VerticalScrollBar;
		OrCalc(MultiSelect, &lStyle, LBS_EXTENDEDSEL);
		OrCalc(MultiColumn, &lStyle, LBS_MULTICOLUMN);
		OrCalc(BlackBorder, &lStyle, WS_BORDER);
		OrCalc(AutoSort, &lStyle, LBS_SORT);

		CurrentHwnd = CreateWindowEx(ClientEdgeBorder ? WS_EX_CLIENTEDGE : 0, "MyListBox", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	添加指定的文本到列表框中的指定位置
	strAdd: 需要添加的文本
	ListIndex: 文本添加到的位置。如果是-1 则添加到列表框的末尾
	如果添加成功则返回对应的列表序号，如果失败则返回LB_ERR (-1)
	*/
	int AddItem(char* strAdd, int ListIndex = -1)
	{
		return (int)SendMessage(CurrentHwnd, LB_INSERTSTRING, (WPARAM)ListIndex, (LPARAM)strAdd);
	}

	/*
	添加指定的数值到列表框中的指定位置
	strAdd: 需要添加的数值
	ListIndex: 文本添加到的位置。如果是-1 则添加到列表框的末尾
	如果添加成功则返回对应的列表序号，如果失败则返回LB_ERR (-1)
	*/
	int AddItem(int strAdd, int ListIndex = -1)
	{
		char Buffer[255];
		itoa(strAdd, Buffer, 10);
		return (int)SendMessage(CurrentHwnd, LB_INSERTSTRING, (WPARAM)ListIndex, (LPARAM)Buffer);
	}

	/*
	删除指定序号的列表项
	ListIndex: 指定的列表项序号
	如果执行成功则返回列表中剩余的项目数，如果失败则返回LB_ERR (-1)
	*/
	int RemoveItem(int ListIndex)
	{
		return (int)SendMessage(CurrentHwnd, LB_DELETESTRING, (WPARAM)ListIndex, 0);
	}

	/* 清空所有列表项 */
	void Clear()
	{
		SendMessage(CurrentHwnd, LB_RESETCONTENT, 0, 0);
	}

	/*
	在列表框中列出指定目录的文件
	FilePath: 文件路径
	FileType: 筛选的文件类型，为DDL_*的常数
	如果执行成功则返回最后一个添加的列表项的序号，如果失败则返回LB_ERR (-1)，如果列表空间不足则返回LB_ERRSPACE (-2)
	*/
	int ListDirFiles(LPCTSTR FilePath, DWORD FileType = 0)
	{
		return (int)SendMessage(CurrentHwnd, LB_DIR, (WPARAM)FileType, (LPARAM)FilePath);
	}

	/*
	在列表框中查找有指定文本的列表项
	StrFind: 需要查找的文本
	StartIndex: 从哪个列表项开始查找。如果为-1则从头开始查找
	FullMatch: 是否需要全字匹配。如果不是，则检测文本的开头是否包含查找的文本
	如果执行成功则返回找到的列表项的序号，如果失败则返回LB_ERR (-1)
	*/
	int FindItem(char* strFind, int StartIndex = -1, bool FullMatch = true)
	{
		return (int)SendMessage(CurrentHwnd, FullMatch ? LB_FINDSTRINGEXACT : LB_FINDSTRING,
			(WPARAM)StartIndex, (LPARAM)strFind);
	}

	/* 获取列表框中的项目数量 */
	int GetListCount()
	{
		return (int)SendMessage(CurrentHwnd, LB_GETCOUNT, 0, 0);
	}

	/*
	获取列表框中当前选取的列表项
	返回值为当前选取的列表项序号。如果没有列表项被选取则返回LB_ERR (-1)
	*/
	int GetSelItem()
	{
		return (int)SendMessage(CurrentHwnd, LB_GETCURSEL, 0, 0);
	}

	/* 获取列表框中其中一个列表项的高度 */
	int GetItemHeight()
	{
		return (int)SendMessage(CurrentHwnd, LB_GETITEMHEIGHT, 0, 0);
	}

	/*
	获取列表框中指定列表项的文本
	ListIndex: 指定的列表项序号
	*/
	char* GetListText(int ListIndex)
	{
		int Length = (int)SendMessage(CurrentHwnd, LB_GETTEXTLEN, ListIndex, 0);
		char* tmp = new char[Length];
		if (SendMessage(CurrentHwnd, LB_GETTEXT, ListIndex, (LPARAM)tmp) != LB_ERR)
		{
			return tmp;
		}
		else
		{
			return "";
		}
	}

	/*
	设置列表框选择的列表项
	ListIndex: 列表项的序号，如果为-1则不选择任何列表项
	bSelect: 是否选择此列表项。若为FALSE，则取消选取该列表项
	如果执行成功则返回选取的列表项的序号，如果执行失败则返回LB_ERR (-1)
	*/
	int SetSelItem(int ListIndex, bool bSelect = TRUE)
	{
		return (int)SendMessage(CurrentHwnd, LB_SETCURSEL, (WPARAM)ListIndex, 0);
	}

	/*
	设置列表框的下拉列表中每一个列表项的高度
	NewHeight: 每个列表项新的高度
	如果执行成功则返回列表项新的高度，如果执行失败则返回LB_ERR (-1)
	*/
	int SetItemHeight(int NewHeight)
	{
		return (int)SendMessage(CurrentHwnd, LB_SETITEMHEIGHT, 0, NewHeight);
	}

	/*
	使指定的列表项能够被显示到列表框中
	ListIndex: 需要显示在列表框中的项目
	如果成功则返回true
	*/
	bool ScrollToItem(int ListIndex)
	{
		return (SendMessage(CurrentHwnd, LB_SETCARETINDEX, (WPARAM)ListIndex, FALSE) == 0);
	}

	/*
	获得当列表项多项选择时的首个列表项（仅适用于多选列表框）
	如果执行失败则返回LB_ERR (-1)
	*/
	int GetFirstItem()
	{
		return (int)(SendMessage(CurrentHwnd, LB_GETANCHORINDEX, 0, 0));
	}

	/*
	获得当前多选的列表项数（仅适用于多选列表框）
	如果执行失败则返回LB_ERR (-1)
	*/
	int GetSelCount()
	{
		return (int)(SendMessage(CurrentHwnd, LB_GETSELCOUNT, 0, 0));
	}

	/*
	获取列表框当前多选的所有列表项（仅适用于多选列表框）
	intBuffer：指向一个指针数组的指针。其定义一般为 int *intBuffer = new int[列表项数];
	如果成功则往intBuffer中写入所有的列表项序号，如果执行失败则返回LB_ERR (-1)
	*/
	int GetMultiSelItems(int* intBuffer)
	{
		int SelCount = (int)(SendMessage(CurrentHwnd, LB_GETSELCOUNT, 0, 0));					//列表项数
		int *Buffer = (int*)GlobalAlloc(GPTR, SelCount * sizeof(int));							//为缓冲区申请内存空间
		int rtn = SendMessage(CurrentHwnd, LB_GETSELITEMS, (WPARAM)SelCount, (LPARAM)Buffer);	//读取选择的列表项到缓冲区
		memcpy(intBuffer, Buffer, SelCount * sizeof(int));										//将缓冲区的内存拷贝到目标内存
		return (rtn == LB_ERR) ? LB_ERR : SelCount;														//判断函数是否执行成功
	}

	/*
	获取列表框可视的第一个列表项的序号
	如果函数执行成功则返回列表框的第一个可视的列表项序号，如果失败则返回LB_ERR (-1)
	*/
	int GetFirstVisibleItem()
	{
		return (int)SendMessage(CurrentHwnd, LB_GETTOPINDEX, 0, 0);
	}

	/*
	获取列表框指定坐标对应的列表项
	X、Y: 指定的坐标
	返回对应的列表项的序号
	*/
	int ItemFromPoint(int X, int Y)
	{
		return LOWORD(SendMessage(CurrentHwnd, LB_ITEMFROMPOINT, 0, MAKELPARAM(X, Y)));
	}

	/*
	多项选择指定的列表项（仅适用于多选列表框）
	ListIndexFrom: 选择的列表项的范围的首个
	ListIndexTo: 选择的列表项的范围的最后一个
	bSelect: 若为TRUE则选取指定范围内的列表项，若为FALSE则不选取指定范围内的列表项
	如果有错误发生则返回false
	*/
	bool SetSelItemRange(int ListIndexFrom, int ListIndexTo, bool bSelect = TRUE)
	{
		return (SendMessage(CurrentHwnd, LB_SELITEMRANGE, (WPARAM)bSelect,
			MAKELPARAM(ListIndexFrom, ListIndexTo)) != CB_ERR);
	}

	/* 设置列表框的每一列的宽度（仅适用于多列列表框） */
	void SetColumnWidth(int NewWidth)
	{
		SendMessage(CurrentHwnd, LB_SETCOLUMNWIDTH, (WPARAM)NewWidth, 0);
	}

	/* 让列表框获得焦点 */
	void SetFocus()
	{
		::SetFocus(CurrentHwnd);
	}
};

//============================================================================
class MyHScroll : public MyControls
{
public:
	int			Min;							//最小值
	int			Max;							//最大值
	int			SmallChange;					//最小更改值
	int			LargeChange;					//最大更改值

	//创建滚动条控件
	HWND Create()
	{
		LONG lStyle = WS_CHILD | SBS_HORZ | SBS_LEFTALIGN;

		CurrentHwnd = CreateWindowEx(0, "MyScrollBar", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		//设置滚动条的范围
		SetScrollRange(CurrentHwnd, SB_CTL, Min, Max, TRUE);

		return CurrentHwnd;
	}

	//获得滚动条的值
	int GetValue()
	{
		return GetScrollPos(CurrentHwnd, SB_CTL);
	}

	/*
	设置滚动条的值
	如果执行成功则返回滚动条之前的值，如果执行失败则返回0
	*/
	int SetValue(int Value)
	{
		return SetScrollPos(CurrentHwnd, SB_CTL, Value, TRUE);
	}

	/*
	获得滚动条范围
	lpMin, lpMax 分别为指向用来接收滚动条最小值和最大值的整数的指针
	如果执行失败则返回0
	*/
	int GetRange(int* lpMin, int* lpMax)
	{
		return GetScrollRange(CurrentHwnd, SB_CTL, lpMin, lpMax);
	}

	/*
	设置滚动条范围
	MinValue, MaxValue 分别为滚动条新的最小值和最大值
	如果执行失败则返回0
	*/
	int SetRange(int MinValue, int MaxValue)
	{
		return SetScrollRange(CurrentHwnd, SB_CTL, MinValue, MaxValue, TRUE);
	}

};

class MyVScroll : public MyHScroll
{
public:
	HWND Create()
	{
		LONG lStyle = WS_CHILD | SBS_VERT;

		CurrentHwnd = CreateWindowEx(0, "MyScrollBar", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		//设置滚动条的范围
		SetScrollRange(CurrentHwnd, SB_CTL, Min, Max, TRUE);

		return CurrentHwnd;
	}
};

//============================================================================
class MyUpDown : public MyControls
{
public:
	int			Min;							//最小值
	int			Max;							//最大值
	int			Accel;							//增加速度
	bool		HorzStyle;						//是否为水平样式

	//创建调节按钮控件
	HWND Create()
	{
		LONG lStyle = WS_CHILD | (HorzStyle ? UDS_HORZ : 0);

		CurrentHwnd = CreateWindowEx(0, "MyUpDown", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		//设置最小值和最大值
		PostMessage(CurrentHwnd, UDM_SETRANGE32, (WPARAM)Min, (LPARAM)Max);

		//设置调节速度
		UDACCEL uda;
		uda.nSec = 1;
		uda.nInc = Accel;
		SendMessage(CurrentHwnd, UDM_SETACCEL, 1, (LPARAM)&uda);

		return CurrentHwnd;
	}

	/* 获取调节按钮的调节速度 */
	int GetAccel()
	{
		UDACCEL uda;
		SendMessage(CurrentHwnd, UDM_GETACCEL, 1, (LPARAM)&uda);
		return uda.nInc;
	}

	/*
	设置调节按钮的调节速度
	Acceleration: 新的调节速度
	如果执行成功则返回TRUE
	*/
	bool SetAccel(int Acceleration)
	{
		UDACCEL uda;
		uda.nSec = 1;
		uda.nInc = Acceleration;
		return (bool)SendMessage(CurrentHwnd, UDM_SETACCEL, 1, (LPARAM)&uda);
	}

	/* 获取调节按钮的值 */
	int GetPos()
	{
		return (int)SendMessage(CurrentHwnd, UDM_GETPOS32, 0, 0);
	}

	/*
	设置调节按钮的值
	NewPos: 新的值
	返回调节按钮之前的值
	*/
	int SetPos(int NewPos)
	{
		return (int)SendMessage(CurrentHwnd, UDM_SETPOS32, 0, (LPARAM)NewPos);
	}

	/*
	获取调节按钮的范围
	lpMin, lpMax 分别为指向用来接收调节按钮最小值和最大值的整数的指针
	*/
	void GetRange(int* lpMin, int* lpMax)
	{
		SendMessage(CurrentHwnd, UDM_GETRANGE32, (WPARAM)lpMin, (LPARAM)lpMax);
	}

	/*
	设置调节按钮的范围
	Min, Max 分别为调节按钮新的最小值和最大值
	*/
	void SetRange(int Min, int Max)
	{
		PostMessage(CurrentHwnd, UDM_SETRANGE32, (WPARAM)Min, (LPARAM)Max);
	}
};

//============================================================================
class MyProgressBar : public MyControls
{
public:
	int			Min;							//最小值
	int			Max;							//最大值
	bool		Smooth;							//是否平滑样式
	bool		VertStyle;						//是否为垂直样式
	COLORREF	BarColor;						//滑块颜色
	COLORREF	BackColor;						//背景颜色

	//创建进度条控件
	HWND Create()
	{
		LONG lStyle = WS_CHILD | (Smooth ? PBS_SMOOTH : 0) | (VertStyle ? PBS_VERTICAL : 0) | PBS_MARQUEE;

		CurrentHwnd = CreateWindowEx(0, "MyProgressBar", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		//设置最小值和最大值
		PostMessage(CurrentHwnd, PBM_SETRANGE32, (WPARAM)Min, (LPARAM)Max);

		//设置滑块颜色和背景颜色
		PostMessage(CurrentHwnd, PBM_SETBARCOLOR, 0, (LPARAM)BarColor);
		PostMessage(CurrentHwnd, PBM_SETBKCOLOR, 0, (LPARAM)BackColor);

		return CurrentHwnd;
	}

	/*
	令进度条增加指定的值
	ValueAdd: 进度条增加的值
	返回进度条之前的值
	*/
	int IncreaseValue(int ValueAdd)
	{
		return (int)PostMessage(CurrentHwnd, PBM_DELTAPOS, (WPARAM)ValueAdd, 0);
	}

	/* 获取滑块颜色 */
	COLORREF GetBarColor()
	{
		return (COLORREF)PostMessage(CurrentHwnd, PBM_GETBARCOLOR, 0, 0);
	}

	/*
	设置滑块的颜色
	NewColor: 新的滑块颜色。
	返回进度条之前的滑块颜色
	*/
	COLORREF SetBarColor(COLORREF NewColor)
	{
		return (COLORREF)PostMessage(CurrentHwnd, PBM_SETBARCOLOR, 0, (LPARAM)NewColor);
	}

	/* 获取背景颜色 */
	COLORREF GetBackColor()
	{
		return (COLORREF)PostMessage(CurrentHwnd, PBM_GETBKCOLOR, 0, 0);
	}

	/*
	设置背景颜色
	NewColor: 新的背景颜色。若为CLR_DEFAULT (0xFF000000L)，则使用系统默认的颜色。
	返回进度条之前的背景颜色
	*/
	COLORREF SetBackColor(COLORREF NewColor)
	{
		return (COLORREF)PostMessage(CurrentHwnd, PBM_SETBKCOLOR, 0, (LPARAM)NewColor);
	}

	/* 获取进度条的值 */
	int GetValue()
	{
		return (int)SendMessage(CurrentHwnd, PBM_GETPOS, 0, 0);
	}

	/*
	设置进度条的值
	NewValue: 新的值。若该值超出进度条的范围系统则会设置为最接近的有效值。
	返回之前的值
	*/
	int SetValue(int NewValue)
	{
		return (int)SendMessage(CurrentHwnd, PBM_SETPOS, (WPARAM)NewValue, 0);
	}

	/* 获取进度条的最小值 */
	int GetMin()
	{
		return (int)SendMessage(CurrentHwnd, PBM_GETRANGE, (WPARAM)TRUE, 0);
	}

	/* 获取进度条的最大值 */
	int GetMax()
	{
		return (int)SendMessage(CurrentHwnd, PBM_GETRANGE, (WPARAM)FALSE, 0);
	}

	/*
	获取进度条的范围
	lpMin, lpMax 分别为指向用来接收进度条最小值和最大值的整数的指针
	*/
	void GetRange(int* lpMin, int* lpMax)
	{
		PBRANGE pr;
		SendMessage(CurrentHwnd, PBM_GETRANGE, 0, (LPARAM)&pr);
		*lpMin = pr.iLow;
		*lpMax = pr.iHigh;
	}

	/*
	设置进度条范围
	MinValue, MaxValue 分别为进度条新的最大值和最小值
	返回一个LOWORD中装有之前的最小值，HIWORD中装有之前的最大值的整数
	*/
	DWORD SetRange(int MinValue, int MaxValue)
	{
		return (DWORD)SendMessage(CurrentHwnd, PBM_SETRANGE32, (WPARAM)MinValue, (LPARAM)MaxValue);
	}
};

//============================================================================
class MySlider : public MyControls
{
public:
	bool		Direction;						//方向。若为true则为水平，若为false则为垂直
	int			MarkPosition;					//刻度位置。0-5分别为左边、右边、上方、下方、都有、无刻度
	bool		NoBar;							//是否不显示滑块
	int			TooltipPos;						//数字标签位置0-4分别为左边、右边、上方、下方、无数字标签
	int			TickFreq;						//刻度间隔
	int			Min;							//最小值
	int			Max;							//最大值
	int			SmallChange;					//慢速更改步长
	int			LargeChange;					//快速更改步长
	bool		BlackBorder;					//黑色边框

	//创建滑块控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD | TBS_AUTOTICKS;
		OrCalc(!Direction, &lStyle, TBS_VERT | TBS_DOWNISLEFT);
		OrCalc(NoBar, &lStyle, TBS_NOTHUMB);
		OrCalc(BlackBorder, &lStyle, WS_BORDER);

		//判断有没有数字标签
		OrCalc(TooltipPos != 4, &lStyle, TBS_TOOLTIPS);

		//判断刻度位置
		switch (MarkPosition)
		{
		case 0:
		case 2:
			lStyle |= TBS_LEFT;
			break;

		case 4:
			lStyle |= TBS_BOTH;
			break;

		case 5:
			lStyle |= TBS_NOTICKS;
			break;
		}

		CurrentHwnd = CreateWindowEx(0, "MySlider", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置滑块的数字标签位置
		SetToolTipPos(TooltipPos);

		//设置滑块的刻度间隔
		SetTickFreq(TickFreq);

		//设置滑块的最小值和最大值
		SetRange(Max, Min);

		//设置慢速更改步长和快速更改步长
		SetSmallChange(SmallChange);
		SetLargeChange(LargeChange);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	设置滑块的数字标签位置
	NewPos: 数字标签位置0-3分别为左边、右边、上方、下方
	*/
	void SetToolTipPos(int NewPos)
	{
		switch (NewPos)
		{
		case 0:
			SendMessage(CurrentHwnd, TBM_SETTIPSIDE, TBTS_LEFT, 0);
			break;

		case 1:
			SendMessage(CurrentHwnd, TBM_SETTIPSIDE, TBTS_RIGHT, 0);
			break;

		case 2:
			SendMessage(CurrentHwnd, TBM_SETTIPSIDE, TBTS_TOP, 0);
			break;

		case 3:
			SendMessage(CurrentHwnd, TBM_SETTIPSIDE, TBTS_BOTTOM, 0);
			break;
		}
	}

	/*
	设置滑块的刻度间隔
	NewTickFreq: 新的刻度间隔
	*/
	void SetTickFreq(int NewTickFreq)
	{
		SendMessage(CurrentHwnd, TBM_SETTICFREQ, (WPARAM)NewTickFreq, 0);
	}

	/*
	设置滑块的最大值
	NewMax: 新的最大值
	*/
	void SetMax(int NewMax)
	{
		SendMessage(CurrentHwnd, TBM_SETRANGEMAX, TRUE, (LPARAM)NewMax);
	}

	/*
	设置滑块的最小值
	NewMin: 新的最小值
	*/
	void SetMin(int NewMin)
	{
		SendMessage(CurrentHwnd, TBM_SETRANGEMIN, TRUE, (LPARAM)NewMin);
	}

	/*
	设置滑块的范围
	NewMax: 新的最大值
	NewMin: 新的最小值
	*/
	void SetRange(int NewMax, int NewMin)
	{
		SendMessage(CurrentHwnd, TBM_SETRANGE, TRUE, MAKELPARAM(NewMin, NewMax));
	}

	/* 获取滑块的最大值 */
	int GetMax()
	{
		return (int)SendMessage(CurrentHwnd, TBM_GETRANGEMAX, 0, 0);
	}

	/* 获取滑块的最小值 */
	int GetMin()
	{
		return (int)SendMessage(CurrentHwnd, TBM_GETRANGEMIN, 0, 0);
	}

	/*
	设置滑块的慢速更改步长
	NewSmallChange: 新的慢速更改步长
	*/
	void SetSmallChange(int NewSmallChange)
	{
		SendMessage(CurrentHwnd, TBM_SETLINESIZE, TRUE, (LPARAM)NewSmallChange);
	}

	/*
	设置滑块的快速更改步长
	NewLargeChange: 新的快速更改步长
	*/
	void SetLargeChange(int NewLargeChange)
	{
		SendMessage(CurrentHwnd, TBM_SETPAGESIZE, TRUE, (LPARAM)NewLargeChange);
	}

	/*
	设置滑块位置
	NewPos: 新的滑块位置
	*/
	void SetPos(int NewPos)
	{
		SendMessage(CurrentHwnd, TBM_SETPOS, TRUE, (LPARAM)NewPos);
	}

	/* 获取滑块位置 */
	int GetPos()
	{
		return (int)SendMessage(CurrentHwnd, TBM_GETPOS, 0, 0);
	}
};

//============================================================================
class MyHotkey : public MyControls
{
public:
	//创建热键控件
	HWND Create()
	{
		CurrentHwnd = CreateWindowEx(0, "MyHotkey", "", WS_VISIBLE | WS_CHILD,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	设置热键控件的热键
	KeyCode: 指定的键盘按键
	KeyModifier: 指定的键盘功能键。Shift = 1; Ctrl = 2; Alt = 4
	*/
	void SetHotkey(BYTE KeyCode, BYTE KeyModifier)
	{
		SendMessage(CurrentHwnd, HKM_SETHOTKEY, MAKEWPARAM(MAKEWORD(KeyCode, KeyModifier), 0), 0);
	}

	/*
	获取热键控件的热键
	lpKeyCode: 用来接收键盘按键的变量的指针
	lpKeyModifier: 用来接收键盘功能键的变量的指针
	*/
	void GetHotKey(LPBYTE KeyCode, LPBYTE KeyModifier)
	{
		LRESULT rtn = SendMessage(CurrentHwnd, HKM_GETHOTKEY, 0, 0);
		*KeyCode = LOBYTE(LOWORD(rtn));
		*KeyModifier = HIBYTE(LOWORD(rtn));
	}
};

//============================================================================
class MyListView : public MyControls
{
public:
	int			Style;							//样式。0-3分别为图标、列表、报告、小图标
	int			Sort;							//自动排序模式。0-2分别为递增、递减、不排序
	int			Align;							//自动对齐。0-2分别为左对齐、顶端对齐、自动
	bool		EditableLabel;					//标签是否可编辑
	bool		MultiSelectItems;				//是否可以多选
	bool		BlackBorder;					//黑色边框

	//创建列表视图控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;
		switch (Style)
		{
		case 0:
			lStyle |= LVS_ICON;
			break;

		case 1:
			lStyle |= LVS_LIST;
			break;

		case 2:
			lStyle |= LVS_REPORT;
			break;

		case 3:
			lStyle |= LVS_SMALLICON;
			break;
		}
		OrCalc(Sort == 0, &lStyle, LVS_SORTASCENDING);
		OrCalc(Sort == 1, &lStyle, LVS_SORTDESCENDING);
		OrCalc(Align == 0, &lStyle, LVS_ALIGNLEFT);
		OrCalc(Align == 2, &lStyle, LVS_AUTOARRANGE);
		OrCalc(EditableLabel, &lStyle, LVS_EDITLABELS);
		OrCalc(!MultiSelectItems, &lStyle, LVS_SINGLESEL);
		OrCalc(BlackBorder, &lStyle, WS_BORDER);

		CurrentHwnd = CreateWindowEx(0, "MyListView", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	添加列表头到列表视图中
	ColumnText: 列表头文本
	Width: 列表头宽度
	Index: 列表头序号
	Position: 列表头文本对齐位置，可以为LVCFMT_*的常数
	*/
	void AddColumn(char* ColumnText, int Width, int Index = 0, int Position = LVCFMT_LEFT)
	{
		LVCOLUMN lvCol;
		lvCol.mask = LVCF_WIDTH | LVCF_TEXT | LVCF_FMT;
		lvCol.fmt = Position;
		lvCol.cx = Width;
		lvCol.pszText = ColumnText;
		lvCol.cchTextMax = 255;
		SendMessage(CurrentHwnd, LVM_INSERTCOLUMN, Index, (LPARAM)&lvCol);
	}

	/*
	添加列表项到列表视图中
	ItemText: 新添加的列表项的文本
	若执行成功则返回新加的列表项的序号，若失败则返回-1
	*/
	int AddItem(char* ItemText)
	{
		LVITEM lvItem = { 0 };
		lvItem.iItem = GetItemCount();
		lvItem.mask = LVIF_TEXT;
		lvItem.pszText = ItemText;
		lvItem.cchTextMax = 255;
		return SendMessage(CurrentHwnd, LVM_INSERTITEM, 0, (LPARAM)&lvItem);
	}

	/*
	更改列表项的文本
	ItemText: 新的列表文本
	ListIndex: 需要更改的列表项的序号
	SubItemIndex: 需要更改的子列表项的序号
	*/
	void SetItemText(char* ItemText, int ListIndex, int SubItemIndex = 0)
	{
		LVITEM lvItem = { 0 };
		lvItem.mask = LVIF_TEXT;
		lvItem.pszText = ItemText;
		lvItem.cchTextMax = 255;
		lvItem.iSubItem = SubItemIndex;
		SendMessage(CurrentHwnd, LVM_SETITEMTEXT, ListIndex, (LPARAM)&lvItem);
	}

	/*
	获取列表项文本
	Index: 指定列表项的序号
	SubItemIndex: 指定子列表项的序号
	*/
	char* GetItemText(int ListIndex, int SubItemIndex = 0)
	{
		char* tmp = new char[255];
		LVITEM lvItem = { 0 };
		lvItem.mask = LVIF_TEXT;
		lvItem.cchTextMax = 255;
		lvItem.pszText = tmp;
		lvItem.iItem = ListIndex;
		lvItem.iSubItem = SubItemIndex;
		SendMessage(CurrentHwnd, LVM_GETITEM, 0, (LPARAM)&lvItem);
		return lvItem.pszText;
	}

	/* 获取列表项数量 */
	int GetItemCount()
	{
		return SendMessage(CurrentHwnd, LVM_GETITEMCOUNT, 0, 0);
	}

	/*
	获取指定列表头的文本
	Index: 指定列表头的序号
	*/
	char* GetColumnText(int Index)
	{
		char* tmp = new char[255];
		LVCOLUMN lvItem = { 0 };
		lvItem.mask = LVCF_TEXT;
		lvItem.cchTextMax = 255;
		lvItem.pszText = tmp;
		SendMessage(CurrentHwnd, LVM_GETCOLUMN, (WPARAM)Index, (LPARAM)&lvItem);
		return lvItem.pszText;
	}

	/*
	获取指定列表头的宽度
	Index: 指定列表头的序号
	*/
	int GetColumnWidth(int Index)
	{
		LVCOLUMN lvItem = { 0 };
		lvItem.mask = LVCF_WIDTH;
		SendMessage(CurrentHwnd, LVM_GETCOLUMN, (WPARAM)Index, (LPARAM)&lvItem);
		return lvItem.cx;
	}

	/*
	设置指定列表头的文本
	Index: 指定列表头的序号
	NewText: 新的文本
	若执行成功则返回TRUE
	*/
	bool SetColumnText(int Index, char* NewText)
	{
		LVCOLUMN lvItem = { 0 };
		lvItem.mask = LVCF_TEXT;
		lvItem.pszText = NewText;
		return (bool)SendMessage(CurrentHwnd, LVM_SETCOLUMN, (WPARAM)Index, (LPARAM)&lvItem);
	}

	/*
	设置指定列表头的宽度
	Index: 列表头的序号
	NewWidth: 新的宽度
	若执行成功则返回TRUE
	*/
	bool SetColumnWidth(int Index, int NewWidth)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_SETCOLUMNWIDTH, Index, NewWidth);
	}

	/*
	进入标签编辑操作
	Index: 需要编辑标签的列表项序号。若为-1，则取消编辑
	若执行失败，返回值为0
	*/
	int EditLabel(int Index)
	{
		return SendMessage(CurrentHwnd, LVM_EDITLABEL, Index, 0);
	}

	/* 退出标签编辑操作 */
	void CancelEditLabel()
	{
		SendMessage(CurrentHwnd, LVM_CANCELEDITLABEL, 0, 0);
	}

	/* 删除所有的列表项 */
	void DeleteAllItems()
	{
		SendMessage(CurrentHwnd, LVM_DELETEALLITEMS, 0, 0);
	}

	/*
	删除指定的列表头
	Index: 列表头的序号
	*/
	bool DeleteColumn(int Index)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_DELETECOLUMN, Index, 0);
	}

	/*
	删除指定的列表项
	Index: 列表项的序号
	*/
	bool DeleteItem(int Index)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_DELETEITEM, Index, 0);
	}

	/*
	确保指定的列表项可视
	Index: 列表项的序号
	*/
	bool EnsureVisible(int Index)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_ENSUREVISIBLE, Index, TRUE);
	}

	/*
	按照指定的文本查找列表项
	ItemText: 指定列表项的文本
	AllMatch: 字符串是否必须完全相同。若为FALSE，若列表项开头包含查找的字符串也可接受
	StartIndex: 从指定的列表项开始查找。默认为从头开始查找
	返回找到的列表项的序号。如果失败则返回-1
	*/
	int FindItem(char* ItemText, bool FullMatch, int StartIndex = -1)
	{
		LVFINDINFO lvfi = { 0 };
		if (!FullMatch)
			lvfi.flags = LVFI_PARTIAL;
		lvfi.flags |= LVFI_STRING;
		lvfi.psz = ItemText;
		return SendMessage(CurrentHwnd, LVM_FINDITEM, StartIndex, (LPARAM)&lvfi);
	}

	/*
	按照指定的坐标查找列表项，仅适用于图标样式
	X: X坐标值
	Y: Y坐标值
	返回找到的列表项的序号。如果失败则返回-1
	*/
	int ItemFromXY(int X, int Y)
	{
		LVFINDINFO lvfi = { 0 };
		lvfi.flags = LVFI_NEARESTXY;
		lvfi.pt.x = X;
		lvfi.pt.y = Y;
		return SendMessage(CurrentHwnd, LVM_FINDITEM, (WPARAM)-1, (LPARAM)&lvfi);
	}

	/*
	设置列表视图的背景颜色
	Color: 新的颜色。若为CLR_NONE则使用系统默认的颜色
	若执行成功则返回TRUE
	*/
	bool SetBackColor(COLORREF Color)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_SETBKCOLOR, 0, (LPARAM)Color);
	}

	/*
	设置列表视图的文本背景颜色
	Color: 新的颜色。若为CLR_NONE则使用系统默认的颜色
	若执行成功则返回TRUE
	*/
	bool SetTextBackColor(COLORREF Color)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_SETTEXTBKCOLOR, 0, (LPARAM)Color);
	}

	/*
	设置列表视图的文本颜色
	Color: 新的颜色
	若执行成功则返回TRUE
	*/
	bool SetTextColor(COLORREF Color)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_SETTEXTCOLOR, 0, (LPARAM)Color);
	}

	/*
	获取列表视图的背景颜色
	返回背景颜色值
	*/
	COLORREF GetBackColor()
	{
		return (COLORREF)SendMessage(CurrentHwnd, LVM_GETBKCOLOR, 0, 0);
	}

	/*
	获取列表视图的文本背景颜色
	返回文本背景颜色值
	*/
	COLORREF GetTextBackColor()
	{
		return (COLORREF)SendMessage(CurrentHwnd, LVM_GETTEXTBKCOLOR, 0, 0);
	}

	/*
	获取列表视图的文本颜色
	返回文本颜色值
	*/
	COLORREF GetTextColor()
	{
		return (COLORREF)SendMessage(CurrentHwnd, LVM_GETTEXTCOLOR, 0, 0);
	}

	/*
	滚动列表视图
	vScroll: 垂直滚动的距离
	hScroll: 水平滚动的距离
	若执行成功则返回TRUE
	*/
	bool Scroll(int vScroll, int hScroll = 0)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_SCROLL, (WPARAM)hScroll, (LPARAM)vScroll);
	}

	/*
	设置图标位置（仅适用于图标视图）
	ListIndex: 列表项序号
	X, Y: 图标的坐标
	*/
	void SetItemPosition(int ListIndex, int X, int Y)
	{
		POINT point;
		point.x = X;
		point.y = Y;
		SendMessage(CurrentHwnd, LVM_SETITEMPOSITION32, (WPARAM)ListIndex, (LPARAM)&point);
	}

	/*
	获取图标位置
	ListIndex: 列表项序号
	pX, pY: 指向用来接收图标坐标的指针
	返回值：若成功则返回TRUE
	*/
	bool GetItemPosotion(int ListIndex, long* pX, long* pY)
	{
		POINT point;
		bool rtn = (bool)SendMessage(CurrentHwnd, LVM_GETITEMPOSITION, ListIndex, (LPARAM)&point);
		*pX = point.x;
		*pY = point.y;
		return rtn;
	}

	/* 获取当前选择的列表头 */
	int GetSelectedColumn()
	{
		return SendMessage(CurrentHwnd, LVM_GETSELECTEDCOLUMN, 0, 0);
	}

	/*
	获取最顶端的列表项。仅适用于列表视图
	若执行成功则返回最顶端的列表项的序号，否则返回0
	*/
	int GetTopIndex()
	{
		return SendMessage(CurrentHwnd, LVM_GETTOPINDEX, 0, 0);
	}

	/*
	设置ListView的视图
	Style: 样式。0-3分别为图标、列表、报告、小图标
	返回值：若执行成功，则返回1；否则返回-1
	*/
	int SetStyle(int Style)
	{
		switch (Style)
		{
		case 0:
			SendMessage(CurrentHwnd, LVM_SETVIEW, (WPARAM)LV_VIEW_ICON, 0);
			break;

		case 1:
			SendMessage(CurrentHwnd, LVM_SETVIEW, (WPARAM)LV_VIEW_LIST, 0);
			break;

		case 2:
			SendMessage(CurrentHwnd, LVM_SETVIEW, (WPARAM)LV_VIEW_DETAILS, 0);
			break;

		case 3:
			SendMessage(CurrentHwnd, LVM_SETVIEW, (WPARAM)LV_VIEW_SMALLICON, 0);
			break;
		}
	}

	/* 获取当前选择的列表项序号。若没有选择列表项，则返回-1 */
	int GetSelectedItem()
	{
		return SendMessage(CurrentHwnd, LVM_GETNEXTITEM, (WPARAM)-1, (LPARAM)LVNI_SELECTED);
	}
};

//============================================================================
class MyTreeView : public MyControls
{
public:
	bool		EditableLabels;					//标签是否可编辑
	bool		HasButtons;						//是否显示节点按钮
	bool		RootHasButtons;					//根节点是否显示按钮
	bool		HasLines;						//是否显示树线
	bool		NoHscroll;						//是否禁止水平滚动
	bool		NoVHscroll;						//是否禁止水平和垂直滚动
	bool		ShowSelAlways;					//失焦时是否显示选择项
	bool		HotTracking;					//是否实时选取
	bool		CheckBoxes;						//是否有多选框
	bool		BlackBorder;					//是否有黑色边框

	//创建树视图控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;
		OrCalc(EditableLabels, &lStyle, TVS_EDITLABELS);
		OrCalc(HasButtons, &lStyle, TVS_HASBUTTONS);
		OrCalc(RootHasButtons, &lStyle, TVS_LINESATROOT);
		OrCalc(HasLines, &lStyle, TVS_HASLINES);
		OrCalc(NoHscroll, &lStyle, TVS_NOHSCROLL);
		OrCalc(NoVHscroll, &lStyle, TVS_NOSCROLL);
		OrCalc(ShowSelAlways, &lStyle, TVS_SHOWSELALWAYS);
		OrCalc(HotTracking, &lStyle, TVS_TRACKSELECT);
		OrCalc(CheckBoxes, &lStyle, TVS_CHECKBOXES);
		OrCalc(BlackBorder, &lStyle, WS_BORDER);

		CurrentHwnd = CreateWindowEx(0, "MyTreeView", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	添加项目到树视图中
	ItemText: 项目的文本
	Parent: 父节点的句柄
	返回值: 创建的项目的句柄
	*/
	HTREEITEM AddItem(char* ItemText, HTREEITEM Parent = 0)
	{
		TVINSERTSTRUCT ti = { 0 };
		ti.hInsertAfter = TVI_LAST;
		ti.hParent = Parent;
		ti.item.mask = TVIF_TEXT;
		ti.item.pszText = ItemText;
		ti.item.cchTextMax = 255;

		return (HTREEITEM)SendMessage(CurrentHwnd, TVM_INSERTITEM, 0, (LPARAM)&ti);
	}

	/*
	删除指定的项目
	Item: 需要删除项目的句柄。若需要删除所有的项目，这个参数设置为NULL
	返回值: 若删除成功则返回TRUE，否则返回FALSE
	*/
	bool RemoveItem(HTREEITEM Item)
	{
		return (bool)SendMessage(CurrentHwnd, TVM_DELETEITEM, 0, (LPARAM)Item);
	}

	/*
	确保指定的项目可视
	Item: 指定的项目的句柄
	*/
	void EnsureVisible(HTREEITEM Item)
	{
		SendMessage(CurrentHwnd, TVM_ENSUREVISIBLE, 0, (LPARAM)Item);
	}

	/*
	展开或者收缩树状图
	Item: 需要被展开或者收缩的列表项
	Mode: 展开或者收缩。1: 收缩；2: 展开；3: 切换展开或者收缩
	*/
	bool ExpandItems(HTREEITEM Item, int Mode)
	{
		return (bool)(SendMessage(CurrentHwnd, TVM_EXPAND, (WPARAM)Mode, (LPARAM)Item) != 0);
	}

	/*
	开始编辑文本
	Item: 需要编辑文本的项目的句柄
	*/
	bool EditLabel(HTREEITEM Item)
	{
		return (bool)(SendMessage(CurrentHwnd, TVM_EDITLABEL, 0, (LPARAM)Item) != 0);
	}

	/*
	取消编辑文本
	SaveChanges: 是否保存对项目的修改
	返回值: 若执行成功返回TRUE，否则返回FALSE
	*/
	bool EndEditLabel(bool SaveChanges)
	{
		return (bool)SendMessage(CurrentHwnd, TVM_ENDEDITLABELNOW, (WPARAM)SaveChanges, 0);
	}

	/*
	获取文本颜色
	返回值: 颜色值。若为-1，则代表使用系统默认颜色
	*/
	COLORREF GetTextColor()
	{
		return (COLORREF)SendMessage(CurrentHwnd, TVM_GETTEXTCOLOR, 0, 0);
	}

	/*
	获取线条颜色
	返回值: 颜色值。若为-1，则代表使用系统默认颜色
	*/
	COLORREF GetLineColor()
	{
		return (COLORREF)SendMessage(CurrentHwnd, TVM_GETLINECOLOR, 0, 0);
	}

	/*
	获取背景颜色
	返回值: 颜色值。若为-1，则代表使用系统默认颜色
	*/
	COLORREF GetBackColor()
	{
		return (COLORREF)SendMessage(CurrentHwnd, TVM_GETBKCOLOR, 0, 0);
	}

	/*
	设置文本颜色
	Color: 新的颜色。若为-1，则代表使用系统默认颜色
	返回值: 之前的颜色值。若为-1，则代表使用系统默认颜色
	*/
	COLORREF SetTextColor(COLORREF Color)
	{
		return (COLORREF)SendMessage(CurrentHwnd, TVM_SETTEXTCOLOR, 0, (LPARAM)Color);
	}

	/*
	设置背景颜色
	Color: 新的颜色。若为-1，则代表使用系统默认颜色
	返回值: 之前的颜色值。若为-1，则代表使用系统默认颜色
	*/
	COLORREF SetBackColor(COLORREF Color)
	{
		return (COLORREF)SendMessage(CurrentHwnd, TVM_SETBKCOLOR, 0, (LPARAM)Color);
	}

	/*
	设置线条颜色
	Color: 新的颜色。若为-1，则代表使用系统默认颜色
	返回值: 之前的颜色值。若为-1，则代表使用系统默认颜色
	*/
	COLORREF SetLineColor(COLORREF Color)
	{
		return (COLORREF)SendMessage(CurrentHwnd, TVM_SETLINECOLOR, 0, (LPARAM)Color);
	}

	/* 获取列表项的数量 */
	int GetListCount()
	{
		return (int)SendMessage(CurrentHwnd, TVM_GETCOUNT, 0, 0);
	}

	/* 获取列表项的高度 */
	int GetItemHeight()
	{
		return (int)SendMessage(CurrentHwnd, TVM_GETITEMHEIGHT, 0, 0);
	}

	/*
	设置列表项的高度
	Height: 新的高度。若为-1则使用系统默认高度
	返回值: 之前的列表项高度
	*/
	int SetItemHeight(int Height)
	{
		return (int)SendMessage(CurrentHwnd, TVM_SETITEMHEIGHT, (WPARAM)Height, 0);
	}

	/*
	获取指定列表项的文本
	Item: 列表项的句柄
	返回值: 指定列表项的文本
	*/
	char* GetItemText(HTREEITEM Item)
	{
		char* tmp = new char[255];
		TVITEM tvItem = { 0 };
		tvItem.mask = TVIF_TEXT;
		tvItem.cchTextMax = 255;
		tvItem.pszText = tmp;
		tvItem.hItem = Item;
		SendMessage(CurrentHwnd, TVM_GETITEM, 0, (LPARAM)&tvItem);
		return tvItem.pszText;
	}

	/*
	设置指定列表项的文本
	Item: 列表项的句柄
	NewText: 新的文本
	若执行成功则返回TRUE
	*/
	bool SetItemText(HTREEITEM Item, char* NewText)
	{
		TVITEM tvItem = { 0 };
		tvItem.mask = TVIF_TEXT;
		tvItem.cchTextMax = 255;
		tvItem.pszText = NewText;
		tvItem.hItem = Item;
		return (bool)SendMessage(CurrentHwnd, TVM_SETITEM, 0, (LPARAM)&tvItem);
	}

	/* 获取当前选择的项目句柄。若没有选择项目或选择的项目无效，则返回0 */
	HTREEITEM GetSelectedItem()
	{
		return (HTREEITEM)SendMessage(CurrentHwnd, TVM_GETNEXTITEM, TVGN_CARET, 0);
	}

	/*
	获取指定列表项的根节点句柄
	Item: 指定的列表项的句柄
	若没有选择项目或选择的项目无效，则返回0
	*/
	HTREEITEM GetParentItem(HTREEITEM Item)
	{
		return (HTREEITEM)SendMessage(CurrentHwnd, TVM_GETNEXTITEM, TVGN_PARENT, (LPARAM)Item);
	}

	/*
	获取指定列表项的上一个可视的列表项的句柄
	Item: 指定的列表项的句柄
	若没有选择项目或选择的项目无效，则返回0
	*/
	HTREEITEM GetPreviousItem(HTREEITEM Item)
	{
		return (HTREEITEM)SendMessage(CurrentHwnd, TVM_GETNEXTITEM, TVGN_LASTVISIBLE, (LPARAM)Item);
	}

	/*
	获取指定列表项的下一个可视的列表项的句柄
	Item: 指定的列表项的句柄
	若没有选择项目或选择的项目无效，则返回0
	*/
	HTREEITEM GetNextItem(HTREEITEM Item)
	{
		return (HTREEITEM)SendMessage(CurrentHwnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, (LPARAM)Item);
	}

	/*
	选择指定的列表项目
	Item: 指定的列表项的句柄。若要取消选择则为NULL
	若执行成功返回TRUE
	*/
	bool SelectItem(HTREEITEM Item)
	{
		return (bool)SendMessage(CurrentHwnd, TVM_SELECTITEM, TVGN_CARET, (LPARAM)Item);
	}

	/* 获取当前可视的项目的数量 */
	int GetVisibleCount()
	{
		return (int)SendMessage(CurrentHwnd, TVM_GETVISIBLECOUNT, 0, 0);
	}

	/*
	设置列表的缩进大小
	NewWidth: 新的缩进大小。若小于系统最小值则设置为系统最小值
	*/
	void SetIndent(int NewWidth)
	{
		SendMessage(CurrentHwnd, TVM_SETINDENT, (WPARAM)NewWidth, 0);
	}

	/* 获取列表的缩进大小 */
	int GetIndent()
	{
		return (int)SendMessage(CurrentHwnd, TVM_GETINDENT, 0, 0);
	}
};

//============================================================================
class MyTab : public MyControls
{
public:
	bool BottomTabs;							//选项卡在底部
	bool ButtonLike;							//按钮样式
	bool FlatButtons;							//扁平按钮
	bool FixedWidth;							//选项卡统一大小
	bool FocusOnButtons;						//按钮显示焦点
	bool ForceLabelLeft;						//文本左对齐
	bool HotTracking;							//实时选取
	bool MultiLine;								//多行选项卡
	bool ScrollOpposite;						//选项卡自动反向
	bool Vertical;								//垂直样式

	//创建选项卡控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;
		OrCalc(BottomTabs, &lStyle, TCS_BOTTOM);
		OrCalc(ButtonLike, &lStyle, TVS_HASBUTTONS);
		OrCalc(FlatButtons, &lStyle, TCS_FLATBUTTONS);
		OrCalc(FixedWidth, &lStyle, TCS_FIXEDWIDTH);
		OrCalc(FocusOnButtons, &lStyle, TCS_FOCUSONBUTTONDOWN);
		OrCalc(ForceLabelLeft, &lStyle, TCS_FORCELABELLEFT);
		OrCalc(HotTracking, &lStyle, TCS_HOTTRACK);
		OrCalc(MultiLine, &lStyle, TCS_MULTILINE);
		OrCalc(ScrollOpposite, &lStyle, TCS_SCROLLOPPOSITE);
		OrCalc(Vertical, &lStyle, TCS_VERTICAL);

		CurrentHwnd = CreateWindowEx(0, "MyTab", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);
		return CurrentHwnd;
	}

	//删除所有的选项卡
	bool DeleteAllItems()
	{
		return (bool)SendMessage(CurrentHwnd, TCM_DELETEALLITEMS, 0, 0);
	}

	/*
	删除指定的选项卡
	Index: 指定的选项卡序号
	*/
	bool DeleteItem(int Index)
	{
		return (bool)SendMessage(CurrentHwnd, TCM_DELETEITEM, Index, 0);
	}

	/*
	取消选择选项卡
	ResetAll: 若为FALSE，取消选取所有的标签；若为TRUE，取消选取当前选择标签之外的标签
	*/
	void DeselectAll(bool ExceptCurrSel)
	{
		SendMessage(CurrentHwnd, TCM_DESELECTALL, (WPARAM)ExceptCurrSel, 0);
	}

	/*
	获取当前选择的选项卡
	返回值: 当前选择的选项卡序号。若没有选择则返回-1
	*/
	int GetSel()
	{
		return (int)SendMessage(CurrentHwnd, TCM_GETCURSEL, 0, 0);
	}

	/*
	获取指定标签的文本
	返回值: 若执行成功返回TRUE，否则返回FALSE
	*/
	char* GetItemText(int Index)
	{
		char* tmp = new char[255];
		TCITEM tci = { 0 };
		tci.mask = TCIF_TEXT;
		tci.cchTextMax = 255;
		tci.pszText = tmp;
		SendMessage(CurrentHwnd, TCM_GETITEM, (WPARAM)Index, (LPARAM)&tci);
		return tci.pszText;
	}

	/*
	获取标签数量
	返回值: 标签数量
	*/
	int GetItemCount()
	{
		return (int)SendMessage(CurrentHwnd, TCM_GETITEMCOUNT, 0, 0);
	}

	/*
	获取一共多少行标签
	返回值: 标签的行数
	*/
	int GetRowCount()
	{
		return (int)SendMessage(CurrentHwnd, TCM_GETROWCOUNT, 0, 0);
	}

	/*
	高亮指定标签
	Index: 指定标签的序号
	HighLight: 是否高亮指定标签。若为TRUE，则高亮标签；若为FALSE，则恢复标签
	若执行成功，返回TRUE；否则返回FALSE
	*/
	bool HighLightItem(int Index, bool HighLight)
	{
		return (bool)SendMessage(CurrentHwnd, TCM_HIGHLIGHTITEM, Index, MAKELPARAM(HighLight, 0));
	}

	/*
	获取指定坐标处的标签序号
	返回获取到的标签序号。若没有匹配的标签，返回-1。
	*/
	int HitTest(int X, int Y)
	{
		TCHITTESTINFO tch;
		tch.flags = TCHT_ONITEM;
		tch.pt.x = X;
		tch.pt.y = Y;
		return (int)SendMessage(CurrentHwnd, TCM_HITTEST, 0, (LPARAM)&tch);
	}

	/*
	添加标签到控件中
	Text: 标签的文本
	Index: 标签的序号
	*/
	int InsertItem(char* Text, int Index = -1)
	{
		TCITEM tci = { 0 };
		tci.mask = TCIF_TEXT;
		tci.cchTextMax = 255;
		tci.pszText = Text;
		if (Index == -1)
			Index = GetItemCount();
		return (int)SendMessage(CurrentHwnd, TCM_INSERTITEM, Index, (LPARAM)&tci);
	}

	/*
	设置获得焦点的标签
	Index: 需要获得焦点的标签
	*/
	void SetFocusIndex(int Index)
	{
		SendMessage(CurrentHwnd, TCM_SETCURFOCUS, Index, 0);
	}

	/*
	设置当前选择的标签
	Index: 标签的序号
	返回之前选择的标签的序号
	*/
	int SetCurrIndex(int Index)
	{
		return (int)SendMessage(CurrentHwnd, TCM_SETCURSEL, Index, 0);
	}

	/*
	设置标签文本
	Index: 标签的序号
	Text: 新的文本
	若执行成功，返回TRUE；否则返回FALSE
	*/
	bool SetItemText(int Index, char* Text)
	{
		TCITEM tci = { 0 };
		tci.mask = TCIF_TEXT;
		tci.cchTextMax = 255;
		tci.pszText = Text;
		return (bool)SendMessage(CurrentHwnd, TCM_SETITEM, Index, (LPARAM)&tci);
	}

	/*
	设置标签的大小
	Width: 标签的宽度
	Height: 标签的高度
	返回值: 低位是之前的宽度，高位是之前的高度
	*/
	int SetItemSize(int Width, int Height)
	{
		return (int)SendMessage(CurrentHwnd, TCM_SETITEMSIZE, 0, MAKELPARAM(Width, Height));
	}

	/*
	设置标签最小的大小
	Width: 标签最小的大小。若为-1，则使用系统默认大小
	返回值: 标签之前的最小大小
	*/
	int SetMinTabWidth(int Width)
	{
		return (int)SendMessage(CurrentHwnd, TCM_SETMINTABWIDTH, 0, (LPARAM)Width);
	}
};

//============================================================================
class MyAnimation : public MyControls
{
public:
	bool		AutoPlay;						//自动播放
	bool		Center;							//视频居中播放
	bool		Transparent;					//视频背景透明
	bool		ClientEdge;						//立体边框
	bool		BlackBorder;					//黑色边框

	//创建动画控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;
		OrCalc(AutoPlay, &lStyle, ACS_AUTOPLAY);
		OrCalc(Center, &lStyle, ACS_CENTER);
		OrCalc(Transparent, &lStyle, ACS_TRANSPARENT);
		OrCalc(ClientEdge, &lStyle, WS_EX_CLIENTEDGE);
		OrCalc(BlackBorder, &lStyle, WS_BORDER);

		CurrentHwnd = CreateWindowEx(0, "SysAnimate32", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	停止播放
	若执行成功返回true
	若成功，返回true
	*/
	bool Stop()
	{
		return SendMessage(CurrentHwnd, ACM_STOP, 0, 0);
	}

	/*
	开始播放
	ReplayTimes: 重复播放的次数。若为-1则代表一直重复
	FrameBegin: 从指定的帧开始播放。若为0代表从头开始
	FrameEnd: 播放到指定的帧。若为-1代表播放到末尾
	若成功，返回true
	*/
	bool Play(int ReplayTimes = -1, short FrameBegin = 0, short FrameEnd = -1)
	{
		return SendMessage(CurrentHwnd, ACM_PLAY, ReplayTimes, MAKELPARAM(FrameBegin, FrameEnd));
	}

	/*
	以文件名打开文件
	FileName: 视频文件路径
	若成功，返回true
	*/
	bool Open(char* FileName)
	{
		return SendMessage(CurrentHwnd, ACM_OPEN, 0, (LPARAM)FileName);
	}

	/*
	以资源文件打开文件
	ResID: 资源文件号
	若成功，返回true
	*/
	bool Open(WORD ResID)
	{
		return SendMessage(CurrentHwnd, ACM_OPEN, 0, (LPARAM)MAKEINTRESOURCE(ResID));
	}

	//获取播放状态
	bool IsPlaying()
	{
		return SendMessage(CurrentHwnd, ACM_ISPLAYING, 0, 0);
	}
};

//============================================================================
class MyRichEdit : public MyControls
{
public:
	char*		Text;							//文本
	bool		AutoHScroll;					//自动水平滚动
	bool		AutoVScroll;					//自动垂直滚动
	int			TextPos;						//文本位置
	bool		ForceNumber;					//强制数字
	bool		IsPassword;						//密码文本
	bool		ReadOnly;						//文本只读
	bool		BlackBorder;					//黑色边框
	bool		ClientEdgeBorder;				//立体边框
	bool		SunkenBorder;					//下沉的边框
	bool		Multiline;						//多行文本
	int			ScrollBars;						//滚动条
	bool		DisableNoScroll;				//显示禁用的滚动条
	bool		NoIME;							//禁用输入法
	bool		SelectionBar;					//左边缘空白

	//创建RTF文本框控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;					//初始化控件样式

		//计算控件的样式
		OrCalc(AutoHScroll, &lStyle, ES_AUTOHSCROLL);			//自动水平滚动
		OrCalc(AutoVScroll, &lStyle, ES_AUTOVSCROLL);			//自动垂直滚动
		switch (TextPos)										//文本位置
		{
		case 1:														//中
			lStyle |= ES_CENTER;
			break;

		case 2:														//右
			lStyle |= ES_RIGHT;
			break;
		}
		OrCalc(ForceNumber, &lStyle, ES_NUMBER);				//强制数字
		OrCalc(IsPassword, &lStyle, ES_PASSWORD);				//密码文本
		OrCalc(ReadOnly, &lStyle, ES_READONLY);					//文本只读
		OrCalc(BlackBorder, &lStyle, WS_BORDER);				//黑色边框
		OrCalc(SunkenBorder, &lStyle, ES_SUNKEN);				//下沉的边框
		OrCalc(Multiline, &lStyle, ES_MULTILINE);				//多行文本
		switch (ScrollBars)										//滚动条
		{
		case 1:														//水平
			lStyle |= WS_HSCROLL;
			break;

		case 2:														//垂直
			lStyle |= WS_VSCROLL;
			break;

		case 3:														//两个都有
			lStyle |= WS_HSCROLL | WS_VSCROLL;
			break;
		}
		OrCalc(DisableNoScroll, &lStyle, ES_DISABLENOSCROLL);		//显示禁用的滚动条
		OrCalc(NoIME, &lStyle, ES_NOIME);							//禁用输入法

		//创建控件
		CurrentHwnd = CreateWindowEx(ClientEdgeBorder ? WS_EX_CLIENTEDGE : 0, "MyRichEdit", Text, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		//判断是否为控件加上左边缘空白样式
		OrCalc(SelectionBar, &lStyle, ES_SELECTIONBAR);
		SetWindowLong(CurrentHwnd, GWL_STYLE, lStyle);

		return CurrentHwnd;
	}

	/*
	设置文本框是否自动检测URL
	Enabled: 是否自动检测URL
	若执行成功，函数返回true
	*/
	bool SetUrlDetect(bool Enabled)
	{
		return (!SendMessage(CurrentHwnd, EM_AUTOURLDETECT,
			Enabled ? AURL_ENABLEURL : 0, 0));
	}

	/*
	检测当前文本框是否能进行粘贴操作
	若能进行操作，返回true
	*/
	bool CanPaste()
	{
		return (!SendMessage(CurrentHwnd, EM_CANPASTE, 0, 0));
	}

	/*
	检测当前文本框是否能进行撤销操作
	若能进行操作，返回true
	*/
	bool CanRedo()
	{
		return (!SendMessage(CurrentHwnd, EM_CANREDO, 0, 0));
	}

	/*
	获取文本框中选取文本的范围
	lpMin: 选取的开头
	lpMax: 选取的末尾
	若选取开头为0，选取末尾为-1，则说明全选
	*/
	void GetSelRange(int* lpMin, int* lpMax)
	{
		CHARRANGE cr;
		SendMessage(CurrentHwnd, EM_EXGETSEL, 0, (LPARAM)&cr);
		*lpMin = cr.cpMin;
		*lpMax = cr.cpMax;
	}

	/*
	设置文本框中选取文本的范围
	Min: 选取的开头
	Max: 选取的末尾
	若选取开头为0，选取末尾为-1，则说明全选
	返回实际选取的字符数
	*/
	int SetSelRange(int Min, int Max)
	{
		CHARRANGE cr;
		cr.cpMin = Min;
		cr.cpMax = Max;
		return (int)SendMessage(CurrentHwnd, EM_EXSETSEL, 0, (LPARAM)&cr);
	}

	/*
	设置文本框可输入的最大字符数
	iCount: 最大文本数。若为0，则使用系统默认
	*/
	void SetLimitText(int iCount)
	{
		SendMessage(CurrentHwnd, EM_EXLIMITTEXT, 0, (LPARAM)iCount);
	}

	/*
	在文本框中查找文本
	Find: 需要查找的文本
	MatchCase: 是否需要大小写匹配
	WholeWord: 是否需要全字匹配
	Begin: 开始搜索文本的位置
	End: 结束搜索文本的位置
	返回查找到的文本起始部分。若没有找到，则返回-1
	*/
	int SearchText(char* Find, bool MatchCase = false, bool WholeWord = false, int Begin = 0, int End = -1)
	{
		FINDTEXTEX ft = { 0 };
		ft.chrg.cpMin = Begin;
		ft.chrg.cpMax = End;
		ft.lpstrText = Find;
		return SendMessage(CurrentHwnd, EM_FINDTEXTEX,
			(MatchCase ? FR_MATCHCASE : 0) | (WholeWord ? FR_WHOLEWORD : 0),
			(LPARAM)&ft);
	}

	/* 获取当前选择的字符的格式 */
	CHARFORMAT GetCharFormat()
	{
		CHARFORMAT cf = { 0 };
		cf.cbSize = sizeof(cf);
		cf.dwMask = CFM_ALL;
		SendMessage(CurrentHwnd, EM_GETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
		return cf;
	}

	/* 设置当前选择的字符的格式 */
	bool SetCharFormat(CHARFORMAT cf)
	{
		return (bool)SendMessage(CurrentHwnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
	}

	/*
	设置文本框文本
	NewText: 新的文本
	*/
	void SetText(char* NewText)
	{
		Text = NewText;
		SetWindowText(CurrentHwnd, NewText);
	}

	/*
	以数字设置文本框文本
	NewText: 有效的数字表达式
	*/
	void SetText(int NewText)
	{
		char Buffer[255];
		itoa(NewText, Buffer, 10);
		SetWindowText(CurrentHwnd, Buffer);
	}

	/* 获取文本 */
	char* GetText()
	{
		int Length = GetWindowTextLength(CurrentHwnd) + 1;
		char* tmp = new char[Length];
		GetWindowText(CurrentHwnd, tmp, Length);
		return tmp;
	}

	/* 复制操作 */
	void Copy()
	{
		SendMessage(CurrentHwnd, WM_COPY, 0, 0);
	}

	/* 剪切操作 */
	void Cut()
	{
		SendMessage(CurrentHwnd, WM_CUT, 0, 0);
	}

	/* 撤销操作 */
	void Undo()
	{
		SendMessage(CurrentHwnd, EM_UNDO, 0, 0);
	}

	/* 重复操作 */
	void Redo()
	{
		SendMessage(CurrentHwnd, EM_REDO, 0, 0);
	}

	/* 清除选择的文本 */
	void Clear()
	{
		SendMessage(CurrentHwnd, WM_CLEAR, 0, 0);
	}

	/* 粘贴操作 */
	void Paste()
	{
		SendMessage(CurrentHwnd, WM_PASTE, 0, 0);
	}
};

//============================================================================
class MyTimePicker : public MyControls
{
public:
	bool		LongDateFormat;					//完整时间格式
	bool		RightAlign;						//在右边弹出日历
	bool		CheckBoxes;						//复选框样式
	bool		TimeFormat;						//时间选择器
	bool		UpDownButton;					//使用调节按钮

	//创建日期时间选择器控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;					//初始化控件样式

		//计算控件的样式
		OrCalc(LongDateFormat, &lStyle, DTS_LONGDATEFORMAT);	//完整时间格式
		OrCalc(RightAlign, &lStyle, DTS_RIGHTALIGN);			//往右边弹出日历
		OrCalc(CheckBoxes, &lStyle, DTS_SHOWNONE);				//复选框样式
		OrCalc(TimeFormat, &lStyle, DTS_TIMEFORMAT);			//时间选择器
		OrCalc(UpDownButton, &lStyle, DTS_UPDOWN);				//使用调节按钮

		//创建控件
		CurrentHwnd = CreateWindowEx(0, "MyTimePicker", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	获取可供选择的日期时间范围
	lpMin: 接收起始时间的SYSTEMTIME结构体指针
	lpMax: 接收结束时间的SYSTEMTIME结构体指针
	*/
	void GetRange(SYSTEMTIME* lpMin, SYSTEMTIME* lpMax)
	{
		if (lpMin && lpMax)
		{
			SYSTEMTIME st[2];
			SendMessage(CurrentHwnd, DTM_GETRANGE, 0, (LPARAM)&st);
			*lpMin = st[0];
			*lpMax = st[1];
		}
	}

	/*
	设置可供选的日期时间范围
	Begin: 起始时间的SYSTEMTIME结构体指针，若为NULL则不设置起始时间
	End: 结束时间的SYSTEMTIME结构体指针，若为NULL则不设置结束时间
	若执行成功，返回true
	*/
	bool SetRange(SYSTEMTIME* Begin, SYSTEMTIME* End)
	{
		SYSTEMTIME st[2];
		DWORD lr = 0;

		if (Begin)
		{
			st[0] = *Begin;
			lr |= GDTR_MIN;
		}
		if (End)
		{
			st[1] = *End;
			lr |= GDTR_MAX;
		}

		return SendMessage(CurrentHwnd, DTM_SETRANGE, (WPARAM)lr, (LPARAM)st);
	}

	/*
	获取当前控件所选择的时间
	Time: 用来接收选择的时间的SYSTEMTIME结构体指针
	若执行成功，返回true
	*/
	bool GetTime(SYSTEMTIME* Time)
	{
		if (Time)			//if (Time != NULL)
			return !SendMessage(CurrentHwnd, DTM_GETSYSTEMTIME, 0, (LPARAM)Time);		//return rtn == GDT_VALID;
		else
			return false;
	}

	/*
	设置当前控件所选择的时间
	Time: 用来设置选择的时间的SYSTEMTIME结构体指针。若为NULL，则不选择时间
	若执行成功，返回true
	*/
	bool SetTime(SYSTEMTIME* Time)
	{
		if (&Time)
			return SendMessage(CurrentHwnd, DTM_SETSYSTEMTIME, GDT_VALID, (LPARAM)Time);
		else
			return SendMessage(CurrentHwnd, DTM_SETSYSTEMTIME, GDT_NONE, (LPARAM)Time);
	}

	/*
	设置日期时间格式
	FormatString: 格式字符串。若为NULL，则使用系统默认格式
	若执行成功，返回true
	*/
	bool SetFormat(char* FormatString)
	{
		return SendMessage(CurrentHwnd, DTM_SETFORMAT, 0, (LPARAM)FormatString);
	}
};

//============================================================================
class MyMonthCalendar : public MyControls
{
public:
	bool		MultiSelect;					//连续选取
	int			MultiSelectLimit;				//连续选取数量
	bool		WeekNumbers;					//显示第几周
	bool		NoTodayCircle;					//不圈选今天
	bool		NoToday;						//不显示今天
	bool		BlackBorder;					//黑色边框
	bool		ClientEdgeBorder;				//立体边框

	//创建月历控件
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;					//初始化控件样式

		//计算控件的样式
		OrCalc(MultiSelect, &lStyle, MCS_MULTISELECT);
		OrCalc(WeekNumbers, &lStyle, MCS_WEEKNUMBERS);
		OrCalc(NoTodayCircle, &lStyle, MCS_NOTODAYCIRCLE);
		OrCalc(NoToday, &lStyle, MCS_NOTODAY);
		OrCalc(BlackBorder, &lStyle, WS_BORDER);

		//创建控件
		CurrentHwnd = CreateWindowEx(ClientEdgeBorder ? WS_EX_CLIENTEDGE : 0, "MyMonthCalendar", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		//设置控件的可以选择的最大数量
		SetMaxSelCount(MultiSelectLimit);

		return CurrentHwnd;
	}

	/* 获取当前日历选择的日期 */
	SYSTEMTIME GetCurrentSel()
	{
		SYSTEMTIME st = { 0 };
		SendMessage(CurrentHwnd, MCM_GETCURSEL, 0, (LPARAM)&st);
		return st;
	}

	/*
	设置当前选择的日期
	若执行成功，返回true
	*/
	bool SetCurrentSel(SYSTEMTIME* SelDate)
	{
		if (SelDate)
			return (bool)(SendMessage(CurrentHwnd, MCM_SETCURSEL, 0, (LPARAM)SelDate));
		else
			return false;
	}

	/*
	获取一周中的第一天
	返回值是当前日历中一周的第一天，0代表星期一，1代表星期二，以此类推
	*/
	int GetFirstDayOfWeek()
	{
		return (int)LOWORD((DWORD)SendMessage(CurrentHwnd, MCM_GETFIRSTDAYOFWEEK, 0, 0));
	}

	/*
	设置一周中的第一天
	NewDay: 一周中的第一天。0代表星期一，1代表星期二，以此类推
	*/
	void SetFirstDatOfWeek(int NewDay)
	{
		SendMessage(CurrentHwnd, MCM_SETFIRSTDAYOFWEEK, 0, NewDay);
	}

	/* 获取月历控件可以选择日期的最大数目 */
	int GetMaxSelCount()
	{
		return (int)SendMessage(CurrentHwnd, MCM_GETMAXSELCOUNT, 0, 0);
	}

	/*
	设置月历控件可以选择日期的最大数目
	若执行成功，返回true
	*/
	bool SetMaxSelCount(int MaxCount)
	{
		return (bool)SendMessage(CurrentHwnd, MCM_SETMAXSELCOUNT, (WPARAM)MaxCount, 0);
	}

	/*
	获取选择的范围
	stBegin: 选择的开始的日期
	stEnd: 选择的结束的事件
	若执行成功，返回true
	*/
	bool GetSelRange(SYSTEMTIME* stBegin, SYSTEMTIME* stEnd)
	{
		SYSTEMTIME st[2] = { 0 };
		LRESULT rtn = 0;
		if (stBegin && stEnd)
		{
			rtn = SendMessage(CurrentHwnd, MCM_GETSELRANGE, 0, (LPARAM)&st[0]);
			*stBegin = st[0];
			*stEnd = st[1];
		}
		return rtn;
	}

	/*
	设置选择的范围
	stBegin: 选择的开始的日期
	stEnd: 选择的结束的事件
	若执行成功，返回true
	*/
	bool SetSelRange(SYSTEMTIME* stBegin, SYSTEMTIME* stEnd)
	{
		SYSTEMTIME st[2] = { 0 };
		LRESULT rtn = 0;
		if (stBegin && stEnd)
		{
			st[0] = *stBegin;
			st[1] = *stEnd;
			rtn = SendMessage(CurrentHwnd, MCM_SETSELRANGE, 0, (LPARAM)&st[0]);
		}
		return rtn;
	}

	/*
	获取允许选择日期的范围
	stBegin: 选择的开始的日期
	stEnd: 选择的结束的事件
	*/
	void GetRange(SYSTEMTIME* stBegin, SYSTEMTIME* stEnd)
	{
		SYSTEMTIME st[2] = { 0 };
		if (stBegin && stEnd)
		{
			SendMessage(CurrentHwnd, MCM_GETSELRANGE, 0, (LPARAM)&st[0]);
			*stBegin = st[0];
			*stEnd = st[1];
		}
	}

	/*
	设置允许选择日期的范围
	stBegin: 选择的开始的日期
	stEnd: 选择的结束的事件
	若执行成功，返回true
	*/
	bool SetRange(SYSTEMTIME* stBegin, SYSTEMTIME* stEnd)
	{
		SYSTEMTIME st[2] = { 0 };
		DWORD lr = 0;

		if (stBegin)
		{
			st[0] = *stBegin;
			lr |= GDTR_MIN;
		}
		if (stEnd)
		{
			st[1] = *stEnd;
			lr |= GDTR_MAX;
		}

		return (bool)SendMessage(CurrentHwnd, MCM_SETRANGE, lr, (WPARAM)&st[0]);
	}

	/* 获取今天的日期 */
	SYSTEMTIME GetToday()
	{
		SYSTEMTIME st = { 0 };
		SendMessage(CurrentHwnd, MCM_GETTODAY, 0, (LPARAM)&st);
		return st;
	}

	/* 设置今天的日期 */
	void SetToday(SYSTEMTIME NewDate)
	{
		SendMessage(CurrentHwnd, MCM_SETTODAY, 0, (LPARAM)&NewDate);
	}
};

//============================================================================
class MyIpAddress : public MyControls
{
public:
	//创建IP地址控件
	HWND Create()
	{
		//创建控件
		CurrentHwnd = CreateWindowEx(0, "SysIPAddress32", "", WS_VISIBLE | WS_CHILD,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//设置控件的可视和激活状态
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	//清空IP地址
	void Clear()
	{
		SendMessage(CurrentHwnd, IPM_CLEARADDRESS, 0, 0);
	}

	//获取当前的IP地址
	char* GetIpAddress()
	{
		char* tmp = new char[20];
		GetWindowText(CurrentHwnd, tmp, 20);
		return tmp;
	}

	/*
	判断当前控件是否为空
	若当前控件为空，则返回true
	*/
	bool IsBlank()
	{
		return SendMessage(CurrentHwnd, IPM_ISBLANK, 0, 0);
	}

	/*
	设置当前的IP地址
	Field0 - 3: 分别对应IP地址的四个部分
	*/
	void SetIpAddress(BYTE Field0, BYTE Field1, BYTE Field2, BYTE Field3)
	{
		SendMessage(CurrentHwnd, IPM_SETADDRESS, 0, MAKEIPADDRESS(Field0, Field1, Field2, Field3));
	}

	/*
	让控件获取焦点
	Field: 需要获取焦点的位置序号。其中0位最左方
	*/
	void SetFocus(int Field)
	{
		SendMessage(CurrentHwnd, IPM_SETFOCUS, (WPARAM)Field, 0);
	}

	/*
	设置IP地址对应位置的限制
	Field: 需要设置限制的位置序号。其中0位最左方
	Min: 需要设置的下限值
	Max: 需要设置的上限值
	若执行成功，返回true
	*/
	bool SetRange(int Field, int Min, int Max)
	{
		return SendMessage(CurrentHwnd, IPM_SETRANGE, (WPARAM)Field, MAKEIPRANGE(Min, Max));
	}
};

//============================================================================
class MyWindow
{
public:
	char*		ClassName;						//类名
	char*		Caption;						//窗体标题
	COLORREF	BackColor;						//窗体背景颜色
	DWORD		Style;							//样式
	DWORD		ExStyle;						//扩展样式
	HWND		CurrentHwnd;					//当前句柄
	bool		Visible;						//可视
	bool		Enabled;						//激活
	LONG		Left;							//水平位置
	LONG		Top;							//垂直位置
	LONG		Width;							//窗体宽度
	LONG		Height;							//窗体高度
	RECT		WindowPos;						//当前窗体位置
	HDC			hDC;							//设备上下文句柄
	int			CurrentX;						//当前窗体文本输出的X轴坐标
	int			CurrentY;						//当前窗体文本输出的Y轴坐标
	HINSTANCE   hInstance;						//当前程序实例句柄
	//-----------------------------------------------------------------------
	//窗体所有的控件需要在这里定义
【AllControlsHere】
	//-----------------------------------------------------------------------
	/* 主过程 */
	void InitClass()
	{
【WindowInitCodeHere】
		/* 加载RichEdit动态库 */
		LoadLibrary("RichEd20.dll");
		/* 加载通用控件动态库 */
		LoadLibrary("comctl32.dll");
		/*=======================================*/
		/* 对所有控件进行超类化 */
		ReregisterClass("STATIC", "MyStatic", &PrevStaticProc, StaticProc);								//STATIC
		ReregisterClass("EDIT", "MyEdit", &PrevEditProc, EditProc);										//EDIT
		ReregisterClass("BUTTON", "MyButton", &PrevButtonProc, ButtonProc);								//BUTTON
		ReregisterClass("COMBOBOX", "MyCombo", &PrevComboProc, ComboProc);								//COMBOBOX
		ReregisterClass("LISTBOX", "MyListBox", &PrevListProc, ListProc);								//LISTBOX
		ReregisterClass("SCROLLBAR", "MyScrollBar", &PrevScrollBarProc, ScrollBarProc);					//SCROLLBAR
		ReregisterClass("msctls_updown32", "MyUpDown", &PrevUpDownProc, UpDownProc);					//msctls_updown32
		ReregisterClass("msctls_progress32", "MyProgressBar", &PrevProgressBarProc, ProgressBarProc);	//msctls_progress32
		ReregisterClass("msctls_trackbar32", "MySlider", &PrevSliderProc, SliderProc);					//msctls_trackbar32
		ReregisterClass("msctls_hotkey32", "MyHotkey", &PrevHotkeyProc, HotkeyProc);					//msctls_hotkey32
		ReregisterClass("SysListView32", "MyListView", &PrevListViewProc, ListViewProc);				//SysListView32
		ReregisterClass("SysTreeView32", "MyTreeView", &PrevTreeViewProc, TreeViewProc);				//SysTreeView32
		ReregisterClass("SysTabControl32", "MyTab", &PrevTabProc, TabProc);								//SysTabControl32
		ReregisterClass("RichEdit20A", "MyRichEdit", &PrevRichEditProc, RichEditProc);					//RichEdit20A
		ReregisterClass("SysDateTimePick32", "MyTimePicker", &PrevTimePickerProc, TimePickerProc);		//SysDateTimePick32
		ReregisterClass("SysMonthCal32", "MyMonthCalendar", &PrevMonthCalendarProc, MonthCalendarProc); //SysMonthCal32
	}

	/*
	创建当前的窗体。请在主程序入口调用该过程。
	mHinstance: 由主程序传入的实例句柄（hInstance）
	*/
	HWND Create(HINSTANCE mHinstance)
	{
		InitClass();
		
		WNDCLASS MyClass;																	//窗体类
		MSG Msg;																			//窗体消息

		MyClass.cbClsExtra = 0;
		MyClass.cbWndExtra = 0;
		MyClass.hbrBackground = (HBRUSH)CreateSolidBrush(BackColor);						//背景颜色
		MyClass.hCursor = LoadCursor(0, IDC_ARROW);											//光标
		MyClass.hIcon = LoadIcon(0, IDI_APPLICATION);										//图标
		MyClass.hInstance = mHinstance;														//实例句柄
		MyClass.lpfnWndProc = WndProc;														//消息处理
		MyClass.lpszClassName = TEXT(ClassName);											//类名
		MyClass.lpszMenuName = NULL;
		MyClass.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;

		RegisterClass(&MyClass);															//注册类

		CurrentHwnd = CreateWindowEx(ExStyle, TEXT(ClassName), TEXT(Caption), Style,		//创建窗体
			Left, Top, Width, Height,
			0, NULL, 0, 0);

		tme.cbSize = sizeof(TRACKMOUSEEVENT);												//设置鼠标追踪
		tme.dwFlags = TME_LEAVE;
		tme.hwndTrack = CurrentHwnd;
		TrackMouseEvent(&tme);																//启动鼠标追踪
		UpdateWindow(CurrentHwnd);															//刷新窗体

		hDC = GetDC(CurrentHwnd);															//获取窗体的上下文句柄（hDC）
		hInstance = mHinstance;																//记录当前的实例句柄

		if (CurrentHwnd == 0)																//如果窗体的句柄为0说明创建失败
		{
			UnregisterClass(ClassName, mHinstance);
			return 0;
		}

		CreateAllControls();																//创建所有的控件
		Form_Load();																		//触发窗体加载事件
		
		while (GetMessage(&Msg, CurrentHwnd, 0, 0) > 0)										//窗体主消息循环
		{
			//循环处理所有的消息，直至窗体关闭
			TranslateMessage(&Msg);
			DispatchMessage(&Msg);
		}

		UnregisterClass(ClassName, mHinstance);									//窗体关闭后卸载类

		return CurrentHwnd;																	//函数返回
	}

	/* 窗体的各种方法 */

	//正常关闭窗体
	void Close()
	{
		SendMessage(CurrentHwnd, WM_CLOSE, 0, 0);		//发送WM_CLOSE消息以关闭窗体
	}

	//强行关闭窗体
	void Destroy()
	{
		DestroyWindow(CurrentHwnd);						//强行摧毁窗体				
		PostQuitMessage(0);								//直接退出线程
	}

	/*
	设置窗体标题
	NewCaption: 新的标题字符串
	*/
	void SetCaption(char* NewCaption)
	{
		SetWindowText(CurrentHwnd, NewCaption);			//更改窗体标题
	}

	/*
	用数值来设置窗体标题。
	NewCaption: 有效的数值表达式
	*/
	void SetCaption(int NewCaption)						
	{
		char tmp[255];
		itoa(NewCaption, tmp, 10);
		SetCaption(tmp);
	}

	/*
	设置背景颜色。
	NewColor: 新的颜色
	*/
	void SetBackColor(COLORREF NewColor)
	{
		HBRUSH ColorBrush = CreateSolidBrush(NewColor);							//创建指定颜色的刷子
		SetClassLongPtr(CurrentHwnd, GCLP_HBRBACKGROUND, (LONG)ColorBrush);		//设置类的背景颜色
		SendMessage(CurrentHwnd, WM_ERASEBKGND, (WPARAM)hDC, 0);				//发送WM_ERASEBKGND以刷新窗体
	}

	/*
	设置窗口是否可视。
	IsVisible: 窗体是否可视
	*/
	void SetVisible(bool IsVisible)
	{
		ShowWindow(CurrentHwnd, IsVisible);
		Visible = IsVisible;
	}

	/*
	设置窗口是否禁用。
	IsEnabled: 是否启用窗体
	*/
	void SetEnabled(bool IsEnabled)
	{
		EnableWindow(CurrentHwnd, IsEnabled);
		Enabled = IsEnabled;
	}

	/*
	设置窗体的字体样式。
	参数：
	FontSize: 文本大小
	FontBoldWidth: 文本粗体程度
	FontItalic: 是否为斜体
	FontUnderline: 是否有下划线
	FontStrikeOut: 是否有删除线
	FontName: 字体名称
	*/
	void SetFont(int FontSize, int FontBoldWidth, bool FontItalic, 
		bool FontUnderline, bool FontStrikeOut, char* FontName)
	{
		HFONT MyFont = CreateFont(FontSize, 0, 0, 0, 
			FontBoldWidth, FontItalic, FontUnderline, FontStrikeOut,
			DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, 
			CLEARTYPE_QUALITY, VARIABLE_PITCH, TEXT(FontName));

		SelectObject(hDC, MyFont);
	}

	/*
	设置窗体的样式。
	StyleAdd: 添加的样式
	StyleRemove: 去除的样式
	*/
	void SetStyle(LONG StyleAdd, LONG StyleRemove)
	{
		LONG CurrentLong;
		CurrentLong = GetWindowLong(CurrentHwnd, GWL_STYLE);
		CurrentLong &= (~StyleRemove);
		CurrentLong |= StyleAdd;
		SetWindowLong(CurrentHwnd, GWL_STYLE, CurrentLong);
		Style = CurrentLong;
	}

	/*
	设置窗体文本颜色。
	NewColor: 新的颜色
	*/
	void SetForeColor(COLORREF NewColor)
	{
		SetTextColor(hDC, NewColor);
	}

	/*
	设置窗体输出的文本是否透明。（适用于Print()）
	bTransparent: 是否透明
	*/
	void SetFontTransparent(bool bTransparent)
	{
		SetBkMode(hDC, bTransparent ? OPAQUE : TRANSPARENT);
	}

	/*
	更改窗体位置。
	参数：
	NewLeft: 新的X坐标
	NewTop: 新的Y坐标
	NewWidth: 新的宽度
	NewHeight: 新的高度
	*/
	void Move(int NewLeft, int NewTop, int NewWidth, int NewHeight)
	{
		SetWindowPos(CurrentHwnd, 0, NewLeft, NewTop, NewWidth, NewHeight, 0);
		Left = NewLeft;
		Top = NewTop;
		Width = NewWidth;
		Height = NewHeight;
	} 

	/*
	以指定位置输出文本。
	参数：
	TextPrint: 需要输出的文本
	X: 输出的X轴位置
	Y: 输出的Y轴位置
	*/
	void Print(char* TextPrint, int X, int Y)
	{
		RECT rect;
		SIZE size;

		GetTextExtentPoint32(hDC, TextPrint, strlen(TextPrint), &size);
		SetRect(&rect, X, Y, X + size.cx, Y + size.cy);
		DrawText(hDC, TEXT(TextPrint), -1, &rect, DT_NOCLIP);
	}

	/*
	以窗体默认位置输出文本。
	TextPrint: 需要输出的文本
	*/
	void Print(char* TextPrint)
	{
		RECT rect;
		SIZE size;

		GetTextExtentPoint32(hDC, TextPrint, strlen(TextPrint), &size);
		SetRect(&rect, CurrentX, CurrentY, CurrentX + size.cx, CurrentY + size.cy);
		DrawText(hDC, TEXT(TextPrint), -1, &rect, DT_NOCLIP);

		CurrentX = 0;
		CurrentY += size.cy;
	}

	//隐藏窗体
	void Hide()
	{
		ShowWindow(CurrentHwnd, SW_HIDE);
	}
	
	//显示窗体
	void Show()
	{
		ShowWindow(CurrentHwnd, SW_SHOW);
	}

	/*
	设置窗体最前端
	bTopMost: 是否最前端
	*/
	void SetTopMost(bool bTopMost)
	{
		SetWindowPos(CurrentHwnd, bTopMost ? HWND_TOPMOST : HWND_NOTOPMOST, 
			0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
	}

	/*
	设置窗体透明
	Degree: 窗体透明的程度，范围为0到255。255表示不透明；0表示完全透明。
	*/
	void SetTransparent(unsigned char Degree)
	{
		SetWindowLong(CurrentHwnd, GWL_EXSTYLE, 
			GetWindowLong(CurrentHwnd, GWL_EXSTYLE) | WS_EX_LAYERED);
		SetLayeredWindowAttributes(CurrentHwnd, 0, Degree, LWA_ALPHA);
	}

} MainWindow;

/* 获取当前窗体的位置和大小 */
void GetCurrentRect()
{
	GetWindowRect(MainWindow.CurrentHwnd, &MainWindow.WindowPos);
	MainWindow.Left = MainWindow.WindowPos.left;
	MainWindow.Top = MainWindow.WindowPos.top;
	MainWindow.Width = MainWindow.WindowPos.right - MainWindow.WindowPos.left;
	MainWindow.Height = MainWindow.WindowPos.bottom - MainWindow.WindowPos.top;
}

/* 获取当前窗体的句柄 */
HWND GetCurrentHwnd()
{
	return MainWindow.CurrentHwnd;
}

/* 获取当前窗体的实例句柄 */
HINSTANCE GetCurrentHinstance()
{
	return MainWindow.hInstance;
}


/* 为指定句柄的窗体增加或者去掉指定的样式
 * TargetHwnd：目标窗体的句柄
 * StyleAdd：需要增加的样式
 * StyleRemove：需要去掉的样式
 */
void SetWindowLongEx(HWND TargetHwnd, DWORD StyleAdd = 0, DWORD StyleRemove = 0)
{
	SetWindowLong(TargetHwnd, GWL_STYLE, GetWindowLong(TargetHwnd, GWL_STYLE) | StyleAdd & (~StyleRemove));
}

/*
 * 根据指定的布尔型表达式判断是否执行运算
 * bExpression: 指定的布尔型表达式
 * StyleLong: 需要更改数值的目标
 * StyleAdd: 当布尔型表达式为真的时候将该样式添加到目标数值中
*/
void OrCalc(bool bExpression, long* StyleLong, long StyleAdd)
{
	if (bExpression)
	{
		*StyleLong = *StyleLong | StyleAdd;
	}
}

/*
 * 断点命中过程
 * CodeLn: 断点对应的代码行数
 */
void Breakpoint(long CodeLn)
{
	SendMessage(DEBUGGER_HWND, MY_DEBUGGER_BREAKPOINT, (WPARAM)CodeLn, 0);		//往调试器发送消息，告诉他有断点命中
	SuspendProcess();															//自己挂起自己
}

/*
 * 挂起进程过程
 * ProcessID: 进程ID
 */
void SuspendProcess()
{
	NtSuspendProcess pfnNtSuspendProcess = (NtSuspendProcess)GetProcAddress(
		GetModuleHandle("ntdll"), "NtSuspendProcess");							//获取API的函数地址
	pfnNtSuspendProcess(GetCurrentProcess());									//挂起进程
}

/*
 * 监视断点命中过程
 * WatchIndex：监视的序号
 * VarAddr：变量的地址
 * DataType：变量数据类型
 */
void WatchBreakpoint(int WatchIndex, void* VarAddr, SIZE_T nSize)
{
	SendMessage(DEBUGGER_HWND, MY_DEBUGGER_MEMDATA,								//往调试器发送消息，告诉他需要更新监视信息
		MAKEWPARAM(WatchIndex, nSize), (LPARAM)VarAddr);
}

/*
* 重新注册类过程
* ClassName: 需要重新被注册的类名称
* NewClassName: 新注册的类名称
* lpfnPrevWndProc: 旧的WndProc地址
* lpfnWndProc: 新的WndProc地址
*/
void ReregisterClass(LPCSTR ClassName, LPCSTR NewClassName,
	WNDPROC *lpfnPrevWndProc, WNDPROC lpfnWndProc)
{
	WNDCLASSEX ctlClass;											//控件类
	ZeroMemory(&ctlClass, sizeof(WNDCLASSEX));						//初始化控件类变量

	GetClassInfoEx(GetCurrentHinstance(), ClassName, &ctlClass);	//读取类信息
	*lpfnPrevWndProc = ctlClass.lpfnWndProc;						//记录原WndProc地址
	ctlClass.lpfnWndProc = lpfnWndProc;								//替换掉WndProc地址
	ctlClass.lpszClassName = NewClassName;							//更改类名
	ctlClass.cbSize = sizeof(WNDCLASSEX);							//更改cbSize
	RegisterClassEx(&ctlClass);										//重新注册类
}