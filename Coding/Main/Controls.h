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

typedef LONG(NTAPI *NtSuspendProcess)(IN HANDLE ProcessHandle);		//�������API

const UINT MY_DEBUGGER_BREAKPOINT = 0x8888;		//�������ϵ�������Ϣ
const UINT MY_DEBUGGER_MEMDATA = 0x9999;		//�ڴ��ȡ��Ϣ
const HWND DEBUGGER_HWND = (HWND)��DebuggerHwnd��;

//�������ƶ��ٶȳ���
const int  HS_LargeChange[��NumberOfHS��] = { ��ArrayOfHSLarge�� };
const int  HS_SmallChange[��NumberOfHS��] = { ��ArrayOfHSSmall�� };
const int  VS_LargeChange[��NumberOfVS��] = { ��ArrayOfVSLarge�� };
const int  VS_SmallChange[��NumberOfVS��] = { ��ArrayOfVSSmall�� };

/* ��������ж����� */
POINT CurUpPos;									//����ɿ�ʱ����Ļ������
RECT  CurrentWindowRect;						//��ǰ�����λ�úʹ�С
/* ���λ��׷�� */
TRACKMOUSEEVENT tme;
/* �����ƶ��б�� */
bool bIsMoving = false;							//�Ƿ������ƶ�����
/* ������Ĵ�С��� */
bool bIsResizing = false;						//�Ƿ����ڸ��Ĵ�С
/* ���пؼ�ԭ�ȵ�WndProc��ַ */
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
/* ���峣�� */
/* ���̹��ܼ� */
const int SHIFT_KEY = 0x1;						//Shift��
const int CTRL_KEY = 0x2;						//Ctrl��
const int ALT_KEY = 0x4;						//Alt��

/* ��갴�� */
const int LEFT_BUTTON = 0x1;					//���
const int RIGHT_BUTTON = 0x2;					//�Ҽ�
const int MIDDLE_BUTTON = 0x4;					//�м�

/* ���ְ��� */
const int WHEEL_LEFT_BUTTON = 0x1;				//���
const int WHEEL_RIGHT_BUTTON = 0x2;				//�Ҽ�
const int WHEEL_MIDDLE_BUTTON = 0x4;			//�м�
const int WHEEL_SHIFT_KEY = 4;					//Shift��
const int WHEEL_CTRL_KEY = 8;					//Ctrl��

//============================================================================
/* ���崰�������¼��ĺ���ԭ�� */
void Form_Load();								//�������
void Form_Activate();							//�����ý���
void Form_Deactivate();							//����ʧȥ����
void Form_KeyDown(int, int, bool);				//���尴�°���
void Form_KeyUp(int, int);						//�����ɿ�����
void Form_MouseDown(int, int, int, int);		//��갴������
void Form_MouseMove(int, int, int, int);		//����ƶ�
void Form_MouseUp(int, int, int, int);			//��갴���ɿ�
void Form_MouseWheel(int, int, int, int, int);	//������
void Form_Click();								//�����
void Form_DoubleClick(int, int, int, int);		//���˫��
void Form_MouseLeave();							//����뿪
void Form_Paint();								//�����ػ�
int Form_BeginResize(int);						//���忪ʼ�ı��С
void Form_Resize(int, int, int);				//����ı��С
void Form_FinishResizing();						//����ı��С����
int Form_BeginMove();							//���忪ʼ�ƶ�
void Form_Move(int, int);						//�����ƶ�
void Form_FinishMoving();						//�����ƶ�����
int Form_QueryUnload();							//���彫Ҫ�ر�
void Form_Unload();								//����ر�
/*---------------------------------------------------------------------*/
void CreateAllControls();						//�������еĿؼ�����
void GetCurrentRect();							//��ȡ��ǰ�����λ�úʹ�С
HWND GetCurrentHwnd();							//��ȡ��ǰ����ľ��
HINSTANCE GetCurrentHinstance();				//��ȡ��ǰ�����ʵ�����
void SetWindowLongEx(HWND, DWORD, DWORD);		//���ô�����ʽ
void OrCalc(bool, long*, long);					//��������̣����ڼ��㴰����ʽ��
void Breakpoint(long);							//�ϵ����й���
void WatchBreakpoint(int, void*, SIZE_T);		//���ӵ����й���
void SuspendProcess();							//������̹���
void ReregisterClass(LPCSTR, LPCSTR,
	WNDPROC*, WNDPROC);							//����ע�������
/*---------------------------------------------------------------------*/
/* ���������пؼ��¼��ĺ���ԭ�� */
��AllEventsDefHere��

//============================================================================
/* ���㵱ǰ��Shiftֵ */
int GetShiftValue()
{
	int ShiftValue = 0;
	if (GetAsyncKeyState(VK_SHIFT))		ShiftValue |= SHIFT_KEY;		//Shift��
	if (GetAsyncKeyState(VK_CONTROL))	ShiftValue |= CTRL_KEY;			//Ctrl��
	if (GetAsyncKeyState(VK_MENU))		ShiftValue |= ALT_KEY;			//Alt��
	return ShiftValue;
}

//============================================================================
/* ��������Ϣ���� */
LRESULT CALLBACK WndProc(HWND hWnd,	UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int TargetHmenu = 0;										//��ĳЩ�¼��л�õĿؼ���hMenu
	int CurrentPos = 0;											//�������¼���������¼����λ�õı���

	switch (uMsg)
	{
	case WM_MOUSEMOVE:											//����ƶ�
		TrackMouseEvent(&tme);										//����׷������ƶ�
		Form_MouseMove(wParam & ~(MK_CONTROL | MK_SHIFT),			//ȡ����갴��״̬
			wParam & ~(wParam & ~(MK_CONTROL | MK_SHIFT)),			//ȡ��ϵͳ���ܼ�����״̬
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//ȡ���������
		break;

	case WM_MOUSELEAVE:											//����뿪����
		Form_MouseLeave();
		TrackMouseEvent(NULL);										//ֹͣ׷������ƶ�
		break;

	case WM_LBUTTONDOWN:										//�������
		SetCapture(hWnd);											//���ô�����겶��
		Form_MouseDown(1, GetShiftValue(),							//�����������������Ϣ
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//ȡ���������
		break;

	case WM_LBUTTONUP:											//����ɿ�
		GetCursorPos(&CurUpPos);									//��ȡ�������
		ReleaseCapture();											//ֹͣ������겶��
		GetWindowRect(hWnd, &CurrentWindowRect);					//��ô�������
		Form_MouseUp(1, GetShiftValue(),							//������������ɿ���Ϣ
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//ȡ���������
		/* �����걣���ڴ��巶Χ�ڵĻ��ʹ�������¼� */
		if ((CurUpPos.x > CurrentWindowRect.left) && (CurUpPos.x < CurrentWindowRect.right) &&
			(CurUpPos.y > CurrentWindowRect.top) && (CurUpPos.y < CurrentWindowRect.bottom))
		{
			Form_Click();
		}
		break;

	case WM_RBUTTONDOWN:										//�Ҽ�����
		SetCapture(hWnd);											//���ô�����겶��
		Form_MouseDown(2, GetShiftValue(), 							//���������Ҽ�������Ϣ
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//ȡ���������
		break;

	case WM_RBUTTONUP:											//�Ҽ��ɿ�
		GetCursorPos(&CurUpPos);									//��ȡ�������
		ReleaseCapture();											//ֹͣ������겶��
		GetWindowRect(hWnd, &CurrentWindowRect);					//��ô�������
		Form_MouseUp(2, GetShiftValue(), 							//���������Ҽ��ɿ���Ϣ
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//ȡ���������
		/* �����걣���ڴ��巶Χ�ڵĻ��ʹ�������¼� */
		if ((CurUpPos.x > CurrentWindowRect.left) && (CurUpPos.x < CurrentWindowRect.right) &&
			(CurUpPos.y > CurrentWindowRect.top) && (CurUpPos.y < CurrentWindowRect.bottom))
		{
			Form_Click();
		}
		break;

	case WM_MBUTTONDOWN:										//�м�����
		SetCapture(hWnd);											//���ô�����겶��
		Form_MouseDown(4, GetShiftValue(), 							//���������м�������Ϣ
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//ȡ���������
		break;

	case WM_MBUTTONUP:											//�м��ɿ�
		GetCursorPos(&CurUpPos);									//��ȡ�������
		ReleaseCapture();											//ֹͣ������겶��
		GetWindowRect(hWnd, &CurrentWindowRect);					//��ô�������
		Form_MouseUp(4, GetShiftValue(), 							//���������м��ɿ���Ϣ
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));			//ȡ���������
		/* �����걣���ڴ��巶Χ�ڵĻ��ʹ�������¼� */
		if ((CurUpPos.x > CurrentWindowRect.left) && (CurUpPos.x < CurrentWindowRect.right) &&
			(CurUpPos.y > CurrentWindowRect.top) && (CurUpPos.y < CurrentWindowRect.bottom))
		{
			Form_Click();
		}
		break;

	case WM_MOUSEWHEEL:											//������
		//������������¼�
		Form_MouseWheel(GET_WHEEL_DELTA_WPARAM(wParam),								//����0�����Ϲ��� ���¹�����֮
			GET_KEYSTATE_WPARAM(wParam) & ~(MK_CONTROL | MK_SHIFT),					//��ȡ��갴��״̬
			GET_KEYSTATE_WPARAM(wParam) & ~(wParam & ~(MK_CONTROL | MK_SHIFT)),		//��ȡϵͳ���ܼ�״̬
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));							//��ȡ�������
		break;

	case WM_LBUTTONDBLCLK:
		Form_DoubleClick(1,	GET_KEYSTATE_WPARAM(wParam) & ~(wParam & ~(MK_CONTROL | MK_SHIFT)),		//��ȡϵͳ���ܼ�״̬
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));											//��ȡ�������
		break;

	case WM_RBUTTONDBLCLK:
		Form_DoubleClick(2, GET_KEYSTATE_WPARAM(wParam) & ~(wParam & ~(MK_CONTROL | MK_SHIFT)),		//��ȡϵͳ���ܼ�״̬
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));											//��ȡ�������
		break;

	case WM_MBUTTONDBLCLK:
		Form_DoubleClick(4, GET_KEYSTATE_WPARAM(wParam) & ~(wParam & ~(MK_CONTROL | MK_SHIFT)),		//��ȡϵͳ���ܼ�״̬
			GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));											//��ȡ�������
		break;

	case WM_SETFOCUS:											//�����ý���
		Form_Activate();
		break;

	case WM_KILLFOCUS:											//����ʧȥ����
		if (bIsMoving)												//���״ָ̬���������ƶ� ���� ����ʧȥ����
		{															//˵�������Ѿ�ֹͣ�ƶ������������ƶ��¼�
			Form_FinishMoving();
		}
		Form_Deactivate();											//��������ʧȥ�����¼�
		break;

	case WM_KEYDOWN:											//���̰��°���
		Form_KeyDown(wParam, GetShiftValue(),						//��ȡAscii���ϵͳ���ܼ�״̬ 
			(bool)((lParam >> 30) != 0));							//��lParam��30λ�Եõ��Ƿ񳤰�����
		break;

	case WM_KEYUP:												//�����ɿ�����
		Form_KeyUp(wParam, GetShiftValue());
		break;

	case WM_ERASEBKGND:											//�����ػ�
		Form_Paint();
		break;

	case WM_SIZING:												//���忪ʼ�ı��С
		bIsResizing = true;											//��¼Ϊ�������ڸ��Ĵ�С
		if (Form_BeginResize(wParam) != 0)							//����������ط�0ֵ��ȡ�����Ĵ�С
		{
			bIsResizing = false;										//��¼Ϊ����δ�ڸ��Ĵ�С
			ReleaseCapture();											//�ͷŴ�����겶��ȡ�����Ĵ�С
		}
		break;

	case WM_SIZE:												//����ı��С
		GetCurrentRect();											//��¼��ǰ����λ�úʹ�С
		Form_Resize(wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));	//���������ƶ��¼�
		break;

	case WM_SYSCOMMAND:
		if ((wParam == SC_MOVE) || (wParam == SC_MOVE + 2))		//���ص����忪ʼ�ƶ���Ϣ
		{
			bIsMoving = true;										//��¼Ϊ���������ƶ�
			if (Form_BeginMove() != 0)								//����������ط�0ֵ��ȡ���ƶ�
			{
				bIsMoving = false;
				return 0;
			}
			break;													//�����������ƶ�
		}
		break;														//���������¼��ͷ���
		
	case WM_COMMAND:											//�ؼ��¼�
		switch (LOWORD(wParam))										//��ȡ�ؼ���hMenu
		{
		��ControlEventsHere��
		}
		break;

	case WM_NOTIFY:												//�ؼ��¼�
		switch ((*(NMHDR*)lParam).idFrom)							//��ȡ�ؼ���hMenu
		{
		��ControlNotifyCodeHere��
		}
		break;

	case WM_HSCROLL:											//��������Ϣ
	case WM_VSCROLL:
		TargetHmenu = (int)GetMenu((HWND)lParam);					//��ȡ�ؼ�hMenu
		if (TargetHmenu)											//hMenu����Ϊ��
		{
			switch (LOWORD(wParam))										//�Բ�ͬ�Ĺ�����ʽ���д���
			{
			case SB_THUMBPOSITION:										//�û��϶�����
			case SB_THUMBTRACK:
				CurrentPos = HIWORD(wParam);								//��ȡ����λ��
				break;

			case SB_PAGELEFT:											//��������ƶ�
				CurrentPos = GetScrollPos((HWND)lParam, SB_CTL);			//��ȡ����λ��
				CurrentPos = CurrentPos - ((uMsg == WM_HSCROLL) ?
					HS_LargeChange[TargetHmenu - 1] :
					VS_LargeChange[TargetHmenu - 1]);
				break;

			case SB_PAGERIGHT:											//���ҿ����ƶ�
				CurrentPos = GetScrollPos((HWND)lParam, SB_CTL);			//��ȡ����λ��
				CurrentPos = CurrentPos + ((uMsg == WM_HSCROLL) ?
					HS_LargeChange[TargetHmenu - 1] :
					VS_LargeChange[TargetHmenu - 1]);
				break;

			case SB_LINELEFT:											//���������ƶ�
				CurrentPos = GetScrollPos((HWND)lParam, SB_CTL);			//��ȡ����λ��
				CurrentPos = CurrentPos - ((uMsg == WM_HSCROLL) ?
					HS_SmallChange[TargetHmenu - 1] :
					VS_SmallChange[TargetHmenu - 1]);
				break;

			case SB_LINERIGHT:											//���������ƶ�
				CurrentPos = GetScrollPos((HWND)lParam, SB_CTL);			//��ȡ����λ��
				CurrentPos = CurrentPos + ((uMsg == WM_HSCROLL) ?
					HS_SmallChange[TargetHmenu - 1] :
					VS_SmallChange[TargetHmenu - 1]);
				break;

			case SB_ENDSCROLL:											//ֹͣ�϶�
				CurrentPos = GetScrollPos((HWND)lParam, SB_CTL);			//��ȡ����λ��
				break;
			}
			SetScrollPos((HWND)lParam, SB_CTL, CurrentPos, TRUE);		//���»���λ��
		}
		break;
		
	case WM_MOVE:												//�����ƶ�
		GetCurrentRect();											//��¼��ǰ����λ�úʹ�С
		Form_Move(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));		//���������ƶ��¼�
		break;
	
	case WM_CAPTURECHANGED:										//������겶����Ϣ
		if (bIsMoving && (GetAsyncKeyState(VK_LBUTTON) == 0))		//���״ָ̬���������ƶ� ���� �������Ѿ��ɿ�
		{															//˵�������Ѿ�ֹͣ�ƶ������������ƶ��¼�
			bIsMoving = false;
			Form_FinishMoving();
		}
		break;

	case WM_NCHITTEST:											//��������ڴ���������λ�õ��¼�
		if (bIsMoving && (GetCapture() != GetCurrentHwnd()))		//���״ָ̬���������ƶ� ���� ����ȴ�������겶��
		{															//˵�������Ѿ�ֹͣ�ƶ������������ƶ��¼�
			bIsMoving = false;
			Form_FinishMoving();
		}
		break;

	case WM_EXITSIZEMOVE:										//�����˳�������С�����˳��ƶ�״̬
		if (bIsResizing)											//����������ڸ��Ĵ�С
		{
			bIsResizing = false;
			Form_FinishResizing();										//��������ı��С�����¼�
		}
		break;

	case WM_CLOSE:												//���彫Ҫ�ر�
		if (Form_QueryUnload() != 0)								//����ȡ���ر�
		{
			return 0;													//����������ط�0ֵ��ȡ������ر�
		}
		break;													//����������ر�

	case WM_DESTROY:											//����ر�
		Form_Unload();
		return 0;
	}

	return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* STATIC�ؼ���Ϣ���� */
LRESULT CALLBACK StaticProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)								//����ֵ2
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);					//��ȡ��ǰ�ؼ��ı�ʶ��
	switch (CurrCtlHmenu)
	{
	��StaticProcCode��
	}

	return CallWindowProc(PrevStaticProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* EDIT�ؼ���Ϣ���� */
LRESULT CALLBACK EditProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);					//��ȡ��ǰ�ؼ��ı�ʶ��
	switch (CurrCtlHmenu)
	{
	��EditProcCode��
	}

	return CallWindowProc(PrevEditProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* BUTTON�ؼ���Ϣ���� */
LRESULT CALLBACK ButtonProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);					//��ȡ��ǰ�ؼ��ı�ʶ��
	switch (CurrCtlHmenu)
	{
	��ButtonProcCode��
	}

	return CallWindowProc(PrevButtonProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* COMBOBOX�ؼ���Ϣ���� */
LRESULT CALLBACK ComboProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	��ComboProcCode��
	}

	return CallWindowProc(PrevComboProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* LISTBOX�ؼ���Ϣ���� */
LRESULT CALLBACK ListProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	��ListProcCode��
	}

	return CallWindowProc(PrevListProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* SCROLLBAR�ؼ���Ϣ���� */
LRESULT CALLBACK ScrollBarProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	SCROLLINFO sInfo;

	switch (CurrCtlHmenu)
	{
	��ScrollBarProcCode��
	}

	return CallWindowProc(PrevScrollBarProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* UpDown�ؼ���Ϣ���� */
LRESULT CALLBACK UpDownProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	SCROLLINFO si;

	switch (CurrCtlHmenu)
	{
	��UpDownProcCode��
	}

	return CallWindowProc(PrevUpDownProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* ProgressBar�ؼ���Ϣ���� */
LRESULT CALLBACK ProgressBarProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	��ProgressBarProcCode��
	}
	return CallWindowProc(PrevProgressBarProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* Slider�ؼ���Ϣ���� */
LRESULT CALLBACK SliderProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	��SliderProcCode��
	}
	return CallWindowProc(PrevSliderProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* Hotkey�ؼ���Ϣ���� */
LRESULT CALLBACK HotkeyProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	��HotkeyProcCode��
	}
	return CallWindowProc(PrevHotkeyProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* ListView�ؼ���Ϣ���� */
LRESULT CALLBACK ListViewProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	��ListViewProcCode��
	}
	return CallWindowProc(PrevListViewProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* TreeView�ؼ���Ϣ���� */
LRESULT CALLBACK TreeViewProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	��TreeViewProcCode��
	}
	return CallWindowProc(PrevTreeViewProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* Tab�ؼ���Ϣ���� */
LRESULT CALLBACK TabProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	��TabProcCode��
	}
	return CallWindowProc(PrevTabProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* RichEdit�ؼ���Ϣ���� */
LRESULT CALLBACK RichEditProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);
	switch (CurrCtlHmenu)
	{
	��RichEditProcCode��
	}
	return CallWindowProc(PrevRichEditProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* TimePicker�ؼ���Ϣ���� */
LRESULT CALLBACK TimePickerProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);					//��ȡ��ǰ�ؼ��ı�ʶ��
	switch (CurrCtlHmenu)
	{
	��TimePickerProcCode��
	}
	return CallWindowProc(PrevTimePickerProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* MonthCalendar�ؼ���Ϣ���� */
LRESULT CALLBACK MonthCalendarProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int CurrCtlHmenu = (int)GetMenu(hWnd);					//��ȡ��ǰ�ؼ��ı�ʶ��
	switch (CurrCtlHmenu)
	{
	��MonthCalendarProcCode��
	}
	return CallWindowProc(PrevMonthCalendarProc, hWnd, uMsg, wParam, lParam);
}

//============================================================================
/* �����ʱ����Ϣ���� */
void CALLBACK TimerProc(HWND hWnd, UINT uMsg, UINT iTimerID, DWORD dwTime)
{
	switch (iTimerID)
	{
��AllTimerIDHere��
	}
}

//============================================================================
/* ������ֿؼ�ͨ�õĿؼ��ࣨ����ÿ���ؼ���ͬ�����Ժ͹��̣���
֮��Ŀؼ������Լ̳����������Ժ͹��� */
class MyControls
{
public:
	int			hMenu;							//�ؼ���hMenu��Ψһ��ʶ����
	LONG		Left;							//ˮƽλ��
	LONG		Top;							//��ֱλ��
	LONG		Width;							//������
	LONG		Height;							//����߶�
	HDC			hDC;							//�豸�����ľ��
	HWND		CurrentHwnd;					//��ǰ�ؼ��ľ��
	bool		Visible;						//�Ƿ����
	bool		Enabled;						//�Ƿ񼤻�

	//ɾ���ؼ�
	void Unload()
	{
		DestroyWindow(CurrentHwnd);
	}

	/*
	���Ŀؼ�λ��
	NewLeft: �µ�X����
	NewTop: �µ�Y����
	NewWidth: �µĿ��
	NewHeight: �µĸ߶�
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
	���ÿؼ��Ƿ����
	bVisible: �ؼ��Ƿ����
	*/
	void SetVisible(bool bVisible)
	{
		ShowWindow(CurrentHwnd, bVisible ? SW_SHOW : SW_HIDE);
		Visible = bVisible;
	}

	/*
	���ÿؼ��Ƿ����
	IsEnabled: �Ƿ����ô���
	*/
	void SetEnabled(bool IsEnabled)
	{
		EnableWindow(CurrentHwnd, IsEnabled);
		Enabled = IsEnabled;
	}
};

//============================================================================
/* ���尴ť�ؼ�����򡢵�ѡ�򡢶�ѡ��ͨ�õİ�ť�ؼ��ࣨ�������ǹ��е����Ժ͹��̣���
���ǵĿؼ������Լ̳����������Ժ͹��� */
class ButtonPublicClass : public MyControls
{
public:
	char*		Text;							//�ı�
	DWORD		TextPos;						//�ı�λ��

	/*
	���ð�ť���ı�
	NewText: �µ��ı�
	*/
	void SetText(char* NewText)
	{
		Text = NewText;
		SetWindowText(CurrentHwnd, NewText);
	}

	/*
	���������ð�ť���ı�
	NewText: ��Ч����ֵ���ʽ
	*/
	void SetText(int NewText)
	{
		char Buffer[255];
		itoa(NewText, Buffer, 10);
		SetWindowText(CurrentHwnd, Buffer);
	}

	/* ģ��������Ĳ��� */
	void Click()
	{
		SendMessage(CurrentHwnd, BM_CLICK, 0, 0);
	}

	/* ��ȡ�ı� */
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
	UINT_PTR	ID;								//��ʱ��ID
	UINT		Interval;						//��ʱ���
	bool		Enabled;						//��ʱ���Ƿ񱻼���

	//������ʱ��
	bool Create()
	{
		Enabled = (bool)(SetTimer(GetCurrentHwnd(), ID, Interval, TimerProc) != 0);
		return Enabled;
	}
	
	//ɾ����ʱ��
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
	//����ͼƬ�ؼ�
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
	int			TextPos;						//�ı�λ��
	char*		Caption;						//��ǩ����
	bool		BlackBorder;					//�Ƿ��к�ɫ�߿�
	bool		BlackFilled;					//�Ƿ��Ժ�ɫ���
	bool		AutoNextLine;					//�Ƿ��Զ�����
	bool		AutoEllipsis;					//�Ƿ��Զ����ʡ�Ժ�

	//������ǩ�ؼ�
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD | SS_NOTIFY;		//��ʼ���ؼ���ʽ

		//����ؼ�����ʽ
		OrCalc(BlackBorder, &lStyle, WS_BORDER);				//��ɫ�߿�
		OrCalc(BlackFilled, &lStyle, SS_BLACKRECT);				//��ɫ���
		switch (TextPos)										//�ı�λ��
		{
		case 1:														//��
			lStyle |= SS_CENTER;
			break;

		case 2:														//��
			lStyle |= SS_RIGHT;
			break;
		}
		OrCalc(AutoNextLine, &lStyle, SS_EDITCONTROL);			//�Զ�����
		OrCalc(AutoEllipsis, &lStyle, SS_ENDELLIPSIS);			//�Զ����ʡ�Ժ�

		//�����ؼ�
		CurrentHwnd = CreateWindowEx(0, "MyStatic", Caption, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);
		return CurrentHwnd;
	}

	/*
	���ñ�ǩ����
	NewCaption: �µı����ַ���
	*/
	void SetCaption(char* NewCaption)
	{
		SetWindowText(CurrentHwnd, NewCaption);			//���ı�ǩ����
		Caption = NewCaption;
	}

	/*
	����ֵ�����ñ�ǩ���⡣
	NewCaption: ��Ч����ֵ���ʽ
	*/
	void SetCaption(int NewCaption)
	{
		char tmp[255];
		itoa(NewCaption, tmp, 10);
		SetCaption(tmp);
	}

	/*
	���ñ�ǩ�ı�λ��
	NewPos: ��ǩ���ı�λ�á�
	*/
	void SetTextPos(int NewPos)
	{
		switch (NewPos)									//���ݲ�ͬ����ֵ���ı�ǩ�ı�λ��
		{
		case 0:												//��
			SetWindowLongEx(CurrentHwnd, SS_LEFT, SS_CENTER | SS_RIGHT);
			break;

		case 1:												//��
			SetWindowLongEx(CurrentHwnd, SS_CENTER, SS_LEFT | SS_RIGHT);
			break;

		case 2:												//��
			SetWindowLongEx(CurrentHwnd, SS_RIGHT, SS_LEFT | SS_CENTER);
			break;

		}
		TextPos = NewPos;
	}

	/*
	���ñ�ǩ�ؼ��Ƿ��Զ�����
	bAutoNextLine: �Ƿ��Զ�����
	*/
	void SetAutoNextLine(bool bAutoNextLine)
	{
		bAutoNextLine ?
			SetWindowLongEx(CurrentHwnd, SS_EDITCONTROL, 0):
			SetWindowLongEx(CurrentHwnd, 0, SS_EDITCONTROL);
		AutoNextLine = bAutoNextLine;
	}

	/*
	���ñ�ǩ�ؼ��Ƿ��Զ����ʡ�Ժ�
	bAutoEllipsis: �Ƿ��Զ����ʡ�Ժ�
	*/
	void SetAutoEllipsis(bool bAutoEllipsis)
	{
		bAutoEllipsis ?
			SetWindowLongEx(CurrentHwnd, SS_ENDELLIPSIS, 0):
			SetWindowLongEx(CurrentHwnd, 0, SS_ENDELLIPSIS);
		AutoEllipsis = bAutoEllipsis;
	}

	/*
	���ñ�ǩ�ؼ��Ƿ��Ժ�ɫ���
	bBlackFilled: �Ƿ��Ժ�ɫ���
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
	char*		Text;							//�ı�
	bool		AutoHScroll;					//�Զ�ˮƽ����
	bool		AutoVScroll;					//�Զ���ֱ����
	int			TextPos;						//�ı�λ��
	bool		ForceLowercase;					//ǿ��Сд
	bool		ForceUppercase;					//ǿ�ƴ�д
	bool		ForceNumber;					//ǿ������
	bool		IsPassword;						//�����ı�
	char		PasswordChar;					//�����ַ�
	bool		ReadOnly;						//�ı�ֻ��
	bool		BlackBorder;					//��ɫ�߿�
	bool		ClientEdgeBorder;				//����߿�
	bool		Multiline;						//�����ı�
	int			ScrollBars;						//������

	//�����ı��ؼ�
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;					//��ʼ���ؼ���ʽ

		//����ؼ�����ʽ
		OrCalc(AutoHScroll, &lStyle, ES_AUTOHSCROLL);			//�Զ�ˮƽ����
		OrCalc(AutoVScroll, &lStyle, ES_AUTOVSCROLL);			//�Զ���ֱ����
		switch (TextPos)										//�ı�λ��
		{
		case 1:														//��
			lStyle |= ES_CENTER;
			break;

		case 2:														//��
			lStyle |= ES_RIGHT;
			break;
		}
		OrCalc(ForceLowercase, &lStyle, ES_LOWERCASE);			//ǿ��Сд
		OrCalc(ForceUppercase, &lStyle, ES_UPPERCASE);			//ǿ�ƴ�д
		OrCalc(ForceNumber, &lStyle, ES_NUMBER);				//ǿ������
		OrCalc(IsPassword, &lStyle, ES_PASSWORD);				//�����ı�
		OrCalc(ReadOnly, &lStyle, ES_READONLY);					//�ı�ֻ��
		OrCalc(BlackBorder, &lStyle, WS_BORDER);				//��ɫ�߿�
		OrCalc(Multiline, &lStyle, ES_MULTILINE);				//�����ı�
		switch (ScrollBars)										//������
		{
		case 1:														//ˮƽ
			lStyle |= WS_HSCROLL;
			break;

		case 2:														//��ֱ
			lStyle |= WS_VSCROLL;
			break;

		case 3:														//��������
			lStyle |= WS_HSCROLL | WS_VSCROLL;
			break;
		}

		//�����ؼ�
		CurrentHwnd = CreateWindowEx(ClientEdgeBorder ? WS_EX_CLIENTEDGE : 0, "MyEdit", Text, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		//���������ַ�
		if (IsPassword)		SendMessage(CurrentHwnd, EM_SETPASSWORDCHAR, (WPARAM)PasswordChar, 0);

		return CurrentHwnd;
	}

	/*
	�����ı����ı�
	NewText: �µ��ı�
	*/
	void SetText(char* NewText)
	{
		Text = NewText;
		SetWindowText(CurrentHwnd, NewText);
	}

	/*
	�����������ı����ı�
	NewText: ��Ч�����ֱ��ʽ
	*/
	void SetText(int NewText)
	{
		char Buffer[255];
		itoa(NewText, Buffer, 10);
		SetWindowText(CurrentHwnd, Buffer);
	}

	/*
	�����ı����Ƿ�ǿ��Ҫ��Сд
	bLowerCase: �Ƿ�ǿ��Ҫ��Сд
	*/
	void SetForceLowerCase(bool bLowerCase)
	{
		bLowerCase ?
			SetWindowLongEx(CurrentHwnd, ES_LOWERCASE, 0) :
			SetWindowLongEx(CurrentHwnd, 0, ES_LOWERCASE);
		ForceLowercase = bLowerCase;
	}

	/*
	�����ı����Ƿ�ǿ��Ҫ���д
	bUpperCase: �Ƿ�ǿ��Ҫ���д
	*/
	void SetForceUpperCase(bool bUpperCase)
	{
		bUpperCase ?
			SetWindowLongEx(CurrentHwnd, ES_UPPERCASE, 0) :
			SetWindowLongEx(CurrentHwnd, 0, ES_UPPERCASE);
		ForceUppercase = bUpperCase;
	}

	/*
	�����ı����Ƿ�ǿ��Ҫ������
	bNumber: �Ƿ�ǿ��Ҫ������
	*/
	void SetForceNumber(bool bNumber)
	{
		bNumber ?
			SetWindowLongEx(CurrentHwnd, ES_NUMBER, 0) :
			SetWindowLongEx(CurrentHwnd, 0, ES_NUMBER);
		ForceNumber = bNumber;
	}

	/*
	�����ı����Ƿ��ı�ֻ��
	bReadOnly: �Ƿ��ı�ֻ��
	*/
	void SetReadOnly(bool bReadOnly)
	{
		SendMessage(CurrentHwnd, EM_SETREADONLY, bReadOnly, 0);
		ReadOnly = bReadOnly;
	}

	/* ����ı���ĳ������� */
	void EmptyUndoBuffer()
	{
		SendMessage(CurrentHwnd, EM_EMPTYUNDOBUFFER, 0, 0);
	}

	/* ��ȡ�ı����ܿ����ĵ�һ�� */
	long GetFirstVisibleLine()
	{
		return (long)SendMessage(CurrentHwnd, EM_GETFIRSTVISIBLELINE, 0, 0);
	}

	/* ��ȡ�ı����Ʒ�Χ */
	long GetLimitText()
	{
		return (long)SendMessage(CurrentHwnd, EM_GETLIMITTEXT, 0, 0);
	}

	/*
	�����ı����Ʒ�Χ
	NewLimit: �µ��ı������������ơ����Ϊ0�������ơ�
	*/
	void SetLimitText(long NewLimit)
	{
		SendMessage(CurrentHwnd, EM_SETLIMITTEXT, NewLimit, 0);
	}

	/* ��ȡ�ı������� */
	long GetLineCount()
	{
		return (long)SendMessage(CurrentHwnd, EM_GETLINECOUNT, 0, 0);
	}

	/*
	�����ı������߾�
	NewMargin: �µ���߾�
	*/
	void SetLeftMargin(int NewMargin)
	{
		SendMessage(CurrentHwnd, EM_SETMARGINS, EC_LEFTMARGIN, MAKELPARAM(NewMargin, 0));
	}

	/*
	�����ı�����ұ߾�
	NewMargin: �µ��ұ߾�
	*/
	void SetRightMargin(int NewMargin)
	{
		SendMessage(CurrentHwnd, EM_SETMARGINS, EC_RIGHTMARGIN, MAKELPARAM(0, NewMargin));
	}

	/* ��ȡ�ı������߾� */
	int GetLeftMargin()
	{
		return (int)(LOWORD(SendMessage(CurrentHwnd, EM_GETMARGINS, 0, 0)));
	}

	/* ��ȡ�ı�����ұ߾� */
	int GetRightMargin()
	{
		return (int)(HIWORD(SendMessage(CurrentHwnd, EM_GETMARGINS, 0, 0)));
	}


	/*
	�����ı���������ı�
	NewPasswordChar: �µ������ַ�
	*/
	void SetPasswordChar(char NewPasswordChar)
	{
		SendMessage(CurrentHwnd, EM_SETPASSWORDCHAR, (WPARAM)NewPasswordChar, 0);
		PasswordChar = NewPasswordChar;
	}

	/*
	�����ı���������ı�
	NewPasswordChar: �µ������ַ���Ascii��
	*/
	void SetPasswordChar(int NewPasswordChar)
	{
		SendMessage(CurrentHwnd, EM_SETPASSWORDCHAR, (WPARAM)NewPasswordChar, 0);
		PasswordChar = (char)NewPasswordChar;
	}

	/* ��ȡ�ı���������ı� */
	char GetPasswordChar()
	{
		return (char)SendMessage(CurrentHwnd, EM_GETPASSWORDCHAR, 0, 0);
	}

	/* ��ȡ�ı����ѡȡ�ı����� */
	DWORD GetSelLength()
	{
		DWORD StartPos = 0, EndPos = 0;
		SendMessage(CurrentHwnd, EM_GETSEL, (WPARAM)&StartPos, (LPARAM)&EndPos);
		return (EndPos - StartPos);
	}

	/* ��ȡ�ı����ѡȡ�ı���ͷ */
	DWORD GetSelStart()
	{
		DWORD StartPos = 0, EndPos = 0;
		SendMessage(CurrentHwnd, EM_GETSEL, (WPARAM)&StartPos, (LPARAM)&EndPos);
		return StartPos;
	}

	/*
	�����ı����ѡȡ�ı�����
	���StartPosΪ0��EndPosΪ-1�����ı�ȫѡ��
	���StartPosΪ-1�����ı�ȡ��ѡ��
	���EndPosΪ-1�����ı�ѡȡ��Χ��StartPos���ı�ĩβ��
	*/
	void SetSel(DWORD StartPos, DWORD EndPos)
	{
		SendMessage(CurrentHwnd, EM_SETSEL, (WPARAM)StartPos, (LPARAM)EndPos);
	}

	/*
	�����ı���Ĺ��λ��
	StartPos: �µĹ��λ��
	���StartPosΪ-1�����ı�ȡ��ѡ��
	*/
	void SetSelStart(DWORD StartPos)
	{
		SendMessage(CurrentHwnd, EM_SETSEL, (WPARAM)StartPos, (LPARAM)StartPos);
	}

	/*
	�����ı����ѡȡ��Χ
	SelLength: ѡȡ��Χ�ĳ���
	���SelLength��-1���ı����ѡȡ��Χ�ǵ�ǰ�������λ�õ��ı�ĩβ��
	*/
	void SetSelLength(DWORD SelLength)
	{
		DWORD CurrPos = GetSelStart();
		SendMessage(CurrentHwnd, EM_SETSEL, (WPARAM)CurrPos, (LPARAM)(CurrPos + SelLength));
	}

	/*
	��ȡָ���еĳ���
	lnNumber: �кš����Ϊ-1�򷵻�δѡȡ���ı����ȡ�
	*/
	long GetLineLength(long lnNumber)
	{
		return (long)SendMessage(CurrentHwnd, EM_LINELENGTH, lnNumber, 0);
	}

	/*
	ˮƽ����ָ������������
	CharCount: ˮƽ����������
	����ı����ǵ����ı����򷵻�false������Ƕ����ı����򷵻�true
	*/
	bool HScroll(long CharCount)
	{
		return (bool)SendMessage(CurrentHwnd, EM_LINESCROLL, (WPARAM)CharCount, 0);
	}

	/*
	��ֱ����ָ������������
	LineCount: ��ֱ����������
	����ı����ǵ����ı����򷵻�false������Ƕ����ı����򷵻�true
	*/
	bool VScroll(long LineCount)
	{
		return (bool)SendMessage(CurrentHwnd, EM_LINESCROLL, 0, (LPARAM)LineCount);
	}

	/*
	�ѵ�ǰѡ����ı��滻��ָ�����ַ���
	bCanUndo: �˲����Ƿ�ɳ���
	NewText: �µ��ı�
	*/
	void SetSelText(bool bCanUndo, char* NewText)
	{
		SendMessage(CurrentHwnd, EM_REPLACESEL, (WPARAM)bCanUndo, (LPARAM)NewText);
	}

	/* ��ȡ��ǰ������ڵ��� */
	long GetCurrLine()
	{
		DWORD CurrSel = SendMessage(CurrentHwnd, EM_GETSEL, 0, 0);
		return (long)SendMessage(CurrentHwnd, EM_LINEFROMCHAR, (WPARAM)(CurrSel / 65536), 0);
	}

	/* ��ȡ��ǰ������ڵ��� */
	long GetCurrCol()
	{
		DWORD CurrSel = SendMessage(CurrentHwnd, EM_GETSEL, 0, 0);
		DWORD CurrLnSel = SendMessage(CurrentHwnd, EM_LINEINDEX, (WPARAM)-1, 0);
		return ((CurrSel / 65536) - CurrLnSel + 1);
	}

	/* ���Ʋ��� */
	void Copy()
	{
		SendMessage(CurrentHwnd, WM_COPY, 0, 0);
	}

	/* ���в��� */
	void Cut()
	{
		SendMessage(CurrentHwnd, WM_CUT, 0, 0);
	}

	/* �������� */
	void Undo()
	{
		SendMessage(CurrentHwnd, EM_UNDO, 0, 0);
	}

	/* ���ѡ����ı� */
	void Clear()
	{
		SendMessage(CurrentHwnd, WM_CLEAR, 0, 0);
	}

	/* ճ������ */
	void Paste()
	{
		SendMessage(CurrentHwnd, WM_PASTE, 0, 0);
	}

	/* ��ȡ�ı����� */
	long GetTextLength()
	{
		return (long)SendMessage(CurrentHwnd, WM_GETTEXTLENGTH, 0, 0);
	}

	/* ��������� */
	void ScrollToCaret()
	{
		SendMessage(CurrentHwnd, EM_SCROLLCARET, 0, 0);
	}

	/* ��ȡ�ı� */
	char* GetText()
	{
		int Length = GetWindowTextLength(CurrentHwnd) + 1;
		char* tmp = new char[Length];
		GetWindowText(CurrentHwnd, tmp, Length);
		return tmp;
	}

	/* ʹ�ı����ȡ���� */
	void SetFocus()
	{
		::SetFocus(CurrentHwnd);
	}
};

//============================================================================
class MyFrame : public ButtonPublicClass
{
public:
	//�������ؼ�
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD | BS_GROUPBOX;					//��ʼ���ؼ���ʽ
		lStyle |= TextPos;													//����ı�λ�õ���ʽ��

		//�����ؼ�
		CurrentHwnd = CreateWindowEx(0, "MyButton", Text, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}
};

//============================================================================
class MyButton : public ButtonPublicClass
{
public:
	bool		ClientEdgeBorder;				//����߿�
	bool		Flat;							//��ƽ
	bool		BlackBorder;					//��ɫ�߿�

	//������ǰ�İ�ť�ؼ�
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;								//��ʼ���ؼ���ʽ
		LONG ExStyle = 0;													//��չ��ʽ

		//����ؼ�����ʽ
		lStyle |= TextPos;													//����ı�λ�õ���ʽ��
		OrCalc(ClientEdgeBorder, &ExStyle, WS_EX_CLIENTEDGE);				//����߿�
		OrCalc(Flat, &lStyle, BS_FLAT);										//��ƽ
		OrCalc(BlackBorder, &lStyle, WS_BORDER);							//��ɫ�߿�

		//�����ؼ�
		CurrentHwnd = CreateWindowEx(ExStyle, "MyButton", Text, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}
};

//============================================================================
class MyCheckBox : public ButtonPublicClass
{
public:
	bool		ClientEdgeBorder;				//����߿�
	bool		Flat;							//��ƽ
	bool		BlackBorder;					//��ɫ�߿�
	bool		PushLike;						//��ť��ʽ

	//������ѡ��ؼ�
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX;				//��ʼ���ؼ���ʽ
		LONG ExStyle = 0;													//��չ��ʽ

		//����ؼ�����ʽ
		lStyle |= TextPos;													//����ı�λ�õ���ʽ��
		OrCalc(ClientEdgeBorder, &ExStyle, WS_EX_CLIENTEDGE);				//����߿�
		OrCalc(Flat, &lStyle, BS_FLAT);										//��ƽ
		OrCalc(BlackBorder, &lStyle, WS_BORDER);							//��ɫ�߿�
		OrCalc(PushLike, &lStyle, BS_PUSHLIKE);								//��ť��ʽ

		//�����ؼ�
		CurrentHwnd = CreateWindowEx(ExStyle, "MyButton", Text, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/* ���ø�ѡ��Ĺ�ѡ״̬ */
	void SetChecked(bool bChecked)
	{
		SendMessage(CurrentHwnd, BM_SETCHECK, bChecked ? BST_CHECKED : BST_UNCHECKED, 0);
	}

	/* ��ȡ��ѡ��Ĺ�ѡ״̬ */
	bool GetChecked()
	{
		return (SendMessage(CurrentHwnd, BM_GETCHECK, 0, 0) == BST_CHECKED);
	}
};

//============================================================================
class MyOption : public MyCheckBox
{
public:
	//������ѡ��ؼ�
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD | BS_AUTORADIOBUTTON;			//��ʼ���ؼ���ʽ
		LONG ExStyle = 0;													//��չ��ʽ

		//����ؼ�����ʽ
		lStyle |= TextPos;													//����ı�λ�õ���ʽ��
		OrCalc(ClientEdgeBorder, &ExStyle, WS_EX_CLIENTEDGE);				//����߿�
		OrCalc(Flat, &lStyle, BS_FLAT);										//��ƽ
		OrCalc(BlackBorder, &lStyle, WS_BORDER);							//��ɫ�߿�
		OrCalc(PushLike, &lStyle, BS_PUSHLIKE);								//��ť��ʽ

		//�����ؼ�
		CurrentHwnd = CreateWindowEx(ExStyle, "MyButton", Text, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}
};

//============================================================================
class MyCombo : public MyControls
{
public:
	DWORD		VerticalScrollBar;				//��ֱ������
	bool		AutoHscroll;					//�Զ�ˮƽ����
	bool		ForceLowerCase;					//ǿ��Сд
	bool		ForceUppercase;					//ǿ�ƴ�д
	bool		DropDownStyle;					//�б���ʽ
	bool		AutoSort;						//�Զ�����

	//������Ͽ�ؼ�
	HWND Create()
	{
		LONG lStyle = WS_CHILD | CBS_HASSTRINGS | WS_VISIBLE;				//��ʼ���ؼ���ʽ

		lStyle |= VerticalScrollBar;
		OrCalc(AutoHscroll, &lStyle, CBS_AUTOHSCROLL);
		OrCalc(ForceLowerCase, &lStyle, CBS_LOWERCASE);
		OrCalc(ForceUppercase, &lStyle, CBS_UPPERCASE);
		OrCalc(!DropDownStyle, &lStyle, CBS_DROPDOWN);
		OrCalc(AutoSort, &lStyle, CBS_SORT);

		CurrentHwnd = CreateWindowEx(0, "MyCombo", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	���ָ�����ı�����Ͽ��е�ָ��λ��
	strAdd: ��Ҫ��ӵ��ı�
	ListIndex: �ı���ӵ���λ�á������-1 ����ӵ���Ͽ��ĩβ
	�����ӳɹ��򷵻ض�Ӧ���б���ţ����ʧ���򷵻�CB_ERR (-1)
	*/
	int AddItem(char* strAdd, int ListIndex = -1)
	{
		return (int)SendMessage(CurrentHwnd, CB_INSERTSTRING, (WPARAM)ListIndex, (LPARAM)strAdd);
	}

	/*
	���ָ������ֵ����Ͽ��е�ָ��λ��
	strAdd: ��Ҫ��ӵ���ֵ
	ListIndex: �ı���ӵ���λ�á������-1 ����ӵ���Ͽ��ĩβ
	�����ӳɹ��򷵻ض�Ӧ���б���ţ����ʧ���򷵻�CB_ERR (-1)
	*/
	int AddItem(int strAdd, int ListIndex = -1)
	{
		char Buffer[255];
		itoa(strAdd, Buffer, 10);
		return (int)SendMessage(CurrentHwnd, CB_INSERTSTRING, (WPARAM)ListIndex, (LPARAM)Buffer);
	}

	/*
	ɾ��ָ����ŵ��б���
	ListIndex: ָ�����б������
	���ִ�гɹ��򷵻��б���ʣ�����Ŀ�������ʧ���򷵻�CB_ERR (-1)
	*/
	int RemoveItem(int ListIndex)
	{
		return (int)SendMessage(CurrentHwnd, CB_DELETESTRING, (WPARAM)ListIndex, 0);
	}

	/* ��������б��� */
	void Clear()
	{
		SendMessage(CurrentHwnd, CB_RESETCONTENT, 0, 0);
	}

	/*
	����Ͽ����г�ָ��Ŀ¼���ļ�
	FilePath: �ļ�·��
	FileType: ɸѡ���ļ����ͣ�ΪDDL_*�ĳ���
	���ִ�гɹ��򷵻����һ����ӵ��б������ţ����ʧ���򷵻�CB_ERR (-1)������б�ռ䲻���򷵻�CB_ERRSPACE (-2)
	*/
	int ListDirFiles(LPCTSTR FilePath, DWORD FileType = 0)
	{
		return (int)SendMessage(CurrentHwnd, CB_DIR, (WPARAM)FileType, (LPARAM)FilePath);
	}

	/*
	����Ͽ��в�����ָ���ı����б���
	StrFind: ��Ҫ���ҵ��ı�
	StartIndex: ���ĸ��б��ʼ���ҡ����Ϊ-1���ͷ��ʼ����
	FullMatch: �Ƿ���Ҫȫ��ƥ�䡣������ǣ������ı��Ŀ�ͷ�Ƿ�������ҵ��ı�
	���ִ�гɹ��򷵻��ҵ����б������ţ����ʧ���򷵻�CB_ERR (-1)
	*/
	int FindItem(char* strFind, int StartIndex = -1, bool FullMatch = true)
	{
		return (int)SendMessage(CurrentHwnd, FullMatch ? CB_FINDSTRINGEXACT : CB_FINDSTRING,
			(WPARAM)StartIndex, (LPARAM)strFind);
	}

	/* ��ȡ��Ͽ��е���Ŀ���� */
	int GetListCount()
	{
		return (int)SendMessage(CurrentHwnd, CB_GETCOUNT, 0, 0);
	}

	/*
	��ȡ��Ͽ��е�ǰѡȡ���б���
	����ֵΪ��ǰѡȡ���б�����š����û���б��ѡȡ�򷵻�CB_ERR (-1)
	*/
	int GetSelItem()
	{
		return (int)SendMessage(CurrentHwnd, CB_GETCURSEL, 0, 0);
	}

	/*
	��ȡ��Ͽ���б������״̬
	��������б�����򷵻�true
	*/
	bool GetDroppedState()
	{
		return (bool)SendMessage(CurrentHwnd, CB_GETDROPPEDSTATE, 0, 0);
	}

	/* ��ȡ��Ͽ�������б�Ŀ�� */
	int GetDroppedWidth()
	{
		return (int)SendMessage(CurrentHwnd, CB_GETDROPPEDWIDTH, 0, 0);
	}

	/*
	��ȡ��Ͽ��е��ı����ѡȡ��Χ
	outSelStart: ������ŵõ����ı�ѡȡ����Ŀ�ͷ
	outSelEnd: ������ŵõ����ı�ѡȡ����Ľ�β
	*/
	void GetEditSel(int* outSelStart, int* outSelEnd)
	{
		SendMessage(CurrentHwnd, CB_GETEDITSEL, (WPARAM)outSelStart, (LPARAM)outSelEnd);
	}

	/* ��ȡ��Ͽ�������һ���б���ĸ߶� */
	int GetItemHeight()
	{
		return (int)SendMessage(CurrentHwnd, CB_GETITEMHEIGHT, 0, 0);
	}

	/*
	��ȡ��Ͽ���ָ���б�����ı�
	ListIndex: ָ�����б������
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

	/* ��ȡ�ı� */
	char* GetText()
	{
		int Length = GetWindowTextLength(CurrentHwnd) + 1;
		char* tmp = new char[Length];
		GetWindowText(CurrentHwnd, tmp, Length);
		return tmp;
	}

	/*
	�����ı�
	NewText: �µ��ı�
	*/
	void SetText(char* NewText)
	{
		SetWindowText(CurrentHwnd, NewText);
	}

	/*
	�����������ı�
	NewText: ��Ч�����ֱ��ʽ
	*/
	void SetText(int NewText)
	{
		char Buffer[255];
		itoa(NewText, Buffer, 10);
		SetWindowText(CurrentHwnd, Buffer);
	}

	/* ��ȡ��Ͽ�������б������ٿ��Կ������б��� */
	int GetMinimumVisibleItems()
	{
		return (int)SendMessage(CurrentHwnd, CB_GETMINVISIBLE, 0, 0);
	}

	/*
	��ȡ��Ͽ���ӵĵ�һ���б�������
	�������ִ�гɹ��򷵻���Ͽ�������б��еĵ�һ�����ӵ��б�����ţ����ʧ���򷵻�CB_ERR (-1)
	*/
	int GetFirstVisibleItem()
	{
		return (int)SendMessage(CurrentHwnd, CB_GETTOPINDEX, 0, 0);
	}

	/*
	�����ı�������ı�����
	LimitLength: ���Ƶĳ��ȣ����Ϊ0��ΪϵͳĬ��
	*/
	void SetLimitLength(int LimitLength)
	{
		SendMessage(CurrentHwnd, CB_LIMITTEXT, (WPARAM)LimitLength, 0);
	}

	/*
	������Ͽ�ѡ����б���
	ListIndex: �б������ţ����Ϊ-1��ѡ���κ��б���
	���ִ�гɹ��򷵻�ѡȡ���б������ţ����ִ��ʧ���򷵻�CB_ERR (-1)
	*/
	int SetSelItem(int ListIndex)
	{
		return (int)SendMessage(CurrentHwnd, CB_SETCURSEL, ListIndex, 0);
	}

	/*
	������Ͽ�ѡ����б���
	ListText: ��Ҫ��ѡ����б�����ı���ϵͳ�����ı��Ŀ�ͷ�Ƿ�������ҵ��ı�
	StartItem: ��ʼ�������б�������-1���ͷ��ʼ����
	���ִ�гɹ��򷵻�ѡȡ���б������ţ����ִ��ʧ���򷵻�CB_ERR (-1)
	*/
	int SetSelItem(char* ListText, int StartItem = -1)
	{
		return (int)SendMessage(CurrentHwnd, CB_SELECTSTRING, (WPARAM)StartItem, (LPARAM)ListText);
	}

	/*
	������Ͽ�������б�Ŀ��
	NewWidth: �����б��µĿ��
	���ִ�гɹ��򷵻������б��µĿ�ȣ����ִ��ʧ���򷵻�CB_ERR (-1)
	*/
	int SetDroppedWidth(int NewWidth)
	{
		return (int)SendMessage(CurrentHwnd, CB_SETDROPPEDWIDTH, (WPARAM)NewWidth, 0);
	}

	/*
	������Ͽ����ı����ѡȡ��Χ
	SelStart: �ı�ѡȡ�Ŀ�ͷ�����Ϊ-1��ȡ��ѡȡ�ı�
	SelEnd: �ı�ѡȡ�Ľ�β�����Ϊ-1���ı�ѡȡ�ķ�ΧΪSelStart��ĩβ
	���ִ�гɹ��򷵻�TRUE�����ִ��ʧ���򷵻�CB_ERR (-1)
	*/
	bool SetSelText(int SelStart, int SelEnd)
	{
		return (SendMessage(CurrentHwnd, CB_SETEDITSEL, 0, MAKELPARAM(SelStart, SelEnd)) == TRUE);
	}

	/*
	������Ͽ�������б���ÿһ���б���ĸ߶�
	NewHeight: ÿ���б����µĸ߶�
	���ִ�гɹ��򷵻��б����µĸ߶ȣ����ִ��ʧ���򷵻�CB_ERR (-1)
	*/
	int SetItemHeight(int NewHeight)
	{
		return (int)SendMessage(CurrentHwnd, CB_SETITEMHEIGHT, 0, NewHeight);
	}

	/*
	������Ͽ�������б�ÿ��������ʾ���б�����
	ItemCount: ������ʾ���б�����
	���ִ�гɹ��򷵻�true
	*/
	bool SetMinimumVisibleItems(int ItemCount)
	{
		return (SendMessage(CurrentHwnd, CB_SETMINVISIBLE, (WPARAM)ItemCount, 0) == TRUE);
	}

	/*
	ʹָ�����б����ܹ�����ʾ����Ͽ�������б���
	ListIndex: ��Ҫ��ʾ���б���е���Ŀ
	����ɹ��򷵻�true
	*/
	bool ScrollToItem(int ListIndex)
	{
		return (SendMessage(CurrentHwnd, CB_SETTOPINDEX, (WPARAM)ListIndex, 0) == 0);
	}

	/*
	��ʾ����������Ͽ�������б�
	bShow: �Ƿ���ʾ��Ͽ�������б�
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
	DWORD		VerticalScrollBar;				//��ֱ������
	bool		MultiSelect;					//�����ѡ
	bool		MultiColumn;					//�������
	bool		ClientEdgeBorder;				//����߿�
	bool		BlackBorder;					//��ɫ�߿�
	bool		AutoSort;						//�Զ�����

	//�����б��ؼ�
	HWND Create()
	{
		LONG lStyle = WS_CHILD | LBS_HASSTRINGS | LBS_NOINTEGRALHEIGHT | LBS_NOTIFY;		//��ʼ���ؼ���ʽ

		lStyle |= VerticalScrollBar;
		OrCalc(MultiSelect, &lStyle, LBS_EXTENDEDSEL);
		OrCalc(MultiColumn, &lStyle, LBS_MULTICOLUMN);
		OrCalc(BlackBorder, &lStyle, WS_BORDER);
		OrCalc(AutoSort, &lStyle, LBS_SORT);

		CurrentHwnd = CreateWindowEx(ClientEdgeBorder ? WS_EX_CLIENTEDGE : 0, "MyListBox", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	���ָ�����ı����б���е�ָ��λ��
	strAdd: ��Ҫ��ӵ��ı�
	ListIndex: �ı���ӵ���λ�á������-1 ����ӵ��б���ĩβ
	�����ӳɹ��򷵻ض�Ӧ���б���ţ����ʧ���򷵻�LB_ERR (-1)
	*/
	int AddItem(char* strAdd, int ListIndex = -1)
	{
		return (int)SendMessage(CurrentHwnd, LB_INSERTSTRING, (WPARAM)ListIndex, (LPARAM)strAdd);
	}

	/*
	���ָ������ֵ���б���е�ָ��λ��
	strAdd: ��Ҫ��ӵ���ֵ
	ListIndex: �ı���ӵ���λ�á������-1 ����ӵ��б���ĩβ
	�����ӳɹ��򷵻ض�Ӧ���б���ţ����ʧ���򷵻�LB_ERR (-1)
	*/
	int AddItem(int strAdd, int ListIndex = -1)
	{
		char Buffer[255];
		itoa(strAdd, Buffer, 10);
		return (int)SendMessage(CurrentHwnd, LB_INSERTSTRING, (WPARAM)ListIndex, (LPARAM)Buffer);
	}

	/*
	ɾ��ָ����ŵ��б���
	ListIndex: ָ�����б������
	���ִ�гɹ��򷵻��б���ʣ�����Ŀ�������ʧ���򷵻�LB_ERR (-1)
	*/
	int RemoveItem(int ListIndex)
	{
		return (int)SendMessage(CurrentHwnd, LB_DELETESTRING, (WPARAM)ListIndex, 0);
	}

	/* ��������б��� */
	void Clear()
	{
		SendMessage(CurrentHwnd, LB_RESETCONTENT, 0, 0);
	}

	/*
	���б�����г�ָ��Ŀ¼���ļ�
	FilePath: �ļ�·��
	FileType: ɸѡ���ļ����ͣ�ΪDDL_*�ĳ���
	���ִ�гɹ��򷵻����һ����ӵ��б������ţ����ʧ���򷵻�LB_ERR (-1)������б�ռ䲻���򷵻�LB_ERRSPACE (-2)
	*/
	int ListDirFiles(LPCTSTR FilePath, DWORD FileType = 0)
	{
		return (int)SendMessage(CurrentHwnd, LB_DIR, (WPARAM)FileType, (LPARAM)FilePath);
	}

	/*
	���б���в�����ָ���ı����б���
	StrFind: ��Ҫ���ҵ��ı�
	StartIndex: ���ĸ��б��ʼ���ҡ����Ϊ-1���ͷ��ʼ����
	FullMatch: �Ƿ���Ҫȫ��ƥ�䡣������ǣ������ı��Ŀ�ͷ�Ƿ�������ҵ��ı�
	���ִ�гɹ��򷵻��ҵ����б������ţ����ʧ���򷵻�LB_ERR (-1)
	*/
	int FindItem(char* strFind, int StartIndex = -1, bool FullMatch = true)
	{
		return (int)SendMessage(CurrentHwnd, FullMatch ? LB_FINDSTRINGEXACT : LB_FINDSTRING,
			(WPARAM)StartIndex, (LPARAM)strFind);
	}

	/* ��ȡ�б���е���Ŀ���� */
	int GetListCount()
	{
		return (int)SendMessage(CurrentHwnd, LB_GETCOUNT, 0, 0);
	}

	/*
	��ȡ�б���е�ǰѡȡ���б���
	����ֵΪ��ǰѡȡ���б�����š����û���б��ѡȡ�򷵻�LB_ERR (-1)
	*/
	int GetSelItem()
	{
		return (int)SendMessage(CurrentHwnd, LB_GETCURSEL, 0, 0);
	}

	/* ��ȡ�б��������һ���б���ĸ߶� */
	int GetItemHeight()
	{
		return (int)SendMessage(CurrentHwnd, LB_GETITEMHEIGHT, 0, 0);
	}

	/*
	��ȡ�б����ָ���б�����ı�
	ListIndex: ָ�����б������
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
	�����б��ѡ����б���
	ListIndex: �б������ţ����Ϊ-1��ѡ���κ��б���
	bSelect: �Ƿ�ѡ����б����ΪFALSE����ȡ��ѡȡ���б���
	���ִ�гɹ��򷵻�ѡȡ���б������ţ����ִ��ʧ���򷵻�LB_ERR (-1)
	*/
	int SetSelItem(int ListIndex, bool bSelect = TRUE)
	{
		return (int)SendMessage(CurrentHwnd, LB_SETCURSEL, (WPARAM)ListIndex, 0);
	}

	/*
	�����б��������б���ÿһ���б���ĸ߶�
	NewHeight: ÿ���б����µĸ߶�
	���ִ�гɹ��򷵻��б����µĸ߶ȣ����ִ��ʧ���򷵻�LB_ERR (-1)
	*/
	int SetItemHeight(int NewHeight)
	{
		return (int)SendMessage(CurrentHwnd, LB_SETITEMHEIGHT, 0, NewHeight);
	}

	/*
	ʹָ�����б����ܹ�����ʾ���б����
	ListIndex: ��Ҫ��ʾ���б���е���Ŀ
	����ɹ��򷵻�true
	*/
	bool ScrollToItem(int ListIndex)
	{
		return (SendMessage(CurrentHwnd, LB_SETCARETINDEX, (WPARAM)ListIndex, FALSE) == 0);
	}

	/*
	��õ��б������ѡ��ʱ���׸��б���������ڶ�ѡ�б��
	���ִ��ʧ���򷵻�LB_ERR (-1)
	*/
	int GetFirstItem()
	{
		return (int)(SendMessage(CurrentHwnd, LB_GETANCHORINDEX, 0, 0));
	}

	/*
	��õ�ǰ��ѡ���б��������������ڶ�ѡ�б��
	���ִ��ʧ���򷵻�LB_ERR (-1)
	*/
	int GetSelCount()
	{
		return (int)(SendMessage(CurrentHwnd, LB_GETSELCOUNT, 0, 0));
	}

	/*
	��ȡ�б��ǰ��ѡ�������б���������ڶ�ѡ�б��
	intBuffer��ָ��һ��ָ�������ָ�롣�䶨��һ��Ϊ int *intBuffer = new int[�б�����];
	����ɹ�����intBuffer��д�����е��б�����ţ����ִ��ʧ���򷵻�LB_ERR (-1)
	*/
	int GetMultiSelItems(int* intBuffer)
	{
		int SelCount = (int)(SendMessage(CurrentHwnd, LB_GETSELCOUNT, 0, 0));					//�б�����
		int *Buffer = (int*)GlobalAlloc(GPTR, SelCount * sizeof(int));							//Ϊ�����������ڴ�ռ�
		int rtn = SendMessage(CurrentHwnd, LB_GETSELITEMS, (WPARAM)SelCount, (LPARAM)Buffer);	//��ȡѡ����б��������
		memcpy(intBuffer, Buffer, SelCount * sizeof(int));										//�����������ڴ濽����Ŀ���ڴ�
		return (rtn == LB_ERR) ? LB_ERR : SelCount;														//�жϺ����Ƿ�ִ�гɹ�
	}

	/*
	��ȡ�б����ӵĵ�һ���б�������
	�������ִ�гɹ��򷵻��б��ĵ�һ�����ӵ��б�����ţ����ʧ���򷵻�LB_ERR (-1)
	*/
	int GetFirstVisibleItem()
	{
		return (int)SendMessage(CurrentHwnd, LB_GETTOPINDEX, 0, 0);
	}

	/*
	��ȡ�б��ָ�������Ӧ���б���
	X��Y: ָ��������
	���ض�Ӧ���б�������
	*/
	int ItemFromPoint(int X, int Y)
	{
		return LOWORD(SendMessage(CurrentHwnd, LB_ITEMFROMPOINT, 0, MAKELPARAM(X, Y)));
	}

	/*
	����ѡ��ָ�����б���������ڶ�ѡ�б��
	ListIndexFrom: ѡ����б���ķ�Χ���׸�
	ListIndexTo: ѡ����б���ķ�Χ�����һ��
	bSelect: ��ΪTRUE��ѡȡָ����Χ�ڵ��б����ΪFALSE��ѡȡָ����Χ�ڵ��б���
	����д������򷵻�false
	*/
	bool SetSelItemRange(int ListIndexFrom, int ListIndexTo, bool bSelect = TRUE)
	{
		return (SendMessage(CurrentHwnd, LB_SELITEMRANGE, (WPARAM)bSelect,
			MAKELPARAM(ListIndexFrom, ListIndexTo)) != CB_ERR);
	}

	/* �����б���ÿһ�еĿ�ȣ��������ڶ����б�� */
	void SetColumnWidth(int NewWidth)
	{
		SendMessage(CurrentHwnd, LB_SETCOLUMNWIDTH, (WPARAM)NewWidth, 0);
	}

	/* ���б���ý��� */
	void SetFocus()
	{
		::SetFocus(CurrentHwnd);
	}
};

//============================================================================
class MyHScroll : public MyControls
{
public:
	int			Min;							//��Сֵ
	int			Max;							//���ֵ
	int			SmallChange;					//��С����ֵ
	int			LargeChange;					//������ֵ

	//�����������ؼ�
	HWND Create()
	{
		LONG lStyle = WS_CHILD | SBS_HORZ | SBS_LEFTALIGN;

		CurrentHwnd = CreateWindowEx(0, "MyScrollBar", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		//���ù������ķ�Χ
		SetScrollRange(CurrentHwnd, SB_CTL, Min, Max, TRUE);

		return CurrentHwnd;
	}

	//��ù�������ֵ
	int GetValue()
	{
		return GetScrollPos(CurrentHwnd, SB_CTL);
	}

	/*
	���ù�������ֵ
	���ִ�гɹ��򷵻ع�����֮ǰ��ֵ�����ִ��ʧ���򷵻�0
	*/
	int SetValue(int Value)
	{
		return SetScrollPos(CurrentHwnd, SB_CTL, Value, TRUE);
	}

	/*
	��ù�������Χ
	lpMin, lpMax �ֱ�Ϊָ���������չ�������Сֵ�����ֵ��������ָ��
	���ִ��ʧ���򷵻�0
	*/
	int GetRange(int* lpMin, int* lpMax)
	{
		return GetScrollRange(CurrentHwnd, SB_CTL, lpMin, lpMax);
	}

	/*
	���ù�������Χ
	MinValue, MaxValue �ֱ�Ϊ�������µ���Сֵ�����ֵ
	���ִ��ʧ���򷵻�0
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

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		//���ù������ķ�Χ
		SetScrollRange(CurrentHwnd, SB_CTL, Min, Max, TRUE);

		return CurrentHwnd;
	}
};

//============================================================================
class MyUpDown : public MyControls
{
public:
	int			Min;							//��Сֵ
	int			Max;							//���ֵ
	int			Accel;							//�����ٶ�
	bool		HorzStyle;						//�Ƿ�Ϊˮƽ��ʽ

	//�������ڰ�ť�ؼ�
	HWND Create()
	{
		LONG lStyle = WS_CHILD | (HorzStyle ? UDS_HORZ : 0);

		CurrentHwnd = CreateWindowEx(0, "MyUpDown", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		//������Сֵ�����ֵ
		PostMessage(CurrentHwnd, UDM_SETRANGE32, (WPARAM)Min, (LPARAM)Max);

		//���õ����ٶ�
		UDACCEL uda;
		uda.nSec = 1;
		uda.nInc = Accel;
		SendMessage(CurrentHwnd, UDM_SETACCEL, 1, (LPARAM)&uda);

		return CurrentHwnd;
	}

	/* ��ȡ���ڰ�ť�ĵ����ٶ� */
	int GetAccel()
	{
		UDACCEL uda;
		SendMessage(CurrentHwnd, UDM_GETACCEL, 1, (LPARAM)&uda);
		return uda.nInc;
	}

	/*
	���õ��ڰ�ť�ĵ����ٶ�
	Acceleration: �µĵ����ٶ�
	���ִ�гɹ��򷵻�TRUE
	*/
	bool SetAccel(int Acceleration)
	{
		UDACCEL uda;
		uda.nSec = 1;
		uda.nInc = Acceleration;
		return (bool)SendMessage(CurrentHwnd, UDM_SETACCEL, 1, (LPARAM)&uda);
	}

	/* ��ȡ���ڰ�ť��ֵ */
	int GetPos()
	{
		return (int)SendMessage(CurrentHwnd, UDM_GETPOS32, 0, 0);
	}

	/*
	���õ��ڰ�ť��ֵ
	NewPos: �µ�ֵ
	���ص��ڰ�ť֮ǰ��ֵ
	*/
	int SetPos(int NewPos)
	{
		return (int)SendMessage(CurrentHwnd, UDM_SETPOS32, 0, (LPARAM)NewPos);
	}

	/*
	��ȡ���ڰ�ť�ķ�Χ
	lpMin, lpMax �ֱ�Ϊָ���������յ��ڰ�ť��Сֵ�����ֵ��������ָ��
	*/
	void GetRange(int* lpMin, int* lpMax)
	{
		SendMessage(CurrentHwnd, UDM_GETRANGE32, (WPARAM)lpMin, (LPARAM)lpMax);
	}

	/*
	���õ��ڰ�ť�ķ�Χ
	Min, Max �ֱ�Ϊ���ڰ�ť�µ���Сֵ�����ֵ
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
	int			Min;							//��Сֵ
	int			Max;							//���ֵ
	bool		Smooth;							//�Ƿ�ƽ����ʽ
	bool		VertStyle;						//�Ƿ�Ϊ��ֱ��ʽ
	COLORREF	BarColor;						//������ɫ
	COLORREF	BackColor;						//������ɫ

	//�����������ؼ�
	HWND Create()
	{
		LONG lStyle = WS_CHILD | (Smooth ? PBS_SMOOTH : 0) | (VertStyle ? PBS_VERTICAL : 0) | PBS_MARQUEE;

		CurrentHwnd = CreateWindowEx(0, "MyProgressBar", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		//������Сֵ�����ֵ
		PostMessage(CurrentHwnd, PBM_SETRANGE32, (WPARAM)Min, (LPARAM)Max);

		//���û�����ɫ�ͱ�����ɫ
		PostMessage(CurrentHwnd, PBM_SETBARCOLOR, 0, (LPARAM)BarColor);
		PostMessage(CurrentHwnd, PBM_SETBKCOLOR, 0, (LPARAM)BackColor);

		return CurrentHwnd;
	}

	/*
	�����������ָ����ֵ
	ValueAdd: ���������ӵ�ֵ
	���ؽ�����֮ǰ��ֵ
	*/
	int IncreaseValue(int ValueAdd)
	{
		return (int)PostMessage(CurrentHwnd, PBM_DELTAPOS, (WPARAM)ValueAdd, 0);
	}

	/* ��ȡ������ɫ */
	COLORREF GetBarColor()
	{
		return (COLORREF)PostMessage(CurrentHwnd, PBM_GETBARCOLOR, 0, 0);
	}

	/*
	���û������ɫ
	NewColor: �µĻ�����ɫ��
	���ؽ�����֮ǰ�Ļ�����ɫ
	*/
	COLORREF SetBarColor(COLORREF NewColor)
	{
		return (COLORREF)PostMessage(CurrentHwnd, PBM_SETBARCOLOR, 0, (LPARAM)NewColor);
	}

	/* ��ȡ������ɫ */
	COLORREF GetBackColor()
	{
		return (COLORREF)PostMessage(CurrentHwnd, PBM_GETBKCOLOR, 0, 0);
	}

	/*
	���ñ�����ɫ
	NewColor: �µı�����ɫ����ΪCLR_DEFAULT (0xFF000000L)����ʹ��ϵͳĬ�ϵ���ɫ��
	���ؽ�����֮ǰ�ı�����ɫ
	*/
	COLORREF SetBackColor(COLORREF NewColor)
	{
		return (COLORREF)PostMessage(CurrentHwnd, PBM_SETBKCOLOR, 0, (LPARAM)NewColor);
	}

	/* ��ȡ��������ֵ */
	int GetValue()
	{
		return (int)SendMessage(CurrentHwnd, PBM_GETPOS, 0, 0);
	}

	/*
	���ý�������ֵ
	NewValue: �µ�ֵ������ֵ�����������ķ�Χϵͳ�������Ϊ��ӽ�����Чֵ��
	����֮ǰ��ֵ
	*/
	int SetValue(int NewValue)
	{
		return (int)SendMessage(CurrentHwnd, PBM_SETPOS, (WPARAM)NewValue, 0);
	}

	/* ��ȡ����������Сֵ */
	int GetMin()
	{
		return (int)SendMessage(CurrentHwnd, PBM_GETRANGE, (WPARAM)TRUE, 0);
	}

	/* ��ȡ�����������ֵ */
	int GetMax()
	{
		return (int)SendMessage(CurrentHwnd, PBM_GETRANGE, (WPARAM)FALSE, 0);
	}

	/*
	��ȡ�������ķ�Χ
	lpMin, lpMax �ֱ�Ϊָ���������ս�������Сֵ�����ֵ��������ָ��
	*/
	void GetRange(int* lpMin, int* lpMax)
	{
		PBRANGE pr;
		SendMessage(CurrentHwnd, PBM_GETRANGE, 0, (LPARAM)&pr);
		*lpMin = pr.iLow;
		*lpMax = pr.iHigh;
	}

	/*
	���ý�������Χ
	MinValue, MaxValue �ֱ�Ϊ�������µ����ֵ����Сֵ
	����һ��LOWORD��װ��֮ǰ����Сֵ��HIWORD��װ��֮ǰ�����ֵ������
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
	bool		Direction;						//������Ϊtrue��Ϊˮƽ����Ϊfalse��Ϊ��ֱ
	int			MarkPosition;					//�̶�λ�á�0-5�ֱ�Ϊ��ߡ��ұߡ��Ϸ����·������С��޿̶�
	bool		NoBar;							//�Ƿ���ʾ����
	int			TooltipPos;						//���ֱ�ǩλ��0-4�ֱ�Ϊ��ߡ��ұߡ��Ϸ����·��������ֱ�ǩ
	int			TickFreq;						//�̶ȼ��
	int			Min;							//��Сֵ
	int			Max;							//���ֵ
	int			SmallChange;					//���ٸ��Ĳ���
	int			LargeChange;					//���ٸ��Ĳ���
	bool		BlackBorder;					//��ɫ�߿�

	//��������ؼ�
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD | TBS_AUTOTICKS;
		OrCalc(!Direction, &lStyle, TBS_VERT | TBS_DOWNISLEFT);
		OrCalc(NoBar, &lStyle, TBS_NOTHUMB);
		OrCalc(BlackBorder, &lStyle, WS_BORDER);

		//�ж���û�����ֱ�ǩ
		OrCalc(TooltipPos != 4, &lStyle, TBS_TOOLTIPS);

		//�жϿ̶�λ��
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

		//���û�������ֱ�ǩλ��
		SetToolTipPos(TooltipPos);

		//���û���Ŀ̶ȼ��
		SetTickFreq(TickFreq);

		//���û������Сֵ�����ֵ
		SetRange(Max, Min);

		//�������ٸ��Ĳ����Ϳ��ٸ��Ĳ���
		SetSmallChange(SmallChange);
		SetLargeChange(LargeChange);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	���û�������ֱ�ǩλ��
	NewPos: ���ֱ�ǩλ��0-3�ֱ�Ϊ��ߡ��ұߡ��Ϸ����·�
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
	���û���Ŀ̶ȼ��
	NewTickFreq: �µĿ̶ȼ��
	*/
	void SetTickFreq(int NewTickFreq)
	{
		SendMessage(CurrentHwnd, TBM_SETTICFREQ, (WPARAM)NewTickFreq, 0);
	}

	/*
	���û�������ֵ
	NewMax: �µ����ֵ
	*/
	void SetMax(int NewMax)
	{
		SendMessage(CurrentHwnd, TBM_SETRANGEMAX, TRUE, (LPARAM)NewMax);
	}

	/*
	���û������Сֵ
	NewMin: �µ���Сֵ
	*/
	void SetMin(int NewMin)
	{
		SendMessage(CurrentHwnd, TBM_SETRANGEMIN, TRUE, (LPARAM)NewMin);
	}

	/*
	���û���ķ�Χ
	NewMax: �µ����ֵ
	NewMin: �µ���Сֵ
	*/
	void SetRange(int NewMax, int NewMin)
	{
		SendMessage(CurrentHwnd, TBM_SETRANGE, TRUE, MAKELPARAM(NewMin, NewMax));
	}

	/* ��ȡ��������ֵ */
	int GetMax()
	{
		return (int)SendMessage(CurrentHwnd, TBM_GETRANGEMAX, 0, 0);
	}

	/* ��ȡ�������Сֵ */
	int GetMin()
	{
		return (int)SendMessage(CurrentHwnd, TBM_GETRANGEMIN, 0, 0);
	}

	/*
	���û�������ٸ��Ĳ���
	NewSmallChange: �µ����ٸ��Ĳ���
	*/
	void SetSmallChange(int NewSmallChange)
	{
		SendMessage(CurrentHwnd, TBM_SETLINESIZE, TRUE, (LPARAM)NewSmallChange);
	}

	/*
	���û���Ŀ��ٸ��Ĳ���
	NewLargeChange: �µĿ��ٸ��Ĳ���
	*/
	void SetLargeChange(int NewLargeChange)
	{
		SendMessage(CurrentHwnd, TBM_SETPAGESIZE, TRUE, (LPARAM)NewLargeChange);
	}

	/*
	���û���λ��
	NewPos: �µĻ���λ��
	*/
	void SetPos(int NewPos)
	{
		SendMessage(CurrentHwnd, TBM_SETPOS, TRUE, (LPARAM)NewPos);
	}

	/* ��ȡ����λ�� */
	int GetPos()
	{
		return (int)SendMessage(CurrentHwnd, TBM_GETPOS, 0, 0);
	}
};

//============================================================================
class MyHotkey : public MyControls
{
public:
	//�����ȼ��ؼ�
	HWND Create()
	{
		CurrentHwnd = CreateWindowEx(0, "MyHotkey", "", WS_VISIBLE | WS_CHILD,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	�����ȼ��ؼ����ȼ�
	KeyCode: ָ���ļ��̰���
	KeyModifier: ָ���ļ��̹��ܼ���Shift = 1; Ctrl = 2; Alt = 4
	*/
	void SetHotkey(BYTE KeyCode, BYTE KeyModifier)
	{
		SendMessage(CurrentHwnd, HKM_SETHOTKEY, MAKEWPARAM(MAKEWORD(KeyCode, KeyModifier), 0), 0);
	}

	/*
	��ȡ�ȼ��ؼ����ȼ�
	lpKeyCode: �������ռ��̰����ı�����ָ��
	lpKeyModifier: �������ռ��̹��ܼ��ı�����ָ��
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
	int			Style;							//��ʽ��0-3�ֱ�Ϊͼ�ꡢ�б����桢Сͼ��
	int			Sort;							//�Զ�����ģʽ��0-2�ֱ�Ϊ�������ݼ���������
	int			Align;							//�Զ����롣0-2�ֱ�Ϊ����롢���˶��롢�Զ�
	bool		EditableLabel;					//��ǩ�Ƿ�ɱ༭
	bool		MultiSelectItems;				//�Ƿ���Զ�ѡ
	bool		BlackBorder;					//��ɫ�߿�

	//�����б���ͼ�ؼ�
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

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	����б�ͷ���б���ͼ��
	ColumnText: �б�ͷ�ı�
	Width: �б�ͷ���
	Index: �б�ͷ���
	Position: �б�ͷ�ı�����λ�ã�����ΪLVCFMT_*�ĳ���
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
	����б���б���ͼ��
	ItemText: ����ӵ��б�����ı�
	��ִ�гɹ��򷵻��¼ӵ��б������ţ���ʧ���򷵻�-1
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
	�����б�����ı�
	ItemText: �µ��б��ı�
	ListIndex: ��Ҫ���ĵ��б�������
	SubItemIndex: ��Ҫ���ĵ����б�������
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
	��ȡ�б����ı�
	Index: ָ���б�������
	SubItemIndex: ָ�����б�������
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

	/* ��ȡ�б������� */
	int GetItemCount()
	{
		return SendMessage(CurrentHwnd, LVM_GETITEMCOUNT, 0, 0);
	}

	/*
	��ȡָ���б�ͷ���ı�
	Index: ָ���б�ͷ�����
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
	��ȡָ���б�ͷ�Ŀ��
	Index: ָ���б�ͷ�����
	*/
	int GetColumnWidth(int Index)
	{
		LVCOLUMN lvItem = { 0 };
		lvItem.mask = LVCF_WIDTH;
		SendMessage(CurrentHwnd, LVM_GETCOLUMN, (WPARAM)Index, (LPARAM)&lvItem);
		return lvItem.cx;
	}

	/*
	����ָ���б�ͷ���ı�
	Index: ָ���б�ͷ�����
	NewText: �µ��ı�
	��ִ�гɹ��򷵻�TRUE
	*/
	bool SetColumnText(int Index, char* NewText)
	{
		LVCOLUMN lvItem = { 0 };
		lvItem.mask = LVCF_TEXT;
		lvItem.pszText = NewText;
		return (bool)SendMessage(CurrentHwnd, LVM_SETCOLUMN, (WPARAM)Index, (LPARAM)&lvItem);
	}

	/*
	����ָ���б�ͷ�Ŀ��
	Index: �б�ͷ�����
	NewWidth: �µĿ��
	��ִ�гɹ��򷵻�TRUE
	*/
	bool SetColumnWidth(int Index, int NewWidth)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_SETCOLUMNWIDTH, Index, NewWidth);
	}

	/*
	�����ǩ�༭����
	Index: ��Ҫ�༭��ǩ���б�����š���Ϊ-1����ȡ���༭
	��ִ��ʧ�ܣ�����ֵΪ0
	*/
	int EditLabel(int Index)
	{
		return SendMessage(CurrentHwnd, LVM_EDITLABEL, Index, 0);
	}

	/* �˳���ǩ�༭���� */
	void CancelEditLabel()
	{
		SendMessage(CurrentHwnd, LVM_CANCELEDITLABEL, 0, 0);
	}

	/* ɾ�����е��б��� */
	void DeleteAllItems()
	{
		SendMessage(CurrentHwnd, LVM_DELETEALLITEMS, 0, 0);
	}

	/*
	ɾ��ָ�����б�ͷ
	Index: �б�ͷ�����
	*/
	bool DeleteColumn(int Index)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_DELETECOLUMN, Index, 0);
	}

	/*
	ɾ��ָ�����б���
	Index: �б�������
	*/
	bool DeleteItem(int Index)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_DELETEITEM, Index, 0);
	}

	/*
	ȷ��ָ�����б������
	Index: �б�������
	*/
	bool EnsureVisible(int Index)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_ENSUREVISIBLE, Index, TRUE);
	}

	/*
	����ָ�����ı������б���
	ItemText: ָ���б�����ı�
	AllMatch: �ַ����Ƿ������ȫ��ͬ����ΪFALSE�����б��ͷ�������ҵ��ַ���Ҳ�ɽ���
	StartIndex: ��ָ�����б��ʼ���ҡ�Ĭ��Ϊ��ͷ��ʼ����
	�����ҵ����б������š����ʧ���򷵻�-1
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
	����ָ������������б����������ͼ����ʽ
	X: X����ֵ
	Y: Y����ֵ
	�����ҵ����б������š����ʧ���򷵻�-1
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
	�����б���ͼ�ı�����ɫ
	Color: �µ���ɫ����ΪCLR_NONE��ʹ��ϵͳĬ�ϵ���ɫ
	��ִ�гɹ��򷵻�TRUE
	*/
	bool SetBackColor(COLORREF Color)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_SETBKCOLOR, 0, (LPARAM)Color);
	}

	/*
	�����б���ͼ���ı�������ɫ
	Color: �µ���ɫ����ΪCLR_NONE��ʹ��ϵͳĬ�ϵ���ɫ
	��ִ�гɹ��򷵻�TRUE
	*/
	bool SetTextBackColor(COLORREF Color)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_SETTEXTBKCOLOR, 0, (LPARAM)Color);
	}

	/*
	�����б���ͼ���ı���ɫ
	Color: �µ���ɫ
	��ִ�гɹ��򷵻�TRUE
	*/
	bool SetTextColor(COLORREF Color)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_SETTEXTCOLOR, 0, (LPARAM)Color);
	}

	/*
	��ȡ�б���ͼ�ı�����ɫ
	���ر�����ɫֵ
	*/
	COLORREF GetBackColor()
	{
		return (COLORREF)SendMessage(CurrentHwnd, LVM_GETBKCOLOR, 0, 0);
	}

	/*
	��ȡ�б���ͼ���ı�������ɫ
	�����ı�������ɫֵ
	*/
	COLORREF GetTextBackColor()
	{
		return (COLORREF)SendMessage(CurrentHwnd, LVM_GETTEXTBKCOLOR, 0, 0);
	}

	/*
	��ȡ�б���ͼ���ı���ɫ
	�����ı���ɫֵ
	*/
	COLORREF GetTextColor()
	{
		return (COLORREF)SendMessage(CurrentHwnd, LVM_GETTEXTCOLOR, 0, 0);
	}

	/*
	�����б���ͼ
	vScroll: ��ֱ�����ľ���
	hScroll: ˮƽ�����ľ���
	��ִ�гɹ��򷵻�TRUE
	*/
	bool Scroll(int vScroll, int hScroll = 0)
	{
		return (bool)SendMessage(CurrentHwnd, LVM_SCROLL, (WPARAM)hScroll, (LPARAM)vScroll);
	}

	/*
	����ͼ��λ�ã���������ͼ����ͼ��
	ListIndex: �б������
	X, Y: ͼ�������
	*/
	void SetItemPosition(int ListIndex, int X, int Y)
	{
		POINT point;
		point.x = X;
		point.y = Y;
		SendMessage(CurrentHwnd, LVM_SETITEMPOSITION32, (WPARAM)ListIndex, (LPARAM)&point);
	}

	/*
	��ȡͼ��λ��
	ListIndex: �б������
	pX, pY: ָ����������ͼ�������ָ��
	����ֵ�����ɹ��򷵻�TRUE
	*/
	bool GetItemPosotion(int ListIndex, long* pX, long* pY)
	{
		POINT point;
		bool rtn = (bool)SendMessage(CurrentHwnd, LVM_GETITEMPOSITION, ListIndex, (LPARAM)&point);
		*pX = point.x;
		*pY = point.y;
		return rtn;
	}

	/* ��ȡ��ǰѡ����б�ͷ */
	int GetSelectedColumn()
	{
		return SendMessage(CurrentHwnd, LVM_GETSELECTEDCOLUMN, 0, 0);
	}

	/*
	��ȡ��˵��б�����������б���ͼ
	��ִ�гɹ��򷵻���˵��б������ţ����򷵻�0
	*/
	int GetTopIndex()
	{
		return SendMessage(CurrentHwnd, LVM_GETTOPINDEX, 0, 0);
	}

	/*
	����ListView����ͼ
	Style: ��ʽ��0-3�ֱ�Ϊͼ�ꡢ�б����桢Сͼ��
	����ֵ����ִ�гɹ����򷵻�1�����򷵻�-1
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

	/* ��ȡ��ǰѡ����б�����š���û��ѡ���б���򷵻�-1 */
	int GetSelectedItem()
	{
		return SendMessage(CurrentHwnd, LVM_GETNEXTITEM, (WPARAM)-1, (LPARAM)LVNI_SELECTED);
	}
};

//============================================================================
class MyTreeView : public MyControls
{
public:
	bool		EditableLabels;					//��ǩ�Ƿ�ɱ༭
	bool		HasButtons;						//�Ƿ���ʾ�ڵ㰴ť
	bool		RootHasButtons;					//���ڵ��Ƿ���ʾ��ť
	bool		HasLines;						//�Ƿ���ʾ����
	bool		NoHscroll;						//�Ƿ��ֹˮƽ����
	bool		NoVHscroll;						//�Ƿ��ֹˮƽ�ʹ�ֱ����
	bool		ShowSelAlways;					//ʧ��ʱ�Ƿ���ʾѡ����
	bool		HotTracking;					//�Ƿ�ʵʱѡȡ
	bool		CheckBoxes;						//�Ƿ��ж�ѡ��
	bool		BlackBorder;					//�Ƿ��к�ɫ�߿�

	//��������ͼ�ؼ�
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

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	�����Ŀ������ͼ��
	ItemText: ��Ŀ���ı�
	Parent: ���ڵ�ľ��
	����ֵ: ��������Ŀ�ľ��
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
	ɾ��ָ������Ŀ
	Item: ��Ҫɾ����Ŀ�ľ��������Ҫɾ�����е���Ŀ�������������ΪNULL
	����ֵ: ��ɾ���ɹ��򷵻�TRUE�����򷵻�FALSE
	*/
	bool RemoveItem(HTREEITEM Item)
	{
		return (bool)SendMessage(CurrentHwnd, TVM_DELETEITEM, 0, (LPARAM)Item);
	}

	/*
	ȷ��ָ������Ŀ����
	Item: ָ������Ŀ�ľ��
	*/
	void EnsureVisible(HTREEITEM Item)
	{
		SendMessage(CurrentHwnd, TVM_ENSUREVISIBLE, 0, (LPARAM)Item);
	}

	/*
	չ������������״ͼ
	Item: ��Ҫ��չ�������������б���
	Mode: չ������������1: ������2: չ����3: �л�չ����������
	*/
	bool ExpandItems(HTREEITEM Item, int Mode)
	{
		return (bool)(SendMessage(CurrentHwnd, TVM_EXPAND, (WPARAM)Mode, (LPARAM)Item) != 0);
	}

	/*
	��ʼ�༭�ı�
	Item: ��Ҫ�༭�ı�����Ŀ�ľ��
	*/
	bool EditLabel(HTREEITEM Item)
	{
		return (bool)(SendMessage(CurrentHwnd, TVM_EDITLABEL, 0, (LPARAM)Item) != 0);
	}

	/*
	ȡ���༭�ı�
	SaveChanges: �Ƿ񱣴����Ŀ���޸�
	����ֵ: ��ִ�гɹ�����TRUE�����򷵻�FALSE
	*/
	bool EndEditLabel(bool SaveChanges)
	{
		return (bool)SendMessage(CurrentHwnd, TVM_ENDEDITLABELNOW, (WPARAM)SaveChanges, 0);
	}

	/*
	��ȡ�ı���ɫ
	����ֵ: ��ɫֵ����Ϊ-1�������ʹ��ϵͳĬ����ɫ
	*/
	COLORREF GetTextColor()
	{
		return (COLORREF)SendMessage(CurrentHwnd, TVM_GETTEXTCOLOR, 0, 0);
	}

	/*
	��ȡ������ɫ
	����ֵ: ��ɫֵ����Ϊ-1�������ʹ��ϵͳĬ����ɫ
	*/
	COLORREF GetLineColor()
	{
		return (COLORREF)SendMessage(CurrentHwnd, TVM_GETLINECOLOR, 0, 0);
	}

	/*
	��ȡ������ɫ
	����ֵ: ��ɫֵ����Ϊ-1�������ʹ��ϵͳĬ����ɫ
	*/
	COLORREF GetBackColor()
	{
		return (COLORREF)SendMessage(CurrentHwnd, TVM_GETBKCOLOR, 0, 0);
	}

	/*
	�����ı���ɫ
	Color: �µ���ɫ����Ϊ-1�������ʹ��ϵͳĬ����ɫ
	����ֵ: ֮ǰ����ɫֵ����Ϊ-1�������ʹ��ϵͳĬ����ɫ
	*/
	COLORREF SetTextColor(COLORREF Color)
	{
		return (COLORREF)SendMessage(CurrentHwnd, TVM_SETTEXTCOLOR, 0, (LPARAM)Color);
	}

	/*
	���ñ�����ɫ
	Color: �µ���ɫ����Ϊ-1�������ʹ��ϵͳĬ����ɫ
	����ֵ: ֮ǰ����ɫֵ����Ϊ-1�������ʹ��ϵͳĬ����ɫ
	*/
	COLORREF SetBackColor(COLORREF Color)
	{
		return (COLORREF)SendMessage(CurrentHwnd, TVM_SETBKCOLOR, 0, (LPARAM)Color);
	}

	/*
	����������ɫ
	Color: �µ���ɫ����Ϊ-1�������ʹ��ϵͳĬ����ɫ
	����ֵ: ֮ǰ����ɫֵ����Ϊ-1�������ʹ��ϵͳĬ����ɫ
	*/
	COLORREF SetLineColor(COLORREF Color)
	{
		return (COLORREF)SendMessage(CurrentHwnd, TVM_SETLINECOLOR, 0, (LPARAM)Color);
	}

	/* ��ȡ�б�������� */
	int GetListCount()
	{
		return (int)SendMessage(CurrentHwnd, TVM_GETCOUNT, 0, 0);
	}

	/* ��ȡ�б���ĸ߶� */
	int GetItemHeight()
	{
		return (int)SendMessage(CurrentHwnd, TVM_GETITEMHEIGHT, 0, 0);
	}

	/*
	�����б���ĸ߶�
	Height: �µĸ߶ȡ���Ϊ-1��ʹ��ϵͳĬ�ϸ߶�
	����ֵ: ֮ǰ���б���߶�
	*/
	int SetItemHeight(int Height)
	{
		return (int)SendMessage(CurrentHwnd, TVM_SETITEMHEIGHT, (WPARAM)Height, 0);
	}

	/*
	��ȡָ���б�����ı�
	Item: �б���ľ��
	����ֵ: ָ���б�����ı�
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
	����ָ���б�����ı�
	Item: �б���ľ��
	NewText: �µ��ı�
	��ִ�гɹ��򷵻�TRUE
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

	/* ��ȡ��ǰѡ�����Ŀ�������û��ѡ����Ŀ��ѡ�����Ŀ��Ч���򷵻�0 */
	HTREEITEM GetSelectedItem()
	{
		return (HTREEITEM)SendMessage(CurrentHwnd, TVM_GETNEXTITEM, TVGN_CARET, 0);
	}

	/*
	��ȡָ���б���ĸ��ڵ���
	Item: ָ�����б���ľ��
	��û��ѡ����Ŀ��ѡ�����Ŀ��Ч���򷵻�0
	*/
	HTREEITEM GetParentItem(HTREEITEM Item)
	{
		return (HTREEITEM)SendMessage(CurrentHwnd, TVM_GETNEXTITEM, TVGN_PARENT, (LPARAM)Item);
	}

	/*
	��ȡָ���б������һ�����ӵ��б���ľ��
	Item: ָ�����б���ľ��
	��û��ѡ����Ŀ��ѡ�����Ŀ��Ч���򷵻�0
	*/
	HTREEITEM GetPreviousItem(HTREEITEM Item)
	{
		return (HTREEITEM)SendMessage(CurrentHwnd, TVM_GETNEXTITEM, TVGN_LASTVISIBLE, (LPARAM)Item);
	}

	/*
	��ȡָ���б������һ�����ӵ��б���ľ��
	Item: ָ�����б���ľ��
	��û��ѡ����Ŀ��ѡ�����Ŀ��Ч���򷵻�0
	*/
	HTREEITEM GetNextItem(HTREEITEM Item)
	{
		return (HTREEITEM)SendMessage(CurrentHwnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, (LPARAM)Item);
	}

	/*
	ѡ��ָ�����б���Ŀ
	Item: ָ�����б���ľ������Ҫȡ��ѡ����ΪNULL
	��ִ�гɹ�����TRUE
	*/
	bool SelectItem(HTREEITEM Item)
	{
		return (bool)SendMessage(CurrentHwnd, TVM_SELECTITEM, TVGN_CARET, (LPARAM)Item);
	}

	/* ��ȡ��ǰ���ӵ���Ŀ������ */
	int GetVisibleCount()
	{
		return (int)SendMessage(CurrentHwnd, TVM_GETVISIBLECOUNT, 0, 0);
	}

	/*
	�����б��������С
	NewWidth: �µ�������С����С��ϵͳ��Сֵ������Ϊϵͳ��Сֵ
	*/
	void SetIndent(int NewWidth)
	{
		SendMessage(CurrentHwnd, TVM_SETINDENT, (WPARAM)NewWidth, 0);
	}

	/* ��ȡ�б��������С */
	int GetIndent()
	{
		return (int)SendMessage(CurrentHwnd, TVM_GETINDENT, 0, 0);
	}
};

//============================================================================
class MyTab : public MyControls
{
public:
	bool BottomTabs;							//ѡ��ڵײ�
	bool ButtonLike;							//��ť��ʽ
	bool FlatButtons;							//��ƽ��ť
	bool FixedWidth;							//ѡ�ͳһ��С
	bool FocusOnButtons;						//��ť��ʾ����
	bool ForceLabelLeft;						//�ı������
	bool HotTracking;							//ʵʱѡȡ
	bool MultiLine;								//����ѡ�
	bool ScrollOpposite;						//ѡ��Զ�����
	bool Vertical;								//��ֱ��ʽ

	//����ѡ��ؼ�
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

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);
		return CurrentHwnd;
	}

	//ɾ�����е�ѡ�
	bool DeleteAllItems()
	{
		return (bool)SendMessage(CurrentHwnd, TCM_DELETEALLITEMS, 0, 0);
	}

	/*
	ɾ��ָ����ѡ�
	Index: ָ����ѡ����
	*/
	bool DeleteItem(int Index)
	{
		return (bool)SendMessage(CurrentHwnd, TCM_DELETEITEM, Index, 0);
	}

	/*
	ȡ��ѡ��ѡ�
	ResetAll: ��ΪFALSE��ȡ��ѡȡ���еı�ǩ����ΪTRUE��ȡ��ѡȡ��ǰѡ���ǩ֮��ı�ǩ
	*/
	void DeselectAll(bool ExceptCurrSel)
	{
		SendMessage(CurrentHwnd, TCM_DESELECTALL, (WPARAM)ExceptCurrSel, 0);
	}

	/*
	��ȡ��ǰѡ���ѡ�
	����ֵ: ��ǰѡ���ѡ���š���û��ѡ���򷵻�-1
	*/
	int GetSel()
	{
		return (int)SendMessage(CurrentHwnd, TCM_GETCURSEL, 0, 0);
	}

	/*
	��ȡָ����ǩ���ı�
	����ֵ: ��ִ�гɹ�����TRUE�����򷵻�FALSE
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
	��ȡ��ǩ����
	����ֵ: ��ǩ����
	*/
	int GetItemCount()
	{
		return (int)SendMessage(CurrentHwnd, TCM_GETITEMCOUNT, 0, 0);
	}

	/*
	��ȡһ�������б�ǩ
	����ֵ: ��ǩ������
	*/
	int GetRowCount()
	{
		return (int)SendMessage(CurrentHwnd, TCM_GETROWCOUNT, 0, 0);
	}

	/*
	����ָ����ǩ
	Index: ָ����ǩ�����
	HighLight: �Ƿ����ָ����ǩ����ΪTRUE���������ǩ����ΪFALSE����ָ���ǩ
	��ִ�гɹ�������TRUE�����򷵻�FALSE
	*/
	bool HighLightItem(int Index, bool HighLight)
	{
		return (bool)SendMessage(CurrentHwnd, TCM_HIGHLIGHTITEM, Index, MAKELPARAM(HighLight, 0));
	}

	/*
	��ȡָ�����괦�ı�ǩ���
	���ػ�ȡ���ı�ǩ��š���û��ƥ��ı�ǩ������-1��
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
	��ӱ�ǩ���ؼ���
	Text: ��ǩ���ı�
	Index: ��ǩ�����
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
	���û�ý���ı�ǩ
	Index: ��Ҫ��ý���ı�ǩ
	*/
	void SetFocusIndex(int Index)
	{
		SendMessage(CurrentHwnd, TCM_SETCURFOCUS, Index, 0);
	}

	/*
	���õ�ǰѡ��ı�ǩ
	Index: ��ǩ�����
	����֮ǰѡ��ı�ǩ�����
	*/
	int SetCurrIndex(int Index)
	{
		return (int)SendMessage(CurrentHwnd, TCM_SETCURSEL, Index, 0);
	}

	/*
	���ñ�ǩ�ı�
	Index: ��ǩ�����
	Text: �µ��ı�
	��ִ�гɹ�������TRUE�����򷵻�FALSE
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
	���ñ�ǩ�Ĵ�С
	Width: ��ǩ�Ŀ��
	Height: ��ǩ�ĸ߶�
	����ֵ: ��λ��֮ǰ�Ŀ�ȣ���λ��֮ǰ�ĸ߶�
	*/
	int SetItemSize(int Width, int Height)
	{
		return (int)SendMessage(CurrentHwnd, TCM_SETITEMSIZE, 0, MAKELPARAM(Width, Height));
	}

	/*
	���ñ�ǩ��С�Ĵ�С
	Width: ��ǩ��С�Ĵ�С����Ϊ-1����ʹ��ϵͳĬ�ϴ�С
	����ֵ: ��ǩ֮ǰ����С��С
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
	bool		AutoPlay;						//�Զ�����
	bool		Center;							//��Ƶ���в���
	bool		Transparent;					//��Ƶ����͸��
	bool		ClientEdge;						//����߿�
	bool		BlackBorder;					//��ɫ�߿�

	//���������ؼ�
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

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	ֹͣ����
	��ִ�гɹ�����true
	���ɹ�������true
	*/
	bool Stop()
	{
		return SendMessage(CurrentHwnd, ACM_STOP, 0, 0);
	}

	/*
	��ʼ����
	ReplayTimes: �ظ����ŵĴ�������Ϊ-1�����һֱ�ظ�
	FrameBegin: ��ָ����֡��ʼ���š���Ϊ0�����ͷ��ʼ
	FrameEnd: ���ŵ�ָ����֡����Ϊ-1�����ŵ�ĩβ
	���ɹ�������true
	*/
	bool Play(int ReplayTimes = -1, short FrameBegin = 0, short FrameEnd = -1)
	{
		return SendMessage(CurrentHwnd, ACM_PLAY, ReplayTimes, MAKELPARAM(FrameBegin, FrameEnd));
	}

	/*
	���ļ������ļ�
	FileName: ��Ƶ�ļ�·��
	���ɹ�������true
	*/
	bool Open(char* FileName)
	{
		return SendMessage(CurrentHwnd, ACM_OPEN, 0, (LPARAM)FileName);
	}

	/*
	����Դ�ļ����ļ�
	ResID: ��Դ�ļ���
	���ɹ�������true
	*/
	bool Open(WORD ResID)
	{
		return SendMessage(CurrentHwnd, ACM_OPEN, 0, (LPARAM)MAKEINTRESOURCE(ResID));
	}

	//��ȡ����״̬
	bool IsPlaying()
	{
		return SendMessage(CurrentHwnd, ACM_ISPLAYING, 0, 0);
	}
};

//============================================================================
class MyRichEdit : public MyControls
{
public:
	char*		Text;							//�ı�
	bool		AutoHScroll;					//�Զ�ˮƽ����
	bool		AutoVScroll;					//�Զ���ֱ����
	int			TextPos;						//�ı�λ��
	bool		ForceNumber;					//ǿ������
	bool		IsPassword;						//�����ı�
	bool		ReadOnly;						//�ı�ֻ��
	bool		BlackBorder;					//��ɫ�߿�
	bool		ClientEdgeBorder;				//����߿�
	bool		SunkenBorder;					//�³��ı߿�
	bool		Multiline;						//�����ı�
	int			ScrollBars;						//������
	bool		DisableNoScroll;				//��ʾ���õĹ�����
	bool		NoIME;							//�������뷨
	bool		SelectionBar;					//���Ե�հ�

	//����RTF�ı���ؼ�
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;					//��ʼ���ؼ���ʽ

		//����ؼ�����ʽ
		OrCalc(AutoHScroll, &lStyle, ES_AUTOHSCROLL);			//�Զ�ˮƽ����
		OrCalc(AutoVScroll, &lStyle, ES_AUTOVSCROLL);			//�Զ���ֱ����
		switch (TextPos)										//�ı�λ��
		{
		case 1:														//��
			lStyle |= ES_CENTER;
			break;

		case 2:														//��
			lStyle |= ES_RIGHT;
			break;
		}
		OrCalc(ForceNumber, &lStyle, ES_NUMBER);				//ǿ������
		OrCalc(IsPassword, &lStyle, ES_PASSWORD);				//�����ı�
		OrCalc(ReadOnly, &lStyle, ES_READONLY);					//�ı�ֻ��
		OrCalc(BlackBorder, &lStyle, WS_BORDER);				//��ɫ�߿�
		OrCalc(SunkenBorder, &lStyle, ES_SUNKEN);				//�³��ı߿�
		OrCalc(Multiline, &lStyle, ES_MULTILINE);				//�����ı�
		switch (ScrollBars)										//������
		{
		case 1:														//ˮƽ
			lStyle |= WS_HSCROLL;
			break;

		case 2:														//��ֱ
			lStyle |= WS_VSCROLL;
			break;

		case 3:														//��������
			lStyle |= WS_HSCROLL | WS_VSCROLL;
			break;
		}
		OrCalc(DisableNoScroll, &lStyle, ES_DISABLENOSCROLL);		//��ʾ���õĹ�����
		OrCalc(NoIME, &lStyle, ES_NOIME);							//�������뷨

		//�����ؼ�
		CurrentHwnd = CreateWindowEx(ClientEdgeBorder ? WS_EX_CLIENTEDGE : 0, "MyRichEdit", Text, lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		//�ж��Ƿ�Ϊ�ؼ��������Ե�հ���ʽ
		OrCalc(SelectionBar, &lStyle, ES_SELECTIONBAR);
		SetWindowLong(CurrentHwnd, GWL_STYLE, lStyle);

		return CurrentHwnd;
	}

	/*
	�����ı����Ƿ��Զ����URL
	Enabled: �Ƿ��Զ����URL
	��ִ�гɹ�����������true
	*/
	bool SetUrlDetect(bool Enabled)
	{
		return (!SendMessage(CurrentHwnd, EM_AUTOURLDETECT,
			Enabled ? AURL_ENABLEURL : 0, 0));
	}

	/*
	��⵱ǰ�ı����Ƿ��ܽ���ճ������
	���ܽ��в���������true
	*/
	bool CanPaste()
	{
		return (!SendMessage(CurrentHwnd, EM_CANPASTE, 0, 0));
	}

	/*
	��⵱ǰ�ı����Ƿ��ܽ��г�������
	���ܽ��в���������true
	*/
	bool CanRedo()
	{
		return (!SendMessage(CurrentHwnd, EM_CANREDO, 0, 0));
	}

	/*
	��ȡ�ı�����ѡȡ�ı��ķ�Χ
	lpMin: ѡȡ�Ŀ�ͷ
	lpMax: ѡȡ��ĩβ
	��ѡȡ��ͷΪ0��ѡȡĩβΪ-1����˵��ȫѡ
	*/
	void GetSelRange(int* lpMin, int* lpMax)
	{
		CHARRANGE cr;
		SendMessage(CurrentHwnd, EM_EXGETSEL, 0, (LPARAM)&cr);
		*lpMin = cr.cpMin;
		*lpMax = cr.cpMax;
	}

	/*
	�����ı�����ѡȡ�ı��ķ�Χ
	Min: ѡȡ�Ŀ�ͷ
	Max: ѡȡ��ĩβ
	��ѡȡ��ͷΪ0��ѡȡĩβΪ-1����˵��ȫѡ
	����ʵ��ѡȡ���ַ���
	*/
	int SetSelRange(int Min, int Max)
	{
		CHARRANGE cr;
		cr.cpMin = Min;
		cr.cpMax = Max;
		return (int)SendMessage(CurrentHwnd, EM_EXSETSEL, 0, (LPARAM)&cr);
	}

	/*
	�����ı�������������ַ���
	iCount: ����ı�������Ϊ0����ʹ��ϵͳĬ��
	*/
	void SetLimitText(int iCount)
	{
		SendMessage(CurrentHwnd, EM_EXLIMITTEXT, 0, (LPARAM)iCount);
	}

	/*
	���ı����в����ı�
	Find: ��Ҫ���ҵ��ı�
	MatchCase: �Ƿ���Ҫ��Сдƥ��
	WholeWord: �Ƿ���Ҫȫ��ƥ��
	Begin: ��ʼ�����ı���λ��
	End: ���������ı���λ��
	���ز��ҵ����ı���ʼ���֡���û���ҵ����򷵻�-1
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

	/* ��ȡ��ǰѡ����ַ��ĸ�ʽ */
	CHARFORMAT GetCharFormat()
	{
		CHARFORMAT cf = { 0 };
		cf.cbSize = sizeof(cf);
		cf.dwMask = CFM_ALL;
		SendMessage(CurrentHwnd, EM_GETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
		return cf;
	}

	/* ���õ�ǰѡ����ַ��ĸ�ʽ */
	bool SetCharFormat(CHARFORMAT cf)
	{
		return (bool)SendMessage(CurrentHwnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
	}

	/*
	�����ı����ı�
	NewText: �µ��ı�
	*/
	void SetText(char* NewText)
	{
		Text = NewText;
		SetWindowText(CurrentHwnd, NewText);
	}

	/*
	�����������ı����ı�
	NewText: ��Ч�����ֱ��ʽ
	*/
	void SetText(int NewText)
	{
		char Buffer[255];
		itoa(NewText, Buffer, 10);
		SetWindowText(CurrentHwnd, Buffer);
	}

	/* ��ȡ�ı� */
	char* GetText()
	{
		int Length = GetWindowTextLength(CurrentHwnd) + 1;
		char* tmp = new char[Length];
		GetWindowText(CurrentHwnd, tmp, Length);
		return tmp;
	}

	/* ���Ʋ��� */
	void Copy()
	{
		SendMessage(CurrentHwnd, WM_COPY, 0, 0);
	}

	/* ���в��� */
	void Cut()
	{
		SendMessage(CurrentHwnd, WM_CUT, 0, 0);
	}

	/* �������� */
	void Undo()
	{
		SendMessage(CurrentHwnd, EM_UNDO, 0, 0);
	}

	/* �ظ����� */
	void Redo()
	{
		SendMessage(CurrentHwnd, EM_REDO, 0, 0);
	}

	/* ���ѡ����ı� */
	void Clear()
	{
		SendMessage(CurrentHwnd, WM_CLEAR, 0, 0);
	}

	/* ճ������ */
	void Paste()
	{
		SendMessage(CurrentHwnd, WM_PASTE, 0, 0);
	}
};

//============================================================================
class MyTimePicker : public MyControls
{
public:
	bool		LongDateFormat;					//����ʱ���ʽ
	bool		RightAlign;						//���ұߵ�������
	bool		CheckBoxes;						//��ѡ����ʽ
	bool		TimeFormat;						//ʱ��ѡ����
	bool		UpDownButton;					//ʹ�õ��ڰ�ť

	//��������ʱ��ѡ�����ؼ�
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;					//��ʼ���ؼ���ʽ

		//����ؼ�����ʽ
		OrCalc(LongDateFormat, &lStyle, DTS_LONGDATEFORMAT);	//����ʱ���ʽ
		OrCalc(RightAlign, &lStyle, DTS_RIGHTALIGN);			//���ұߵ�������
		OrCalc(CheckBoxes, &lStyle, DTS_SHOWNONE);				//��ѡ����ʽ
		OrCalc(TimeFormat, &lStyle, DTS_TIMEFORMAT);			//ʱ��ѡ����
		OrCalc(UpDownButton, &lStyle, DTS_UPDOWN);				//ʹ�õ��ڰ�ť

		//�����ؼ�
		CurrentHwnd = CreateWindowEx(0, "MyTimePicker", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	/*
	��ȡ�ɹ�ѡ�������ʱ�䷶Χ
	lpMin: ������ʼʱ���SYSTEMTIME�ṹ��ָ��
	lpMax: ���ս���ʱ���SYSTEMTIME�ṹ��ָ��
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
	���ÿɹ�ѡ������ʱ�䷶Χ
	Begin: ��ʼʱ���SYSTEMTIME�ṹ��ָ�룬��ΪNULL��������ʼʱ��
	End: ����ʱ���SYSTEMTIME�ṹ��ָ�룬��ΪNULL�����ý���ʱ��
	��ִ�гɹ�������true
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
	��ȡ��ǰ�ؼ���ѡ���ʱ��
	Time: ��������ѡ���ʱ���SYSTEMTIME�ṹ��ָ��
	��ִ�гɹ�������true
	*/
	bool GetTime(SYSTEMTIME* Time)
	{
		if (Time)			//if (Time != NULL)
			return !SendMessage(CurrentHwnd, DTM_GETSYSTEMTIME, 0, (LPARAM)Time);		//return rtn == GDT_VALID;
		else
			return false;
	}

	/*
	���õ�ǰ�ؼ���ѡ���ʱ��
	Time: ��������ѡ���ʱ���SYSTEMTIME�ṹ��ָ�롣��ΪNULL����ѡ��ʱ��
	��ִ�гɹ�������true
	*/
	bool SetTime(SYSTEMTIME* Time)
	{
		if (&Time)
			return SendMessage(CurrentHwnd, DTM_SETSYSTEMTIME, GDT_VALID, (LPARAM)Time);
		else
			return SendMessage(CurrentHwnd, DTM_SETSYSTEMTIME, GDT_NONE, (LPARAM)Time);
	}

	/*
	��������ʱ���ʽ
	FormatString: ��ʽ�ַ�������ΪNULL����ʹ��ϵͳĬ�ϸ�ʽ
	��ִ�гɹ�������true
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
	bool		MultiSelect;					//����ѡȡ
	int			MultiSelectLimit;				//����ѡȡ����
	bool		WeekNumbers;					//��ʾ�ڼ���
	bool		NoTodayCircle;					//��Ȧѡ����
	bool		NoToday;						//����ʾ����
	bool		BlackBorder;					//��ɫ�߿�
	bool		ClientEdgeBorder;				//����߿�

	//���������ؼ�
	HWND Create()
	{
		LONG lStyle = WS_VISIBLE | WS_CHILD;					//��ʼ���ؼ���ʽ

		//����ؼ�����ʽ
		OrCalc(MultiSelect, &lStyle, MCS_MULTISELECT);
		OrCalc(WeekNumbers, &lStyle, MCS_WEEKNUMBERS);
		OrCalc(NoTodayCircle, &lStyle, MCS_NOTODAYCIRCLE);
		OrCalc(NoToday, &lStyle, MCS_NOTODAY);
		OrCalc(BlackBorder, &lStyle, WS_BORDER);

		//�����ؼ�
		CurrentHwnd = CreateWindowEx(ClientEdgeBorder ? WS_EX_CLIENTEDGE : 0, "MyMonthCalendar", "", lStyle,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		//���ÿؼ��Ŀ���ѡ����������
		SetMaxSelCount(MultiSelectLimit);

		return CurrentHwnd;
	}

	/* ��ȡ��ǰ����ѡ������� */
	SYSTEMTIME GetCurrentSel()
	{
		SYSTEMTIME st = { 0 };
		SendMessage(CurrentHwnd, MCM_GETCURSEL, 0, (LPARAM)&st);
		return st;
	}

	/*
	���õ�ǰѡ�������
	��ִ�гɹ�������true
	*/
	bool SetCurrentSel(SYSTEMTIME* SelDate)
	{
		if (SelDate)
			return (bool)(SendMessage(CurrentHwnd, MCM_SETCURSEL, 0, (LPARAM)SelDate));
		else
			return false;
	}

	/*
	��ȡһ���еĵ�һ��
	����ֵ�ǵ�ǰ������һ�ܵĵ�һ�죬0��������һ��1�������ڶ����Դ�����
	*/
	int GetFirstDayOfWeek()
	{
		return (int)LOWORD((DWORD)SendMessage(CurrentHwnd, MCM_GETFIRSTDAYOFWEEK, 0, 0));
	}

	/*
	����һ���еĵ�һ��
	NewDay: һ���еĵ�һ�졣0��������һ��1�������ڶ����Դ�����
	*/
	void SetFirstDatOfWeek(int NewDay)
	{
		SendMessage(CurrentHwnd, MCM_SETFIRSTDAYOFWEEK, 0, NewDay);
	}

	/* ��ȡ�����ؼ�����ѡ�����ڵ������Ŀ */
	int GetMaxSelCount()
	{
		return (int)SendMessage(CurrentHwnd, MCM_GETMAXSELCOUNT, 0, 0);
	}

	/*
	���������ؼ�����ѡ�����ڵ������Ŀ
	��ִ�гɹ�������true
	*/
	bool SetMaxSelCount(int MaxCount)
	{
		return (bool)SendMessage(CurrentHwnd, MCM_SETMAXSELCOUNT, (WPARAM)MaxCount, 0);
	}

	/*
	��ȡѡ��ķ�Χ
	stBegin: ѡ��Ŀ�ʼ������
	stEnd: ѡ��Ľ������¼�
	��ִ�гɹ�������true
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
	����ѡ��ķ�Χ
	stBegin: ѡ��Ŀ�ʼ������
	stEnd: ѡ��Ľ������¼�
	��ִ�гɹ�������true
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
	��ȡ����ѡ�����ڵķ�Χ
	stBegin: ѡ��Ŀ�ʼ������
	stEnd: ѡ��Ľ������¼�
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
	��������ѡ�����ڵķ�Χ
	stBegin: ѡ��Ŀ�ʼ������
	stEnd: ѡ��Ľ������¼�
	��ִ�гɹ�������true
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

	/* ��ȡ��������� */
	SYSTEMTIME GetToday()
	{
		SYSTEMTIME st = { 0 };
		SendMessage(CurrentHwnd, MCM_GETTODAY, 0, (LPARAM)&st);
		return st;
	}

	/* ���ý�������� */
	void SetToday(SYSTEMTIME NewDate)
	{
		SendMessage(CurrentHwnd, MCM_SETTODAY, 0, (LPARAM)&NewDate);
	}
};

//============================================================================
class MyIpAddress : public MyControls
{
public:
	//����IP��ַ�ؼ�
	HWND Create()
	{
		//�����ؼ�
		CurrentHwnd = CreateWindowEx(0, "SysIPAddress32", "", WS_VISIBLE | WS_CHILD,
			Left, Top, Width, Height, GetCurrentHwnd(), (HMENU)hMenu, GetCurrentHinstance(), 0);

		//���ÿؼ��Ŀ��Ӻͼ���״̬
		SetVisible(Visible);
		SetEnabled(Enabled);

		return CurrentHwnd;
	}

	//���IP��ַ
	void Clear()
	{
		SendMessage(CurrentHwnd, IPM_CLEARADDRESS, 0, 0);
	}

	//��ȡ��ǰ��IP��ַ
	char* GetIpAddress()
	{
		char* tmp = new char[20];
		GetWindowText(CurrentHwnd, tmp, 20);
		return tmp;
	}

	/*
	�жϵ�ǰ�ؼ��Ƿ�Ϊ��
	����ǰ�ؼ�Ϊ�գ��򷵻�true
	*/
	bool IsBlank()
	{
		return SendMessage(CurrentHwnd, IPM_ISBLANK, 0, 0);
	}

	/*
	���õ�ǰ��IP��ַ
	Field0 - 3: �ֱ��ӦIP��ַ���ĸ�����
	*/
	void SetIpAddress(BYTE Field0, BYTE Field1, BYTE Field2, BYTE Field3)
	{
		SendMessage(CurrentHwnd, IPM_SETADDRESS, 0, MAKEIPADDRESS(Field0, Field1, Field2, Field3));
	}

	/*
	�ÿؼ���ȡ����
	Field: ��Ҫ��ȡ�����λ����š�����0λ����
	*/
	void SetFocus(int Field)
	{
		SendMessage(CurrentHwnd, IPM_SETFOCUS, (WPARAM)Field, 0);
	}

	/*
	����IP��ַ��Ӧλ�õ�����
	Field: ��Ҫ�������Ƶ�λ����š�����0λ����
	Min: ��Ҫ���õ�����ֵ
	Max: ��Ҫ���õ�����ֵ
	��ִ�гɹ�������true
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
	char*		ClassName;						//����
	char*		Caption;						//�������
	COLORREF	BackColor;						//���屳����ɫ
	DWORD		Style;							//��ʽ
	DWORD		ExStyle;						//��չ��ʽ
	HWND		CurrentHwnd;					//��ǰ���
	bool		Visible;						//����
	bool		Enabled;						//����
	LONG		Left;							//ˮƽλ��
	LONG		Top;							//��ֱλ��
	LONG		Width;							//������
	LONG		Height;							//����߶�
	RECT		WindowPos;						//��ǰ����λ��
	HDC			hDC;							//�豸�����ľ��
	int			CurrentX;						//��ǰ�����ı������X������
	int			CurrentY;						//��ǰ�����ı������Y������
	HINSTANCE   hInstance;						//��ǰ����ʵ�����
	//-----------------------------------------------------------------------
	//�������еĿؼ���Ҫ�����ﶨ��
��AllControlsHere��
	//-----------------------------------------------------------------------
	/* ������ */
	void InitClass()
	{
��WindowInitCodeHere��
		/* ����RichEdit��̬�� */
		LoadLibrary("RichEd20.dll");
		/* ����ͨ�ÿؼ���̬�� */
		LoadLibrary("comctl32.dll");
		/*=======================================*/
		/* �����пؼ����г��໯ */
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
	������ǰ�Ĵ��塣������������ڵ��øù��̡�
	mHinstance: �����������ʵ�������hInstance��
	*/
	HWND Create(HINSTANCE mHinstance)
	{
		InitClass();
		
		WNDCLASS MyClass;																	//������
		MSG Msg;																			//������Ϣ

		MyClass.cbClsExtra = 0;
		MyClass.cbWndExtra = 0;
		MyClass.hbrBackground = (HBRUSH)CreateSolidBrush(BackColor);						//������ɫ
		MyClass.hCursor = LoadCursor(0, IDC_ARROW);											//���
		MyClass.hIcon = LoadIcon(0, IDI_APPLICATION);										//ͼ��
		MyClass.hInstance = mHinstance;														//ʵ�����
		MyClass.lpfnWndProc = WndProc;														//��Ϣ����
		MyClass.lpszClassName = TEXT(ClassName);											//����
		MyClass.lpszMenuName = NULL;
		MyClass.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;

		RegisterClass(&MyClass);															//ע����

		CurrentHwnd = CreateWindowEx(ExStyle, TEXT(ClassName), TEXT(Caption), Style,		//��������
			Left, Top, Width, Height,
			0, NULL, 0, 0);

		tme.cbSize = sizeof(TRACKMOUSEEVENT);												//�������׷��
		tme.dwFlags = TME_LEAVE;
		tme.hwndTrack = CurrentHwnd;
		TrackMouseEvent(&tme);																//�������׷��
		UpdateWindow(CurrentHwnd);															//ˢ�´���

		hDC = GetDC(CurrentHwnd);															//��ȡ����������ľ����hDC��
		hInstance = mHinstance;																//��¼��ǰ��ʵ�����

		if (CurrentHwnd == 0)																//�������ľ��Ϊ0˵������ʧ��
		{
			UnregisterClass(ClassName, mHinstance);
			return 0;
		}

		CreateAllControls();																//�������еĿؼ�
		Form_Load();																		//������������¼�
		
		while (GetMessage(&Msg, CurrentHwnd, 0, 0) > 0)										//��������Ϣѭ��
		{
			//ѭ���������е���Ϣ��ֱ������ر�
			TranslateMessage(&Msg);
			DispatchMessage(&Msg);
		}

		UnregisterClass(ClassName, mHinstance);									//����رպ�ж����

		return CurrentHwnd;																	//��������
	}

	/* ����ĸ��ַ��� */

	//�����رմ���
	void Close()
	{
		SendMessage(CurrentHwnd, WM_CLOSE, 0, 0);		//����WM_CLOSE��Ϣ�Թرմ���
	}

	//ǿ�йرմ���
	void Destroy()
	{
		DestroyWindow(CurrentHwnd);						//ǿ�дݻٴ���				
		PostQuitMessage(0);								//ֱ���˳��߳�
	}

	/*
	���ô������
	NewCaption: �µı����ַ���
	*/
	void SetCaption(char* NewCaption)
	{
		SetWindowText(CurrentHwnd, NewCaption);			//���Ĵ������
	}

	/*
	����ֵ�����ô�����⡣
	NewCaption: ��Ч����ֵ���ʽ
	*/
	void SetCaption(int NewCaption)						
	{
		char tmp[255];
		itoa(NewCaption, tmp, 10);
		SetCaption(tmp);
	}

	/*
	���ñ�����ɫ��
	NewColor: �µ���ɫ
	*/
	void SetBackColor(COLORREF NewColor)
	{
		HBRUSH ColorBrush = CreateSolidBrush(NewColor);							//����ָ����ɫ��ˢ��
		SetClassLongPtr(CurrentHwnd, GCLP_HBRBACKGROUND, (LONG)ColorBrush);		//������ı�����ɫ
		SendMessage(CurrentHwnd, WM_ERASEBKGND, (WPARAM)hDC, 0);				//����WM_ERASEBKGND��ˢ�´���
	}

	/*
	���ô����Ƿ���ӡ�
	IsVisible: �����Ƿ����
	*/
	void SetVisible(bool IsVisible)
	{
		ShowWindow(CurrentHwnd, IsVisible);
		Visible = IsVisible;
	}

	/*
	���ô����Ƿ���á�
	IsEnabled: �Ƿ����ô���
	*/
	void SetEnabled(bool IsEnabled)
	{
		EnableWindow(CurrentHwnd, IsEnabled);
		Enabled = IsEnabled;
	}

	/*
	���ô����������ʽ��
	������
	FontSize: �ı���С
	FontBoldWidth: �ı�����̶�
	FontItalic: �Ƿ�Ϊб��
	FontUnderline: �Ƿ����»���
	FontStrikeOut: �Ƿ���ɾ����
	FontName: ��������
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
	���ô������ʽ��
	StyleAdd: ��ӵ���ʽ
	StyleRemove: ȥ������ʽ
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
	���ô����ı���ɫ��
	NewColor: �µ���ɫ
	*/
	void SetForeColor(COLORREF NewColor)
	{
		SetTextColor(hDC, NewColor);
	}

	/*
	���ô���������ı��Ƿ�͸������������Print()��
	bTransparent: �Ƿ�͸��
	*/
	void SetFontTransparent(bool bTransparent)
	{
		SetBkMode(hDC, bTransparent ? OPAQUE : TRANSPARENT);
	}

	/*
	���Ĵ���λ�á�
	������
	NewLeft: �µ�X����
	NewTop: �µ�Y����
	NewWidth: �µĿ��
	NewHeight: �µĸ߶�
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
	��ָ��λ������ı���
	������
	TextPrint: ��Ҫ������ı�
	X: �����X��λ��
	Y: �����Y��λ��
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
	�Դ���Ĭ��λ������ı���
	TextPrint: ��Ҫ������ı�
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

	//���ش���
	void Hide()
	{
		ShowWindow(CurrentHwnd, SW_HIDE);
	}
	
	//��ʾ����
	void Show()
	{
		ShowWindow(CurrentHwnd, SW_SHOW);
	}

	/*
	���ô�����ǰ��
	bTopMost: �Ƿ���ǰ��
	*/
	void SetTopMost(bool bTopMost)
	{
		SetWindowPos(CurrentHwnd, bTopMost ? HWND_TOPMOST : HWND_NOTOPMOST, 
			0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
	}

	/*
	���ô���͸��
	Degree: ����͸���ĳ̶ȣ���ΧΪ0��255��255��ʾ��͸����0��ʾ��ȫ͸����
	*/
	void SetTransparent(unsigned char Degree)
	{
		SetWindowLong(CurrentHwnd, GWL_EXSTYLE, 
			GetWindowLong(CurrentHwnd, GWL_EXSTYLE) | WS_EX_LAYERED);
		SetLayeredWindowAttributes(CurrentHwnd, 0, Degree, LWA_ALPHA);
	}

} MainWindow;

/* ��ȡ��ǰ�����λ�úʹ�С */
void GetCurrentRect()
{
	GetWindowRect(MainWindow.CurrentHwnd, &MainWindow.WindowPos);
	MainWindow.Left = MainWindow.WindowPos.left;
	MainWindow.Top = MainWindow.WindowPos.top;
	MainWindow.Width = MainWindow.WindowPos.right - MainWindow.WindowPos.left;
	MainWindow.Height = MainWindow.WindowPos.bottom - MainWindow.WindowPos.top;
}

/* ��ȡ��ǰ����ľ�� */
HWND GetCurrentHwnd()
{
	return MainWindow.CurrentHwnd;
}

/* ��ȡ��ǰ�����ʵ����� */
HINSTANCE GetCurrentHinstance()
{
	return MainWindow.hInstance;
}


/* Ϊָ������Ĵ������ӻ���ȥ��ָ������ʽ
 * TargetHwnd��Ŀ�괰��ľ��
 * StyleAdd����Ҫ���ӵ���ʽ
 * StyleRemove����Ҫȥ������ʽ
 */
void SetWindowLongEx(HWND TargetHwnd, DWORD StyleAdd = 0, DWORD StyleRemove = 0)
{
	SetWindowLong(TargetHwnd, GWL_STYLE, GetWindowLong(TargetHwnd, GWL_STYLE) | StyleAdd & (~StyleRemove));
}

/*
 * ����ָ���Ĳ����ͱ��ʽ�ж��Ƿ�ִ������
 * bExpression: ָ���Ĳ����ͱ��ʽ
 * StyleLong: ��Ҫ������ֵ��Ŀ��
 * StyleAdd: �������ͱ��ʽΪ���ʱ�򽫸���ʽ��ӵ�Ŀ����ֵ��
*/
void OrCalc(bool bExpression, long* StyleLong, long StyleAdd)
{
	if (bExpression)
	{
		*StyleLong = *StyleLong | StyleAdd;
	}
}

/*
 * �ϵ����й���
 * CodeLn: �ϵ��Ӧ�Ĵ�������
 */
void Breakpoint(long CodeLn)
{
	SendMessage(DEBUGGER_HWND, MY_DEBUGGER_BREAKPOINT, (WPARAM)CodeLn, 0);		//��������������Ϣ���������жϵ�����
	SuspendProcess();															//�Լ������Լ�
}

/*
 * ������̹���
 * ProcessID: ����ID
 */
void SuspendProcess()
{
	NtSuspendProcess pfnNtSuspendProcess = (NtSuspendProcess)GetProcAddress(
		GetModuleHandle("ntdll"), "NtSuspendProcess");							//��ȡAPI�ĺ�����ַ
	pfnNtSuspendProcess(GetCurrentProcess());									//�������
}

/*
 * ���Ӷϵ����й���
 * WatchIndex�����ӵ����
 * VarAddr�������ĵ�ַ
 * DataType��������������
 */
void WatchBreakpoint(int WatchIndex, void* VarAddr, SIZE_T nSize)
{
	SendMessage(DEBUGGER_HWND, MY_DEBUGGER_MEMDATA,								//��������������Ϣ����������Ҫ���¼�����Ϣ
		MAKEWPARAM(WatchIndex, nSize), (LPARAM)VarAddr);
}

/*
* ����ע�������
* ClassName: ��Ҫ���±�ע���������
* NewClassName: ��ע���������
* lpfnPrevWndProc: �ɵ�WndProc��ַ
* lpfnWndProc: �µ�WndProc��ַ
*/
void ReregisterClass(LPCSTR ClassName, LPCSTR NewClassName,
	WNDPROC *lpfnPrevWndProc, WNDPROC lpfnWndProc)
{
	WNDCLASSEX ctlClass;											//�ؼ���
	ZeroMemory(&ctlClass, sizeof(WNDCLASSEX));						//��ʼ���ؼ������

	GetClassInfoEx(GetCurrentHinstance(), ClassName, &ctlClass);	//��ȡ����Ϣ
	*lpfnPrevWndProc = ctlClass.lpfnWndProc;						//��¼ԭWndProc��ַ
	ctlClass.lpfnWndProc = lpfnWndProc;								//�滻��WndProc��ַ
	ctlClass.lpszClassName = NewClassName;							//��������
	ctlClass.cbSize = sizeof(WNDCLASSEX);							//����cbSize
	RegisterClassEx(&ctlClass);										//����ע����
}