
// ScrapperDlg.cpp : implementation file
//

#include "pch.h"
#include "framework.h"
#include "Scrapper.h"
#include "ScrapperDlg.h"
#include "afxdialogex.h"
#include "CSettingDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#include <xlsxio_read.h>

// CAboutDlg dialog used for App About

std::vector<TaskExcel> g_Tasks;

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CScrapperDlg dialog



CScrapperDlg::CScrapperDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_SCRAPPER_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CScrapperDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_EXCEL_FILES, m_ListCtrl);
}

BEGIN_MESSAGE_MAP(CScrapperDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &CScrapperDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDC_BUTTON_ADD, &CScrapperDlg::OnBnClickedButtonAdd)
	ON_BN_CLICKED(IDC_BUTTON_EDIT, &CScrapperDlg::OnBnClickedButtonEdit)
	ON_BN_CLICKED(IDC_BUTTON_REMOVE, &CScrapperDlg::OnBnClickedButtonRemove)
	ON_BN_CLICKED(IDC_BUTTON_CLEAR, &CScrapperDlg::OnBnClickedButtonClear)
END_MESSAGE_MAP()


// CScrapperDlg message handlers

BOOL CScrapperDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	m_ListCtrl.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

	m_ListCtrl.InsertColumn(COL_File,		_T("File"),			LVCFMT_LEFT,	250);
	m_ListCtrl.InsertColumn(COL_Rows,		_T("Rows"),			LVCFMT_RIGHT,	80);
	m_ListCtrl.InsertColumn(COL_URL,		_T("URL column"),	LVCFMT_CENTER,	80);
	m_ListCtrl.InsertColumn(COL_Threads,	_T("Threads"),		LVCFMT_RIGHT,	80);
	m_ListCtrl.InsertColumn(COL_Status,		_T("Status"),		LVCFMT_CENTER,	80);

	AdjustListColumn(&m_ListCtrl);
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CScrapperDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CScrapperDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CScrapperDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CScrapperDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	CDialogEx::OnOK();
}

void CScrapperDlg::AdjustListColumn(CListCtrl *list) {
	if (list->GetHeaderCtrl()->GetItemCount() == 0) {
		return;
	}
	CRect rect;
	list->GetClientRect(&rect);

	int totalWidth = rect.Width();
	int columnCount = list->GetHeaderCtrl()->GetItemCount();

	int otherColumnsWidth = 0;
	for (int i = 0; i < columnCount - 1; i++) {
		otherColumnsWidth += list->GetColumnWidth(i);
	}

	int lastColumnWidth = totalWidth - otherColumnsWidth - 2;
	if (lastColumnWidth > 0) {
		list->SetColumnWidth(columnCount - 1, lastColumnWidth);
	}
}

void CScrapperDlg::OnBnClickedButtonAdd() {
	CString filter = _T("Excel Files (*.xlsx)|*.xlsx||");

	CFileDialog dlg(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, filter, this);

	if (dlg.DoModal() != IDOK) {
		return;
	}

	CString filePath = dlg.GetPathName();
	int n = m_ListCtrl.GetItemCount();
	m_ListCtrl.InsertItem(n, dlg.GetPathName());
	m_ListCtrl.SetItemText(n, COL_Rows, _T("--"));
	m_ListCtrl.SetItemText(n, COL_URL, _T("--"));
	m_ListCtrl.SetItemText(n, COL_Threads, _T("--"));
	m_ListCtrl.SetItemText(n, COL_Status, _T("Reading..."));

	TaskExcel task;
	task.file = dlg.GetPathName();
	task.pos = 0;
	InitializeCriticalSection(&task.mutex);

	g_Tasks.push_back(task);

	CSettingDlg settingDlg;
	settingDlg.OpenModal(this, TRUE, n, 1, _T("B"));
}

void CScrapperDlg::RemoveItem(int index) {
	m_ListCtrl.DeleteItem(index);
	g_Tasks.erase(g_Tasks.begin() + index);
}

void CString2Str(CString source, char* target) {
	for (int i = 0; i < source.GetLength(); i++) {
		target[i] = source.GetAt(i);
	}
}

void CScrapperDlg::SetThreadColumn(int index, int nThread, CString column) {
	CString tmp;
	tmp.Format(_T("%d"), nThread);
	m_ListCtrl.SetItemText(index, COL_Threads, tmp);
	m_ListCtrl.SetItemText(index, COL_URL, column);

	char szFile[MAX_PATH] = { 0 };
	CString file = m_ListCtrl.GetItemText(index, COL_File);
	CString2Str(file, szFile);
	
	xlsxioreader xlsxioread;
	if ((xlsxioread = xlsxioread_open(szFile)) == NULL) {
		AfxMessageBox(_T("Error opening .xlsx file"));
		RemoveItem(index);
		return;
	}

	int urlIdx = column.GetAt(0) - 'A';
	char* value;
	xlsxioreadersheet sheet;
	const char* sheetname = NULL;

	g_Tasks[index].col = urlIdx;
	g_Tasks[index].thread = nThread;
	g_Tasks[index].items.clear();

	int totalRow = 0;
	int ValidRow = 0;
	int InvalidRow = 0;
	int EmptyRow = 0;
	if ((sheet = xlsxioread_sheet_open(xlsxioread, sheetname, XLSXIOREAD_SKIP_EMPTY_ROWS)) != NULL) {
		while (xlsxioread_sheet_next_row(sheet)) {
			int tmp = 0;
			while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL) {
				if (tmp == urlIdx) {
					TCHAR* url = (TCHAR*)value;
					if (url[0]) {
						if (wcsstr(url, _T("."))) {
							ValidRow++;

							TaskItem item;
							item.row = totalRow;
							memset(item.url, 0, sizeof(item.url));
							for (int j = 0; j < wcslen(url); j++) {
								item.url[j] = url[j];
							}
							g_Tasks[index].items.push_back(item);

						}
						else {
							InvalidRow++;
						}
					}
					else {
						EmptyRow++;
					}
					xlsxioread_free(value);
					break;
				}
				tmp++;
				xlsxioread_free(value);
			}
			totalRow++;
		}
		xlsxioread_sheet_close(sheet);
	}
	xlsxioread_close(xlsxioread);

	CString status;
	status.Format(_T("Total: %d, Empty: %d, Valid: %d, Invalid: %d"), totalRow, EmptyRow, ValidRow, InvalidRow);
	m_ListCtrl.SetItemText(index, COL_Status, status);

	status.Format(_T("%d"), totalRow);
	m_ListCtrl.SetItemText(index, COL_Rows, status);
}

void CScrapperDlg::OnBnClickedButtonEdit() {
	POSITION pos = m_ListCtrl.GetFirstSelectedItemPosition();
	if (pos == NULL) {
		AfxMessageBox(_T("No item selected."));
		return;
	}
	int nItem = m_ListCtrl.GetNextSelectedItem(pos);

	int thread = _ttoi(m_ListCtrl.GetItemText(nItem, COL_Threads));
	CString col = m_ListCtrl.GetItemText(nItem, COL_URL);

	CSettingDlg settingDlg;
	settingDlg.OpenModal(this, FALSE, nItem, thread, col);
}


void CScrapperDlg::OnBnClickedButtonRemove() {
	POSITION pos = m_ListCtrl.GetFirstSelectedItemPosition();
	if (pos == NULL) {
		AfxMessageBox(_T("No item selected."));
		return;
	}
	int nItem = m_ListCtrl.GetNextSelectedItem(pos);
	m_ListCtrl.DeleteItem(nItem);
}


void CScrapperDlg::OnBnClickedButtonClear() {
	m_ListCtrl.DeleteAllItems();
}
