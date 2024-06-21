
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
#include <xlsxio_write.h>
#include "libxl.h"
using namespace libxl;

// CAboutDlg dialog used for App About

std::vector<TaskExcel> g_Tasks;
int g_nTotalThread = 0;
BOOL g_bStop = FALSE;
BOOL g_bStarted = FALSE;

#define ENABLE_WINDOW(id, x) do { \
	GetDlgItem(id)->EnableWindow(b); \
} while(0)

void CString2Str(CString source, char* target) {
	for (int i = 0; i < source.GetLength(); i++) {
		target[i] = source.GetAt(i);
	}
}

void CString2Wstr(CString source, TCHAR* target) {
	for (int i = 0; i < source.GetLength(); i++) {
		target[i] = source.GetAt(i);
	}
}

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




CScrapperDlg::CScrapperDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_SCRAPPER_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CScrapperDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_EXCEL_FILES, m_ListCtrl);
	DDX_Control(pDX, IDC_PROGRESS2, m_Percent);
	DDX_Check(pDX, IDC_CHECK_PROXY, m_UseProxySetting);
	DDX_Control(pDX, IDC_EDIT_PROXY, m_editProxy);

	DDX_Text(pDX, IDC_EDIT3, m_strTLDs);
	
	
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
	ON_BN_CLICKED(IDCANCEL, &CScrapperDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_CHECK_PROXY, &CScrapperDlg::OnBnClickedCheckProxy)
	ON_EN_CHANGE(IDC_EDIT_PROXY, &CScrapperDlg::OnEnChangeEditProxy)
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
	
	int result = _setmaxstdio(8192);

	m_Percent.ShowWindow(SW_HIDE);

	UpdateData();
	m_UseProxySetting = FALSE;
	m_editProxy.EnableWindow(FALSE);

	FILE* fp;
	fopen_s(&fp, "proxy.config", "rb");
	if (fp) {
		char szBuf[0x100] = { 0 };
		fread(szBuf, 1, 0x100 - 4, fp);
		fclose(fp);

		TCHAR szCP[0x100] = { 0 };
		for (int i = 0; i < strlen(szBuf); i++) {
			szCP[i] = szBuf[i];
		}

		m_editProxy.SetWindowTextW(szCP);
	}
	LoadTLD();
	UpdateData(FALSE);
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

void execute_curl_command(const char* url) {
	
}

DWORD WINAPI CScrapperDlg::ThreadMonitor(LPVOID lpParam) {
	CScrapperDlg* pThis = (CScrapperDlg*)lpParam;
	while (1) {
		if (g_nTotalThread <= 0) {
			break;
		}
	}

	// save
	if (!g_bStop) {
		for (int i = 0; i < g_Tasks.size(); i++) {
			Book* book = xlCreateXMLBook();
			book->setKey(_T("Halil Kural"), _T("windows-2723210a07c4e90162b26966a8jcdboe"));
			if (book) {
				TCHAR szXlsx[MAX_PATH] = { 0 };
				CString2Wstr(g_Tasks[i].file, szXlsx);
				if (book->load(szXlsx)) {
					Sheet* sheet = book->getSheet(0);
					if (sheet) {
						for (int j = 0; j < g_Tasks[i].saves.size(); j++) {
							TCHAR mail[0x100] = { 0 };
							for (int ii = 0; ii < strlen(g_Tasks[i].saves[j].mail); ii++) {
								mail[ii] = g_Tasks[i].saves[j].mail[ii];
							}
							sheet->writeStr(g_Tasks[i].saves[j].row, g_Tasks[i].col - 1, mail);
						}
					}
					book->save(szXlsx);
				}
				book->release();
			}
		}
	}
	pThis->EnableAllButtons(TRUE);
	pThis->Terminated();
	g_bStarted = FALSE;
	return (DWORD)0;
}

int is_valid_email_char(char c) {
	return isalnum(c) || c == '.' || c == '_' || c == '%' || c == '+' || c == '-' || c == '@';
}

int is_valid_email(const char* email) {
	const char* at = strchr(email, '@');
	if (!at) return 0;

	const char* dot = strrchr(email, '.');
	if (!dot || dot < at) return 0;

	if (at == email || *(at + 1) == '\0') return 0;

	if (*(dot + 1) == '\0') return 0;

	return 1;
}


void search_email_addresses(const char* str) {
	const char* p = str;
	while (*p) {
		while (*p && !isalnum(*p)) p++;

		const char* start = p;
		while (*p && is_valid_email_char(*p)) p++;

		if (start != p) {
			char email[256] = { 0 };
			strncpy_s(email, start, p - start);
			email[p - start] = '\0';

			if (is_valid_email(email)) {
				printf("Found email address: %s\n", email);
			}
		}

		if (*p) p++;
	}
}

struct ThreadParam {
	CScrapperDlg* pThis;
	int index;
};

DWORD WINAPI CScrapperDlg::ThreadScrapping(LPVOID lpParam) {
	g_nTotalThread++;
	ThreadParam* param = (ThreadParam *)lpParam;
	int i = param->index;
	CScrapperDlg* pThis = param->pThis;
	while (1) {
		if (g_bStop) {
			break;
		}
		EnterCriticalSection(&g_Tasks[i].mutex);
		if (g_Tasks[i].pos >= g_Tasks[i].items.size()) {
			LeaveCriticalSection(&g_Tasks[i].mutex);
			break;
		}
		TaskItem item = g_Tasks[i].items[g_Tasks[i].pos];
		g_Tasks[i].pos++;
		LeaveCriticalSection(&g_Tasks[i].mutex);

		char command[256];
		snprintf(command, sizeof(command), "curl -s \"%s%s\"", strstr(item.url, "http") ? "" : "https://", item.url);

		SECURITY_ATTRIBUTES sa;
		sa.nLength = sizeof(SECURITY_ATTRIBUTES);
		sa.lpSecurityDescriptor = NULL;
		sa.bInheritHandle = TRUE;

		HANDLE hRead, hWrite;
		if (!CreatePipe(&hRead, &hWrite, &sa, 0)) {
			goto _CONTINUE;
		}

		if (!SetHandleInformation(hRead, HANDLE_FLAG_INHERIT, 0)) {
			goto _CONTINUE;
		}

		STARTUPINFOA si;
		ZeroMemory(&si, sizeof(si));
		si.cb = sizeof(si);
		si.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
		si.wShowWindow = SW_HIDE;
		si.hStdOutput = hWrite;
		si.hStdError = hWrite;

		PROCESS_INFORMATION pi;
		ZeroMemory(&pi, sizeof(pi));

		if (!CreateProcessA(NULL, command, NULL, NULL, TRUE, 0, NULL, NULL, &si, &pi)) {
			goto _CONTINUE;
		}
		CloseHandle(hWrite);
		
		char buf[0x400] = { 0 };
		DWORD bytesRead;
		while (ReadFile(hRead, buf, sizeof(buf) - 1, &bytesRead, NULL) && bytesRead > 0) {
			char* s = buf;
			char* p = strstr(s, "@");
			while (p) {
				char* next = strstr(p + 1, "@");
				char* ss;
				for (ss = p; ss >= s; ss--) {
					if (!is_valid_email_char(*ss)) {
						break;
					}
				}
				s = ss + 1;
				for (ss = s; ss < (buf + bytesRead); ss++) {
					if (!is_valid_email_char(*ss)) {
						*ss = '\0';
						break;
					}
				}
				if (is_valid_email(s)) {
					printf("%s", s);

					TaskSave t;
					t.row = item.row;
					memset(t.mail, 0, sizeof(t.mail));
					strcpy_s(t.mail, s);

					EnterCriticalSection(&g_Tasks[i].mutex);
					g_Tasks[i].saves.push_back(t);
					LeaveCriticalSection(&g_Tasks[i].mutex);
				}
				p = next;
			}
			memset(buf, 0, sizeof(buf));
		}
	_CONTINUE:
		CloseHandle(hRead);
		CloseHandle(pi.hProcess);
		CloseHandle(pi.hThread);
		pThis->UpdatePercent();
	}
	g_nTotalThread--;
	free(param);
	return (DWORD)0;
}

void CScrapperDlg::OnBnClickedOk() {
	g_nTotalThread = 0;
	g_bStop = FALSE;
	g_bStarted = FALSE;

	for (int i = 0; i < g_Tasks.size(); i++) {
		g_Tasks[i].pos = 0;
		g_Tasks[i].saves.clear();
	}

	if (g_Tasks.size() == 0) {
		AfxMessageBox(_T("No Task"));
		return;
	}
	m_Percent.ShowWindow(SW_SHOW);
	UpdatePercent();
	EnableAllButtons(FALSE);

	for (int i = 0; i < g_Tasks.size(); i++) {
		for (int j = 0; j < g_Tasks[i].thread; j++) {
			ThreadParam *param = (ThreadParam *)malloc(sizeof(ThreadParam));
			param->pThis = this;
			param->index = i;
			CreateThread(
				NULL,
				0,
				ThreadScrapping,
				(LPVOID)param,
				0,
				NULL);
		}
	}
	Sleep(500);

	CreateThread(
		NULL,
		0,
		ThreadMonitor,
		this,
		0,
		NULL);

	GetDlgItem(IDCANCEL)->EnableWindow(TRUE);
	GetDlgItem(IDCANCEL)->SetWindowTextW(_T("Stop"));
	g_bStarted = TRUE;
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

void CScrapperDlg::SetThreadColumn(int index, int nThread, CString column) {
	EnableAllButtons(FALSE);
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
		EnableAllButtons(TRUE);
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
	EnableAllButtons(TRUE);
}

void CScrapperDlg::EnableAllButtons(bool b) {
	ENABLE_WINDOW(IDOK, b);
	ENABLE_WINDOW(IDC_BUTTON_ADD, b);
	ENABLE_WINDOW(IDC_BUTTON_EDIT, b);
	ENABLE_WINDOW(IDC_BUTTON_REMOVE, b);
	ENABLE_WINDOW(IDC_BUTTON_CLEAR, b);
	ENABLE_WINDOW(IDCANCEL, b);
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


void CScrapperDlg::OnBnClickedCancel() {
	if (g_bStarted) {
		GetDlgItem(IDCANCEL)->EnableWindow(FALSE);
		
		g_bStop = TRUE;
		return;
	}
	CDialogEx::OnCancel();
}

void CScrapperDlg::UpdatePercent() {
	int total = 0;
	int complete = 0;
	for (int i = 0; i < g_Tasks.size(); i++) {
		total += g_Tasks[i].items.size();
		complete += g_Tasks[i].pos;
	}
	m_Percent.SetRange(0, total);
	m_Percent.SetPos(complete);
}

void CScrapperDlg::Terminated() {
	GetDlgItem(IDCANCEL)->SetWindowTextW(_T("Close"));
	m_Percent.ShowWindow(SW_HIDE);
	if (g_bStop) {
		AfxMessageBox(_T("Stopped"));
	}
	else {
		AfxMessageBox(_T("Finished"));
	}
}

void CScrapperDlg::OnBnClickedCheckProxy() {
	UpdateData();
	m_editProxy.EnableWindow(m_UseProxySetting);
	UpdateData(FALSE);
}

void CScrapperDlg::OnEnChangeEditProxy() {
	UpdateData();
	TCHAR szProxy[0x100] = { 0 };
	m_editProxy.GetWindowText(szProxy, 0x100 - 4);
	char szBuf[0x100] = { 0 };
	for (int i = 0; i < wcslen(szProxy); i++) {
		szBuf[i] = szProxy[i];
	}
	FILE* fp;
	fopen_s(&fp, "proxy.config", "wb");
	if (fp) {
		fwrite(szBuf, 1, strlen(szBuf), fp);
		fclose(fp);
	}
	UpdateData(FALSE);
}

void CScrapperDlg::LoadTLD() {
	FILE* fp;
	fopen_s(&fp, "TLD.txt", "rb");
	if (fp) {
		char buffer[0x100] = { 0 };
		while (fgets(buffer, sizeof(buffer), fp) != NULL) {
			char cp[0x100] = { 0 };
			int offset = 0;
			for (int i = 0; i < strlen(buffer); i++) {
				CString tmp;
				if (buffer[i] >= 'A' && buffer[i] <= 'Z') {
					cp[offset++] = buffer[i] - 'A' + 'a';
					tmp.Format(_T("%c"), buffer[i] - 'A' + 'a');
					m_strTLDs += tmp;
				}
				else if (buffer[i] >= 'a' && buffer[i] <= 'z') {
					cp[offset++] = buffer[i];
					tmp.Format(_T("%c"), buffer[i]);
					m_strTLDs += tmp;
				}
			}
			memset(buffer, 0, sizeof(buffer));

			if (strlen(cp)) {
				m_strTLDs += _T("\r\n");
			}
		}
		fclose(fp);
	}
	UpdateData(FALSE);
}