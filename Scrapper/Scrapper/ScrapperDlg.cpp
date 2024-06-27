
// ScrapperDlg.cpp : implementation file
//

#include "pch.h"
#include "framework.h"
#include "Scrapper.h"
#include "ScrapperDlg.h"
#include "afxdialogex.h"
#include "CSettingDlg.h"
#include <time.h>
#include "FolderDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#include <xlsxio_read.h>
#include <xlsxio_write.h>
#include "libxl.h"
#include <algorithm>

using namespace libxl;

// CAboutDlg dialog used for App About

std::vector<TaskExcel> g_Tasks;
int g_nTotalThread = 0;
BOOL g_bStop = FALSE;
BOOL g_bStarted = FALSE;

BOOL g_bUseProxy = FALSE;
char g_szProxy[0x100] = { 0 };

typedef
struct TLD {
	char v[0x40];
};

typedef
struct Negative {
	char v[0x100];
};

typedef
struct Priority {
	char v[0x100];
	int n;
};

std::vector<TLD> g_TLDs;
std::vector<Negative> g_Negatives;
std::vector<Priority> g_Priorities;

#define ENABLE_WINDOW(id, x) do { \
	GetDlgItem(id)->EnableWindow(b); \
} while(0)

bool comparePriority(const Priority& a, const Priority& b) {
	return a.n < b.n;
}

void CString2Str(CString source, char* target) {
	for (int i = 0; i < source.GetLength(); i++) {
		target[i] = source.GetAt(i);
	}
}

CString Str2CString(char* source) {
	CString ret;
	for (int i = 0; i < strlen(source); i++) {
		CString tmp;
		tmp.Format(_T("%c"), source[i]);
		ret += tmp;
	}
	return ret;
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
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST_EXCEL_FILES, &CScrapperDlg::OnLvnItemchangedListExcelFiles)
	ON_BN_CLICKED(IDC_BUTTON_ADD_FOLDER, &CScrapperDlg::OnBnClickedButtonAddFolder)
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
	m_ListCtrl.InsertColumn(COL_Rows,		_T("Rows"),			LVCFMT_RIGHT,	60);
	m_ListCtrl.InsertColumn(COL_URL,		_T("Website"),		LVCFMT_CENTER,	70);
	m_ListCtrl.InsertColumn(COL_mail,		_T("E-mail"),		LVCFMT_CENTER,	70);
	m_ListCtrl.InsertColumn(COL_Threads,	_T("Threads"),		LVCFMT_RIGHT,	70);
	m_ListCtrl.InsertColumn(COL_Status,		_T("Status"),		LVCFMT_CENTER,	60);

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
	LoadNegative();
	LoadPriority();
	Load();

	CFont m_Font;
	m_Font.CreatePointFont(140, _T("Arial"));


	SetFont(&m_Font);
	m_ListCtrl.SetFont(&m_Font);
	GetDlgItem(IDC_EDIT3)->SetFont(&m_Font);
	GetDlgItem(IDOK)->SetFont(&m_Font);
	GetDlgItem(IDC_BUTTON_ADD)->SetFont(&m_Font);
	GetDlgItem(IDC_BUTTON_REMOVE)->SetFont(&m_Font);
	GetDlgItem(IDC_BUTTON_EDIT)->SetFont(&m_Font);
	GetDlgItem(IDC_BUTTON_CLEAR)->SetFont(&m_Font);
	GetDlgItem(IDCANCEL)->SetFont(&m_Font);
	GetDlgItem(IDC_CHECK_PROXY)->SetFont(&m_Font);
	GetDlgItem(IDC_BUTTON_ADD_FOLDER)->SetFont(&m_Font);
	
	// GetDlgItem(IDC_EDIT_PROXY)->SetFont(&m_Font);
	

	
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
		Sleep(500);
	}

	// save
	if (!g_bStop) {
		for (int i = 0; i < g_Tasks.size(); i++) {
			if (g_Tasks[i].excel) {
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
								for (int ii = 0; ii < strlen(g_Tasks[i].saves[j]->mail); ii++) {
									if (ii >= 0x100 - 4) {
										break;
									}
									mail[ii] = g_Tasks[i].saves[j]->mail[ii];
								}
								sheet->writeStr(g_Tasks[i].saves[j]->row, g_Tasks[i].mail, mail);
							}
						}
						book->save(szXlsx);
					}
					book->release();
				}
			}
			else {
				int failed = 0;
				std::vector<std::vector<std::string>> vals = pThis->ReadCSV(g_Tasks[i].file, &failed);
				if (failed) {
					AfxMessageBox(_T("Failed to save file > ") + g_Tasks[i].file);
				}
				else {
					char szFile[MAX_PATH] = { 0 };
					CString2Str(g_Tasks[i].file, szFile);
					FILE* fp;
					fopen_s(&fp, szFile, "wb");
					if (!fp) {
						AfxMessageBox(_T("Failed to save file > ") + g_Tasks[i].file);
					}
					else {
						for (int ii = 0; ii < vals.size(); ii++) {
							for (int jj = 0; jj < vals[i].size(); jj++) {
								if (jj != g_Tasks[i].mail) {
									fwrite(vals[ii][jj].c_str(), 1, strlen(vals[ii][jj].c_str()), fp);
								}
								else {
									char szData[0x200] = { 0 };
									for (int k = 0; k < g_Tasks[i].saves.size(); k++) {
										if (g_Tasks[i].saves[k]->row == ii) {
											strcpy_s(szData, g_Tasks[i].saves[k]->mail);
										}
									}
									fwrite(szData, 1, strlen(szData), fp);
								}
								fwrite(",", 1, 1, fp);
							}
							char tail[2] = { 0x0d, 0x0a };
							fwrite(tail, 1, 2, fp);
						}
						fclose(fp);
					}
				}
			}
		}
	}
	pThis->EnableAllButtons(TRUE);
	pThis->Terminated();
	g_bStarted = FALSE;
	return (DWORD)0;
}

int is_valid_email_char(char c) {
	if (c >= 'A' && c <= 'Z')
		return 1;
	if (c >= 'a' && c <= 'z')
		return 1;
	if (c >= '0' && c <= '9')
		return 1;

	return c == '.' || c == '_' || c == '%' || c == '+' || c == '-' || c == '@';
}

int is_valid_email(const char* email) {
	const char* at = strchr(email, '@');
	if (!at) return 0;

	const char* dot = strrchr(email, '.');
	if (!dot || dot < at) return 0;

	if (at == email || *(at + 1) == '\0') return 0;

	if (*(dot + 1) == '\0') return 0;

	int nDot = 0;
	for (int i = 0; i < strlen(email); i++) {
		if (email[i] == '.') {
			nDot++;
		}
	}
	if (nDot != 1) {
		return 0;
	}
	int i = 0;
	for (; i < g_TLDs.size(); i++) {
		char ban[0x40] = { 0 };
		if (g_TLDs[i].v[0] != '.') {
			sprintf_s(ban, ".%s", g_TLDs[i].v);
		}
		else {
			sprintf_s(ban, "%s", g_TLDs[i].v);
		}
		if (strstr(email, ban)) {
			break;
		}
	}
	if (i == g_TLDs.size()) {
		return 0;
	}
	char noprefix[][0x10] = {".png", ".jpg", ".gif", "@sentry.io"};
	for (int i = 0; i < 4; i++) {
		if (strstr(email, noprefix[i])) {
			return 0;
		}
	}

	for (int i = 0; i < g_Negatives.size(); i++) {
		if (strstr(email, g_Negatives[i].v)) {
			return 0;
		}
	}
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
		TaskItem *item = g_Tasks[i].items[g_Tasks[i].pos];
		g_Tasks[i].pos++;
		LeaveCriticalSection(&g_Tasks[i].mutex);


		for (int j = 0; j < 2; j++) {
			char command[256] = { 0 };
			if (!g_bUseProxy) {
				snprintf(command, sizeof(command), "curl -k \"%s%s\"", strstr(item->url, "http") ? "" : (j == 0 ? "https://" : "https://www."), item->url);
			}
			else {
				snprintf(command, sizeof(command), "curl -x %s -k \"%s%s\"", g_szProxy, strstr(item->url, "http") ? "" : (j == 0 ? "https://" : "https://www."), item->url);
			}

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
		
			char buf[0x4000] = { 0 };
			DWORD bytesRead;

			time_t start_time = time(NULL);

			bool found = false;

			int min_offset_mail = 1024 * 1024 * 1024;
			int mail_offset = 0;
			char* priority = 0;
			int priority_offset = 0;


			char finalMail[0x400] = { 0 };

			while (ReadFile(hRead, buf, sizeof(buf) - 1, &bytesRead, NULL) && bytesRead > 0) {
				if (difftime(time(NULL), start_time) >= 60) {
					break;
				}
				if (!priority) {
					for (int ii = 0; ii < g_Priorities.size(); ii++) {
						priority = strstr(buf, g_Priorities[ii].v);
						if (priority) {
							priority_offset = mail_offset + abs(priority - buf);
							break;
						}
					}
				}

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
					if (is_valid_email(s) && strlen(s) <= (0x100 - 4)) {
						int tmp = abs(mail_offset + abs(s - buf) - priority_offset);
						
						if (priority && 0) {
							if (min_offset_mail > tmp) {
								min_offset_mail = tmp;
								strcpy_s(finalMail, s);
								found = true;
							}
						}
						else {
							found = true;

							TaskSave* t = (TaskSave*)malloc(sizeof(TaskSave));
							t->row = item->row;
							memset(t->mail, 0, sizeof(t->mail));
							strcpy_s(t->mail, s);
							EnterCriticalSection(&g_Tasks[i].mutex);
							g_Tasks[i].saves.push_back(t);
							LeaveCriticalSection(&g_Tasks[i].mutex);
						}
					}
					p = next;
				}
				memset(buf, 0, sizeof(buf));
				mail_offset += bytesRead;
			}
		_CONTINUE:
			CloseHandle(hRead);
			CloseHandle(pi.hProcess);
			CloseHandle(pi.hThread);

			if (strstr(item->url, "http") || found) {
				break;
			}
			if (g_bStop) {
				break;
			}
		}
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

	g_bUseProxy = m_UseProxySetting;

	FILE* fp;
	fopen_s(&fp, "proxy.config", "rb");
	if (fp) {
		fread(g_szProxy, 1, sizeof(g_szProxy), fp);
		fclose(fp);
	}

	if (g_bUseProxy && strlen(g_szProxy) == 0) {
		AfxMessageBox(_T("Please check proxy setting"));
		return;
	}

	LoadPriority();

	m_Percent.ShowWindow(SW_SHOW);
	UpdatePercent();
	EnableAllButtons(FALSE);
	GetDlgItem(IDC_CHECK_PROXY)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_PROXY)->EnableWindow(FALSE);

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
	CString filter = _T("Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv||");

	CFileDialog dlg(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, filter, this);

	if (dlg.DoModal() != IDOK) {
		return;
	}

	CString filePath = dlg.GetPathName();
	int n = m_ListCtrl.GetItemCount();
	m_ListCtrl.InsertItem(n, dlg.GetPathName());
	m_ListCtrl.SetItemText(n, COL_Rows, _T("--"));
	m_ListCtrl.SetItemText(n, COL_URL, _T("--"));
	m_ListCtrl.SetItemText(n, COL_mail, _T("--"));
	m_ListCtrl.SetItemText(n, COL_Threads, _T("--"));
	m_ListCtrl.SetItemText(n, COL_Status, _T("Reading..."));

	TaskExcel task;
	task.file = dlg.GetPathName();
	task.pos = 0;
	task.excel = (filePath.Find(_T(".xlsx"), 0) != -1 || filePath.Find(_T(".XLSX"), 0) != -1);
	InitializeCriticalSection(&task.mutex);

	g_Tasks.push_back(task);

	CSettingDlg settingDlg;
	settingDlg.OpenModal(this, TRUE, n, 1, _T("B"), _T("B"), FALSE);
}

void CScrapperDlg::RemoveItem(int index) {
	m_ListCtrl.DeleteItem(index);

	for (int j = 0; j < g_Tasks[index].items.size(); j++) {
		free(g_Tasks[index].items[j]);
	}
	for (int j = 0; j < g_Tasks[index].saves.size(); j++) {
		free(g_Tasks[index].saves[j]);
	}

	g_Tasks.erase(g_Tasks.begin() + index);
}

std::vector<std::string> SplitCSVRow(const std::string& row) {
	std::vector<std::string> fields;
	std::istringstream s(row);
	std::string field;
	bool inQuotes = false;
	char prevChar = '\0';

	for (char c; s.get(c); prevChar = c) {
		if (c == '"') {
			inQuotes = !inQuotes;
			field += c;
		}
		else if (c == ',' && !inQuotes) {
			fields.push_back(field);
			field.clear();
		}
		else {
			field += c;
		}
	}
	fields.push_back(field);

	return fields;
}

void CScrapperDlg::SetThreadColumns(int start, int num, int nThread, CString column, CString mail) {
	int failed = 0;
	for (int i = 0; i < num; i++) {
		int ret = SetThreadColumn(start + i - failed, nThread, column, mail);
		if (!ret) {
			failed++;
		}
	}
}

BOOL CScrapperDlg::SetThreadColumn(int index, int nThread, CString column, CString mail) {
	EnableAllButtons(FALSE);
	CString tmp;
	tmp.Format(_T("%d"), nThread);
	m_ListCtrl.SetItemText(index, COL_Threads, tmp);
	m_ListCtrl.SetItemText(index, COL_URL, column);
	m_ListCtrl.SetItemText(index, COL_mail, mail);

	char szFile[MAX_PATH] = { 0 };
	CString file = m_ListCtrl.GetItemText(index, COL_File);
	CString2Str(file, szFile);

	int urlIdx = column.GetAt(0) - 'A';
	g_Tasks[index].col = urlIdx;
	g_Tasks[index].mail = mail.GetAt(0) - 'A';
	g_Tasks[index].thread = nThread;
	g_Tasks[index].items.clear();
	g_Tasks[index].saves.clear();

	int totalRow = 0;
	int ValidRow = 0;
	int InvalidRow = 0;
	int EmptyRow = 0;
	if (g_Tasks[index].excel) {
		xlsxioreader xlsxioread;
		if ((xlsxioread = xlsxioread_open(szFile)) == NULL) {
			AfxMessageBox(_T("Error opening .xlsx file"));
			RemoveItem(index);
			EnableAllButtons(TRUE);
			return FALSE;
		}

		char* value;
		xlsxioreadersheet sheet;
		const char* sheetname = NULL;
	
		
		if ((sheet = xlsxioread_sheet_open(xlsxioread, sheetname, XLSXIOREAD_SKIP_EMPTY_ROWS)) != NULL) {
			while (xlsxioread_sheet_next_row(sheet)) {
				int tmp = 0;
				while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL) {
					if (tmp == urlIdx) {
						TCHAR* url = (TCHAR*)value;
						if (url[0]) {
							if (wcsstr(url, _T(".")) || wcsstr(url, _T("localhost"))) {
								ValidRow++;
								TaskItem* item = NULL;
								item = (struct TaskItem*)malloc(sizeof(TaskItem));
								memset(item, 0, sizeof(struct TaskItem));
								item->row = totalRow;
								memset(item->url, 0, sizeof(item->url));
								for (int j = 0; j < wcslen(url); j++) {
									if (j >= 0x200 - 4) {
										break;
									}
									item->url[j] = url[j];
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
						if (value) {
							xlsxioread_free(value);
						}
						break;
					}
					tmp++;
					if (value) {
						xlsxioread_free(value);
					}
				}
				totalRow++;
			}
			xlsxioread_sheet_close(sheet);
		}
		xlsxioread_close(xlsxioread);
		
	}
	else {
		int failed = 0;
		std::vector<std::vector<std::string>> vals = ReadCSV(g_Tasks[index].file, &failed);
		if (failed) {
			AfxMessageBox(_T("Error opening .csv file"));
			RemoveItem(index);
			EnableAllButtons(TRUE);
			return FALSE;
		}
		totalRow = vals.size();
		for (int i = 0; i < vals.size(); i++) {
			char url[0x100] = { 0 };
			if (urlIdx >= vals[i].size()) {
				InvalidRow++;
				continue;
			}
			strcpy_s(url, vals[i][urlIdx].c_str());
			if (url[0]) {
				if (strstr(url, ".")) {
					ValidRow++;
					TaskItem* item = NULL;
					item = (struct TaskItem*)malloc(sizeof(TaskItem));
					memset(item, 0, sizeof(struct TaskItem));
					item->row = i;
					memset(item->url, 0, sizeof(item->url));
					strcpy_s(item->url, url);
					g_Tasks[index].items.push_back(item);
				}
				else {
					InvalidRow++;
				}
			}
			else {
				EmptyRow++;
			}
		}
	}
	CString status;
	status.Format(_T("Total: %d, Empty: %d, Valid: %d, Invalid: %d"), totalRow, EmptyRow, ValidRow, InvalidRow);
	m_ListCtrl.SetItemText(index, COL_Status, status);

	status.Format(_T("%d"), totalRow);
	m_ListCtrl.SetItemText(index, COL_Rows, status);
	EnableAllButtons(TRUE);
	return TRUE;
}

std::vector<std::vector<std::string>> CScrapperDlg::ReadCSV(CString filePath, int *failed) {
	char szFile[MAX_PATH] = { 0 };
	CString2Str(filePath, szFile);
	std::string cppString(szFile);

	std::vector<std::vector<std::string>> ret;

	std::ifstream file(cppString);
	if (!file.is_open()) {
		*failed = 1;
		return ret;
	}

	std::string line;
	while (std::getline(file, line)) {
		std::vector<std::string> row;
		std::stringstream lineStream(line);
		std::string cell;

		row = SplitCSVRow(line);
		/*
		while (std::getline(lineStream, cell, ',')) {
			row.push_back(cell);
		}
		*/
		ret.push_back(row);
	}
	file.close();
	*failed = 0;
	return ret;
}

void CScrapperDlg::EnableAllButtons(bool b) {
	ENABLE_WINDOW(IDOK, b);
	ENABLE_WINDOW(IDC_BUTTON_ADD, b);
	ENABLE_WINDOW(IDC_BUTTON_EDIT, b);
	ENABLE_WINDOW(IDC_BUTTON_REMOVE, b);
	ENABLE_WINDOW(IDC_BUTTON_CLEAR, b);
	ENABLE_WINDOW(IDCANCEL, b);
	ENABLE_WINDOW(IDC_BUTTON_ADD_FOLDER, b);
	
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
	CString mail = m_ListCtrl.GetItemText(nItem, COL_mail);

	CSettingDlg settingDlg;
	settingDlg.OpenModal(this, FALSE, nItem, thread, col, mail, FALSE);
}


void CScrapperDlg::OnBnClickedButtonRemove() {
	POSITION pos = m_ListCtrl.GetFirstSelectedItemPosition();
	if (pos == NULL) {
		AfxMessageBox(_T("No item selected."));
		return;
	}
	int nItem = m_ListCtrl.GetNextSelectedItem(pos);
	m_ListCtrl.DeleteItem(nItem);
	RemoveItem(nItem);
}


void CScrapperDlg::OnBnClickedButtonClear() {
	m_ListCtrl.DeleteAllItems();
	for (int i = 0; i < g_Tasks.size(); i++) {
		for (int j = 0; j < g_Tasks[i].items.size(); j++) {
			free(g_Tasks[i].items[j]);
		}
		for (int j = 0; j < g_Tasks[i].saves.size(); j++) {
			free(g_Tasks[i].saves[j]);
		}
	}
	g_Tasks.clear();
}

void CScrapperDlg::Save() {
	FILE* fp;
	fopen_s(&fp, "save.dat", "wb");
	if (fp) {
		int n = g_Tasks.size();
		fwrite(&n, 1, sizeof(int), fp);
		for (int i = 0; i < g_Tasks.size(); i++) {
			char szPath[MAX_PATH] = { 0 };
			CString2Str(g_Tasks[i].file, szPath);
			fwrite(szPath, 1, MAX_PATH, fp);
			fwrite(&g_Tasks[i].thread, 1, sizeof(int), fp);
			fwrite(&g_Tasks[i].col, 1, sizeof(int), fp);
			fwrite(&g_Tasks[i].mail, 1, sizeof(int), fp);
		}
		fclose(fp);
	}
}

void CScrapperDlg::Load() {
	FILE* fp;
	fopen_s(&fp, "save.dat", "rb");
	if (fp) {
		int n = 0;
		fread(&n, 1, sizeof(int), fp);
		for (int i = 0; i < n; i++) {
			char szPath[MAX_PATH] = { 0 };
			fread(szPath, 1, MAX_PATH, fp);
			int thread = 0;
			int col = 0;
			int mail = 0;
			fread(&thread, 1, sizeof(int), fp);
			fread(&col, 1, sizeof(int), fp);
			fread(&mail, 1, sizeof(int), fp);

			m_ListCtrl.InsertItem(m_ListCtrl.GetItemCount(), Str2CString(szPath));
			m_ListCtrl.SetItemText(n, COL_Rows, _T("--"));
			m_ListCtrl.SetItemText(n, COL_URL, _T("--"));
			m_ListCtrl.SetItemText(n, COL_mail, _T("--"));
			m_ListCtrl.SetItemText(n, COL_Threads, _T("--"));
			m_ListCtrl.SetItemText(n, COL_Status, _T("Reading..."));

			TaskExcel task;
			task.file = Str2CString(szPath);
			task.pos = 0;
			task.excel = (task.file.Find(_T(".xlsx"), 0) != -1 || task.file.Find(_T(".XLSX"), 0) != -1);
			InitializeCriticalSection(&task.mutex);

			g_Tasks.push_back(task);

			CString strCol;
			strCol.Format(_T("%c"), col + 'A');
			CString strMail;
			strMail.Format(_T("%c"), mail + 'A');

			SetThreadColumn(m_ListCtrl.GetItemCount() - 1, thread, strCol, strMail);
		}
		fclose(fp);
	}
}

void CScrapperDlg::OnBnClickedCancel() {
	if (g_bStarted) {
		GetDlgItem(IDCANCEL)->EnableWindow(FALSE);
		
		g_bStop = TRUE;
		return;
	}
	Save();
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
	GetDlgItem(IDCANCEL)->EnableWindow(TRUE);
	m_Percent.ShowWindow(SW_HIDE);
	GetDlgItem(IDC_CHECK_PROXY)->EnableWindow(TRUE);
	GetDlgItem(IDC_EDIT_PROXY)->EnableWindow(m_UseProxySetting);
	
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
	g_TLDs.clear();
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
				TLD tld;
				memset(tld.v, 0, sizeof(tld.v));
				strcpy_s(tld.v, cp);
				g_TLDs.push_back(tld);
				m_strTLDs += _T("\r\n");
			}
		}
		fclose(fp);
	}
	UpdateData(FALSE);
}

void CScrapperDlg::LoadNegative() {
	g_Negatives.clear();
	FILE* fp;
	fopen_s(&fp, "Negative.txt", "rb");
	if (fp) {
		char buffer[0x100] = { 0 };
		while (fgets(buffer, sizeof(buffer), fp) != NULL) {
			char cp[0x100] = { 0 };
			int offset = 0;
			for (int i = 0; i < strlen(buffer); i++) {
				if (offset >= (0x100 - 4)) {
					break;
				}
				if (buffer[i] >= 'A' && buffer[i] <= 'Z') {
					cp[offset++] = buffer[i] - 'A' + 'a';
				}
				else {
					if (buffer[i] != 0x0d && buffer[i] != 0x0a) {
						cp[offset++] = buffer[i];
					}
				}
			}
			memset(buffer, 0, sizeof(buffer));

			if (strlen(cp)) {
				Negative neg;
				memset(neg.v, 0, sizeof(neg.v));
				strcpy_s(neg.v, cp);
				g_Negatives.push_back(neg);
			}
		}
		fclose(fp);
	}
	UpdateData(FALSE);
}


void CScrapperDlg::LoadPriority() {
	g_Priorities.clear();
	FILE* fp;
	fopen_s(&fp, "Priority.txt", "rb");
	if (fp) {
		char buffer[0x100] = { 0 };
		while (fgets(buffer, sizeof(buffer), fp) != NULL) {
			char cp[0x100] = { 0 };
			int offset = 0;
			for (int i = 0; i < strlen(buffer); i++) {
				if (offset >= (0x100 - 4)) {
					break;
				}
				if (buffer[i] != 0x0d && buffer[i] != 0x0a) {
					cp[offset++] = buffer[i];
				}
			}
			memset(buffer, 0, sizeof(buffer));

			if (strlen(cp)) {
				Priority p;
				memset(p.v, 0, sizeof(p.v));
				strcpy_s(p.v, cp);
				p.n = 0;
				g_Priorities.push_back(p);
			}
		}
		fclose(fp);
	}
	// std::sort(g_Priorities.begin(), g_Priorities.end(), comparePriority);
	UpdateData(FALSE);
}


void CScrapperDlg::OnLvnItemchangedListExcelFiles(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
}


void CScrapperDlg::OnBnClickedButtonAddFolder() {
	CFolderPickerDialog dlg;
	if (dlg.DoModal() == IDOK) {
		CString folderPath = dlg.GetFolderPath();
		CFileFind finder;
		CString searchPath = folderPath + _T("\\*.xlsx");

		BOOL bWorking = finder.FindFile(searchPath);

		int start = -1;
		int num = 0;

		while (bWorking) {
			bWorking = finder.FindNextFile();
			if (!finder.IsDots() && !finder.IsDirectory()) {
				CString path = finder.GetFilePath();

				int n = m_ListCtrl.GetItemCount();
				if (start == -1) {
					start = n;
				}
				m_ListCtrl.InsertItem(n, path);
				m_ListCtrl.SetItemText(n, COL_Rows, _T("--"));
				m_ListCtrl.SetItemText(n, COL_URL, _T("--"));
				m_ListCtrl.SetItemText(n, COL_mail, _T("--"));
				m_ListCtrl.SetItemText(n, COL_Threads, _T("--"));
				m_ListCtrl.SetItemText(n, COL_Status, _T("Reading..."));

				TaskExcel task;
				task.file = path;
				task.pos = 0;
				task.excel = (path.Find(_T(".xlsx"), 0) != -1 || path.Find(_T(".XLSX"), 0) != -1);
				InitializeCriticalSection(&task.mutex);
				g_Tasks.push_back(task);
				num++;
			}
		}
		finder.Close();

		if (num == 0) {
			AfxMessageBox(_T("No xlsx file."));
			return;
		}
		CSettingDlg settingDlg;
		settingDlg.OpenModal(this, TRUE, 0, 1, _T("B"), _T("B"), TRUE, start, num);
	}
}
