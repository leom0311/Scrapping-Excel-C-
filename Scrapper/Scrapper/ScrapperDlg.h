
// ScrapperDlg.h : header file
//

#pragma once

#include <vector>
enum {
	COL_File = 0,
	COL_Rows,
	COL_URL,
	COL_Threads,
	COL_Status,
	COL_CNT
};

struct TaskItem {
	int row;
	char url[0x100];
};
struct TaskSave {
	int row;
	char mail[0x40];
};

struct TaskExcel {
	CRITICAL_SECTION mutex;
	CString file;
	int thread;
	int col;

	int pos;
	std::vector<TaskItem> items;
	std::vector<TaskSave> saves;
};

// CScrapperDlg dialog
class CScrapperDlg : public CDialogEx
{
// Construction
public:
	CScrapperDlg(CWnd* pParent = nullptr);	// standard constructor

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_SCRAPPER_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	void AdjustListColumn(CListCtrl* list);
	void SetThreadColumn(int index, int nThread, CString column);
	void RemoveItem(int index);
	void EnableAllButtons(bool b);
	void UpdatePercent();
	void Terminated();
	void LoadTLD();

	static DWORD WINAPI ThreadScrapping(LPVOID lpParam);
	static DWORD WINAPI ThreadMonitor(LPVOID lpParam);
public:
	CListCtrl m_ListCtrl;
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedButtonAdd();
	afx_msg void OnBnClickedButtonEdit();
	afx_msg void OnBnClickedButtonRemove();
	afx_msg void OnBnClickedButtonClear();
	afx_msg void OnBnClickedCancel();
	CProgressCtrl m_Percent;
	afx_msg void OnBnClickedCheckProxy();
	BOOL m_UseProxySetting;
	CEdit m_editProxy;
	afx_msg void OnEnChangeEditProxy();
	CString m_strTLDs;
};
