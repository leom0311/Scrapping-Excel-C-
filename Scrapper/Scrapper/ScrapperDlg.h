
// ScrapperDlg.h : header file
//

#pragma once

enum {
	COL_File = 0,
	COL_Rows,
	COL_URL,
	COL_Threads,
	COL_Status
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
public:
	CListCtrl m_ListCtrl;
	afx_msg void OnBnClickedOk();
};
