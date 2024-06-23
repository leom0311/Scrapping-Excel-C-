#pragma once
#include "afxdialogex.h"


// CSettingDlg dialog
class CScrapperDlg;
class CSettingDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CSettingDlg)

public:
	CSettingDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~CSettingDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_SETTING };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CScrapperDlg* m_wndParent;

	BOOL m_bIsNew;
	int m_nID;
	int m_nThread;
	CString m_sColumn;
	CString m_sMail;
public:
	void OpenModal(CScrapperDlg *parent, BOOL isNew, int nId, int nThread, CString sColumn, CString mail);
public:
	virtual BOOL OnInitDialog();
	CComboBox m_ThreadCnt;
	CComboBox m_UrlColumn;
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	CComboBox m_MailColumn;
};
