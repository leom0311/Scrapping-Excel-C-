// CSettingDlg.cpp : implementation file
//

#include "pch.h"
#include "Scrapper.h"
#include "afxdialogex.h"
#include "CSettingDlg.h"
#include "ScrapperDlg.h"

// CSettingDlg dialog

IMPLEMENT_DYNAMIC(CSettingDlg, CDialogEx)

CSettingDlg::CSettingDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DIALOG_SETTING, pParent)
{

}

CSettingDlg::~CSettingDlg()
{
}

void CSettingDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_COMBO_THREAD, m_ThreadCnt);
	DDX_Control(pDX, IDC_COMBO_URL_COLUMN, m_UrlColumn);
}


BEGIN_MESSAGE_MAP(CSettingDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &CSettingDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CSettingDlg::OnBnClickedCancel)
END_MESSAGE_MAP()


// CSettingDlg message handlers

BOOL CSettingDlg::OnInitDialog() {
	CDialogEx::OnInitDialog();

	m_ThreadCnt.AddString(_T("1"));
	m_ThreadCnt.AddString(_T("8"));
	m_ThreadCnt.AddString(_T("16"));
	m_ThreadCnt.AddString(_T("64"));
	m_ThreadCnt.AddString(_T("128"));

	for (int i = 'B'; i <= 'Z'; i++) {
		CString tmp;
		tmp.Format(_T("%c"), i);
		m_UrlColumn.AddString(tmp);
	}

	CString strThread;
	strThread.Format(_T("%d"), m_nThread);
	int index = m_ThreadCnt.FindStringExact(-1, strThread);
	if (index != CB_ERR) {
		m_ThreadCnt.SetCurSel(index);
	}
	else {
		m_ThreadCnt.SetCurSel(0);
	}

	index = m_UrlColumn.FindStringExact(-1, m_sColumn);
	if (index != CB_ERR) {
		m_UrlColumn.SetCurSel(index);
	}
	else {
		m_UrlColumn.SetCurSel(0);
	}

	return TRUE;
}

void CSettingDlg::OpenModal(CScrapperDlg* parent, BOOL isNew, int nId, int nThread, CString sColumn) {
	m_wndParent = parent;
	m_bIsNew = isNew;
	m_nID = nId;
	m_nThread = nThread;
	m_sColumn = sColumn;
	DoModal();
}

void CSettingDlg::OnBnClickedOk() {
	int nSel = m_ThreadCnt.GetCurSel();
	int nThread = 1;
	if (nSel != CB_ERR) {
		CString strSelectedItem;
		m_ThreadCnt.GetLBText(nSel, strSelectedItem);
		nThread = _ttoi(strSelectedItem);
	}
	else {
		return;
	}

	nSel = m_UrlColumn.GetCurSel();
	CString strSelectedColumn;
	if (nSel != CB_ERR) {
		m_UrlColumn.GetLBText(nSel, strSelectedColumn);
	}
	else {
		return;
	}
	
	m_wndParent->SetThreadColumn(m_nID, nThread, strSelectedColumn);
	CDialogEx::OnOK();
}

void CSettingDlg::OnBnClickedCancel() {
	if (m_bIsNew) {
		m_wndParent->RemoveItem(m_nID);
	}
	CDialogEx::OnCancel();
}
