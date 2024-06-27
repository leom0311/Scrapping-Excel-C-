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
	, m_Website(_T(""))
	, m_Mail(_T(""))
{

}

CSettingDlg::~CSettingDlg()
{
}

void CSettingDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_COMBO_THREAD, m_ThreadCnt);
	// DDX_Control(pDX, IDC_COMBO_URL_COLUMN, m_UrlColumn);
	// DDX_Control(pDX, IDC_COMBO_MAIL_COLUMN, m_MailColumn);
	DDX_Text(pDX, IDC_EDIT1_WEBSITE, m_Website);
	DDX_Text(pDX, IDC_EDIT3_EMAIL, m_Mail);
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

	/*
	for (int i = 'B'; i <= 'Z'; i++) {
		CString tmp;
		tmp.Format(_T("%c"), i);
		m_UrlColumn.AddString(tmp);
		m_MailColumn.AddString(tmp);
	}
	*/

	
	CString strThread;
	strThread.Format(_T("%d"), m_nThread);
	int index = m_ThreadCnt.FindStringExact(-1, strThread);
	if (index != CB_ERR) {
		m_ThreadCnt.SetCurSel(index);
	}
	else {
		m_ThreadCnt.SetCurSel(0);
	}

	/*
	index = m_MailColumn.FindStringExact(-1, m_sMail);
	if (index != CB_ERR) {
		m_MailColumn.SetCurSel(index);
	}
	else {
		m_MailColumn.SetCurSel(0);
	}

	index = m_UrlColumn.FindStringExact(-1, m_sColumn);
	if (index != CB_ERR) {
		m_UrlColumn.SetCurSel(index);
	}
	else {
		m_UrlColumn.SetCurSel(0);
	}
	*/
	m_Website = m_sColumn;
	m_Mail = m_sMail;

	CFont m_Font;
	m_Font.CreatePointFont(140, _T("Arial"));

	SetFont(&m_Font);
	GetDlgItem(IDC_STATIC)->SetFont(&m_Font);
	GetDlgItem(IDC_STATIC_1)->SetFont(&m_Font);
	GetDlgItem(IDC_STATIC_2)->SetFont(&m_Font);
	GetDlgItem(IDC_COMBO_THREAD)->SetFont(&m_Font);
	GetDlgItem(IDC_EDIT1_WEBSITE)->SetFont(&m_Font);
	GetDlgItem(IDC_EDIT3_EMAIL)->SetFont(&m_Font);

	UpdateData(FALSE);
	return TRUE;
}

void CSettingDlg::OpenModal(CScrapperDlg* parent, BOOL isNew, int nId, int nThread, CString sColumn, CString mail, bool fromFolder, int start, int num) {
	m_wndParent = parent;
	m_bIsNew = isNew;
	m_nID = nId;
	m_nThread = nThread;
	m_sColumn = sColumn;
	m_sMail = mail;

	m_bFromFolder = fromFolder;
	m_nStart = start;
	m_nNum = num;
	DoModal();
}

void CSettingDlg::OnBnClickedOk() {
	UpdateData();
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
	if (m_Website == _T("")) {
		AfxMessageBox(_T("Input Website column"));
		return;
	}
	if (m_Mail == _T("")) {
		AfxMessageBox(_T("Input E-main column"));
		return;
	}

	/*
	nSel = m_UrlColumn.GetCurSel();
	CString strSelectedColumn;
	if (nSel != CB_ERR) {
		m_UrlColumn.GetLBText(nSel, strSelectedColumn);
	}
	else {
		return;
	}
	
	nSel = m_MailColumn.GetCurSel();
	CString strSelectedMail;
	if (nSel != CB_ERR) {
		m_MailColumn.GetLBText(nSel, strSelectedMail);
	}
	else {
		return;
	}
	*/
	CString site;
	site.AppendChar(m_Website.GetAt(0));

	CString mail;
	mail.AppendChar(m_Mail.GetAt(0));
	if (m_bFromFolder) {
		m_wndParent->SetThreadColumns(m_nStart, m_nNum, nThread, site, mail);
	}
	else {
		m_wndParent->SetThreadColumn(m_nID, nThread, site, mail);
	}
	CDialogEx::OnOK();
}

void CSettingDlg::OnBnClickedCancel() {
	if (m_bIsNew) {
		m_wndParent->RemoveItem(m_nID);
	}
	CDialogEx::OnCancel();
}
