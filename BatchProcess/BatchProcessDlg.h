// BatchProcessDlg.h : header file
//

#pragma once
#include "lead.h"
//#include "leads.h"
#include "hlistbox.h"
#include "afxwin.h"
//#include ".\statelist.h"
//#include "processorders.h"
#include "ESMTP.h"
#include "order.h"
#include "leaddb.h"
#include "processthread.h"
#include ".\fax.h"
#include ".\faxthread.h"
#include ".\faxjob.h"
#include "timerthread.h"
#include "faxjob.h"
#include "afxcmn.h"
#include "manualprocessthread.h"
#include "logdb.h"
#include ".\logentry.h"
#include ".\logthread.h"
#include "logentry.h"
#include "loggerthread.h"
#include "processqueue.h"
#include ".\stats.h"
#import "C:\Program Files\Common Files\System\ADO\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")

// CBatchProcessDlg dialog
class CBatchProcessDlg : public CDialog
{
// Construction
public:
	CBatchProcessDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_BATCHPROCESS_DIALOG };

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
	CHListBox m_output;
	afx_msg void OnBnClickedButtonRead();
	void PrintOut(CString str);
private:
	//CArray<Lead,Lead&> leads;
	//Leads leads;
	//LeadList leads;
	//LeadList verifiedLeads;
public:
	long ReadLeads(void);
	CString GetCString(_bstr_t bstrStart);
	afx_msg void OnEnterIdle(UINT nWhy, CWnd* pWho);
	CString GetLong_String(long num);
	CString GetCString(_variant_t var);
	
	CString GetField(_RecordsetPtr rs, CString field);
	CString GetInt_String(int num);
	CString GetField(_RecordsetPtr rs, CString field, int vt_type);
	void PrintOut(CLead aLead);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	afx_msg void OnBnClickedButtonPrintList();
	afx_msg void OnBnClickedButtonPrintInternet();
	CButton m_verified;
	afx_msg void OnBnClickedButtonTest();
	afx_msg LRESULT OnReadLead(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnInitDB(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnStartLeadFileRead(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnLeadsCount(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnComError(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnFailCreateRecordInstance(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnStartFaxString(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnCharFaxString(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnEndFaxString(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnFaxSuccess(WPARAM wp, LPARAM lp);
	//afx_msg LRESULT OnLogWriteDone(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnFaxFail(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnResetTimer(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnStartTimer(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnFireTimer(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnEmailFailChar(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnEmailFailEnd(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnEmailFailStart(WPARAM wp, LPARAM lp);
	

	CEdit m_states;
private:
	// Worker Thread for order processing
	//CProcessOrders processOrders;
public:
	
	afx_msg void OnDestroy();
	afx_msg void OnBnClickedCancel();
protected:
	virtual void OnCancel();
public:
	afx_msg void OnBnClickedButtonTestOrder();
	COrder testOrder;
	afx_msg void OnBnClickedCheckVerified();
	afx_msg void OnEnChangeEditStates();
	CEdit m_companyid;
	afx_msg void OnEnChangeEditCompanyid();
	CEdit m_contact;
	afx_msg void OnEnChangeEditContact();
	CEdit m_orderid;
	afx_msg void OnEnChangeEditOrderid();
	afx_msg void OnBnClickedOk();
	CEdit m_faxnumber;
	afx_msg void OnEnChangeEditFaxnumber();
	CEdit m_quantity;
	afx_msg void OnEnChangeEditQuantity();
	CEdit m_email;
	afx_msg void OnEnChangeEditEmail();
	CString GetText(CEdit *aBox);
	CEdit m_agent;
	afx_msg void OnEnChangeEditAgent();
	CEdit m_rate;
	afx_msg void OnEnChangeEditRate();
	CEdit m_zip;
	afx_msg void OnEnChangeEditZipFilter();
	CEdit m_loan;
	afx_msg void OnEnChangeEditDesiredLoanFilter();
	CEdit m_areacodes;
	afx_msg void OnEnChangeEditAreacodeFilter();
	CButton m_onepass;
	afx_msg void OnBnClickedCheckOnepass();
	CButton m_excel;
	afx_msg void OnBnClickedCheckExcel();
	CButton m_fax;
	afx_msg void OnBnClickedCheckFax();
	CButton m_word;
	afx_msg void OnBnClickedCheckWord();
	CButton m_onelead;
	afx_msg void OnBnClickedCheckOnelead();
	void LoadScreenFromOrder(COrder *aOrder);
	
	CEdit m_cashout;
	CEdit m_2ndmortgage;
	CEdit m_debtcon;
	_RecordsetPtr pRstInternetOrders;
	afx_msg void OnBnClickedButtonOpenInternetOrders();
	CButton m_emailnote;
	CButton m_print;
	CEdit m_lookingto;
	CEdit m_lang;
	CEdit m_creditscore;
	CEdit m_hometype;
	CEdit m_location;
	CButton m_priority;
	CButton m_mini1003;
	CButton m_adcampaign;
	CEdit m_agentid;
	CButton m_subprime;
	CEdit m_sql;
protected:
	CString message;
public:
	LRESULT OnStartMessageString(WPARAM wp, LPARAM lp);
	LRESULT OnEndMessageString(WPARAM wp, LPARAM lp);
	LRESULT OnCharMessageString(WPARAM wp, LPARAM lp);
	
	afx_msg void OnBnClickedButtonOpenInternetOnepass();
	//CArray<CLead,CLead&> leadsMasterList;
	CLeadDB leadDBMap;
	CEdit m_SFR;
	CEdit m_Mobile;
	CEdit m_Commercial;
	CEdit m_city;
	CEdit m_ltv;
	CButton m_ltvAbove;
	//CProcessThread processOrdersThread;
	afx_msg void OnBnClickedButtonProcess();
	afx_msg void OnBnClickedButtonStop();
	LRESULT OnOrderTrans(WPARAM wp, LPARAM lp);
	void SetEditBox(CEdit *pEdit, int iMessage);
	void SetCheckBox(CButton* pCheck, int state);
	CString sSQL;
	void SetSQL(int iMessage);
	CEdit m_lastOrderQuantity;
	CStatic m_totalOrdersLabel;
	CEdit m_companyTotalOrders;
	CEdit m_yield;
	bool SendMail(CString brokerEmail, CString agentname, CString agentEmail, CString company, CString contact, int div, CString orderid);
	bool testing;
	CString faxmessage;
	bool SendFax(CString faxnum, CString agentname, CString agentEmail, CString company, CString contact, int div, CString filename);
	CFaxThread* pt_faxer;
	CArray<CFaxJob,CFaxJob> myfaxJobs;
	CArray<CLogEntry, CLogEntry> myLogEntrys;
	bool sendingFax;
	void SendNextFax(void);
	void QueueFaxJob(CFaxJob aJob);
	int iFailCount;
	int iCountDown;
	void ResetTimer(void);
	CTimerThread killtimer;
	//void AddToFailLog(int div, CString str);
	void MarkOrderComplete(CFaxJob aJob);
	CFaxJob theFaxJob;
	LRESULT OnFaxJobStart(WPARAM wp, LPARAM lp);
	LRESULT OnFaxJobEnd(WPARAM wp, LPARAM lp);
	LRESULT OnFaxJobCamID(WPARAM wp, LPARAM lp);
	LRESULT OnFaxJobOrderID(WPARAM wp, LPARAM lp);
	LRESULT OnFaxJobDiv(WPARAM wp, LPARAM lp);
	LRESULT OnFaxJobSendCoverpage(WPARAM wp, LPARAM lp);
	LRESULT OnFaxJobAgentEmailChar(WPARAM wp, LPARAM lp);
	LRESULT OnFaxJobAgentNameChar(WPARAM wp, LPARAM lp);
	LRESULT OnFaxJobCompanyChar(WPARAM wp, LPARAM lp);
	LRESULT OnFaxJobContactChar(WPARAM wp, LPARAM lp);
	LRESULT OnFaxJobFaxNumChar(WPARAM wp, LPARAM lp);
	LRESULT OnFaxJobFilenameChar(WPARAM wp, LPARAM lp);
	LRESULT OnFaxJobStatesChar(WPARAM wp, LPARAM lp);
	CEdit m_campaignid;
	afx_msg void OnSize(UINT nType, int cx, int cy);
	LRESULT OnProcessVerified(WPARAM wp, LPARAM lp);
	LRESULT OnProcessInternet(WPARAM wp, LPARAM lp);
	LRESULT OnProcessOnePass(WPARAM wp, LPARAM lp);
	LRESULT OnProcessPriority(WPARAM wp, LPARAM lp);
	LRESULT OnProcessSentZero(WPARAM wp, LPARAM lp);
	LRESULT OnProcessNewCustomer(WPARAM wp, LPARAM lp);
	LRESULT OnProcessNormal(WPARAM wp, LPARAM lp);
	LRESULT OnRetrySending(WPARAM wp, LPARAM lp);
	CEdit m_leadtype;
	CEdit m_ordertype;
	LRESULT OnProcessDone(WPARAM wp, LPARAM lp);
	CProgressCtrl m_progress;
	LRESULT OnProgress(WPARAM wp, LPARAM lp);
	LRESULT OnProgressAll(WPARAM wp, LPARAM lp);
	CManualProcessThread manualProcessThread;
	afx_msg void OnBnClickedButtonVerified();
	afx_msg void OnBnClickedButtonInternet();
	afx_msg void OnBnClickedButtonOnepass();
	int m_groupIndex;
	afx_msg void OnCbnSelchangeComboGroup();
	CComboBox m_group;
	void OnManualProcess(void);
	CButton m_button_all;
	CButton m_button_stop;
	CButton m_button_verified;
	CButton m_button_internet;
	CButton m_button_onepass;
	void ProcessButtonsEnable(bool bState);
	//CLogDB myLog;
	//void AddToLog(int div, CString str);
	//void WriteNextEntry(void);
	//CLogThread* pt_logger;
	//bool writingLog;
	//CLogEntry myEmailFailLogEntry;
	//CLoggerThread theLogger;
	void PrintFail(CString str);
	afx_msg void OnFileExit();
	afx_msg void OnHelpAbout();
	CProgressCtrl m_progress_all;
	CProcessQueue theProcessQueue;
	bool autorun;
	bool CheckForDone(void);
	bool processing;
	afx_msg void OnBnClickedButtonRetry();
	CButton m_retry;
	afx_msg void OnBnClickedButtonGenOrders();
	CButton m_gen;
	void WriteLog(CString str);
	afx_msg void OnBnClickedButtonStats();
	CButton m_stats;
	CComboBox m_comboStates;
	afx_msg void OnCbnSelchangeComboStates();
	int m_comboStatesIndex;
};
