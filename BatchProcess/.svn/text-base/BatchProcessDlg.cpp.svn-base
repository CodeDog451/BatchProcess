// BatchProcessDlg.cpp : implementation file
//

#include "stdafx.h"
#include "BatchProcess.h"
#include "BatchProcessDlg.h"
#include ".\batchprocessdlg.h"
#include <stdlib.h>
#include ".\order.h"
#include ".\rtfwriter.h"
#include ".\location.h"
#include ".\settings.h"
#include ".\csvwriter.h"
#include ".\genorders.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CBatchProcessDlg dialog



CBatchProcessDlg::CBatchProcessDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CBatchProcessDlg::IDD, pParent)	
	, message(_T(""))
	, sSQL(_T(""))
	, testing(false)
	, faxmessage(_T(""))
	, pt_faxer(NULL)
	, sendingFax(false)
	, iFailCount(0)
	, iCountDown(0)
	, m_groupIndex(0)
	//, pt_logger(NULL)
	//, writingLog(false)
	, autorun(false)
	, processing(false)
	, m_comboStatesIndex(0)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CBatchProcessDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_OUTPUT, m_output);
	DDX_Control(pDX, IDC_CHECK_VERIFIED, m_verified);
	DDX_Control(pDX, IDC_EDIT_STATES, m_states);
	DDX_Control(pDX, IDC_EDIT_COMPANYID, m_companyid);
	DDX_Control(pDX, IDC_EDIT_CONTACT, m_contact);
	DDX_Control(pDX, IDC_EDIT_ORDERID, m_orderid);
	DDX_Control(pDX, IDC_EDIT_FAXNUMBER, m_faxnumber);
	DDX_Control(pDX, IDC_EDIT_QUANTITY, m_quantity);
	DDX_Control(pDX, IDC_EDIT_EMAIL, m_email);
	DDX_Control(pDX, IDC_EDIT_AGENT, m_agent);
	DDX_Control(pDX, IDC_EDIT_RATE, m_rate);
	DDX_Control(pDX, IDC_EDIT_ZIP_FILTER, m_zip);
	DDX_Control(pDX, IDC_EDIT_DESIRED_LOAN_FILTER, m_loan);
	DDX_Control(pDX, IDC_EDIT_AREACODE_FILTER, m_areacodes);
	DDX_Control(pDX, IDC_CHECK_ONEPASS, m_onepass);
	DDX_Control(pDX, IDC_CHECK_EXCEL, m_excel);
	DDX_Control(pDX, IDC_CHECK_FAX, m_fax);
	DDX_Control(pDX, IDC_CHECK_WORD, m_word);
	DDX_Control(pDX, IDC_CHECK_ONELEAD, m_onelead);

	DDX_Control(pDX, IDC_EDIT_CASHOUT, m_cashout);
	DDX_Control(pDX, IDC_EDIT_2NDMORTGAGE, m_2ndmortgage);
	DDX_Control(pDX, IDC_EDIT_DEBTCON, m_debtcon);
	DDX_Control(pDX, IDC_CHECK_EMAILNOTE, m_emailnote);
	DDX_Control(pDX, IDC_CHECK_PRINT, m_print);
	DDX_Control(pDX, IDC_EDIT_LOOKINGTO, m_lookingto);
	DDX_Control(pDX, IDC_EDIT_LANG, m_lang);
	DDX_Control(pDX, IDC_EDIT_CREDIT, m_creditscore);
	DDX_Control(pDX, IDC_EDIT_HOMETYPE, m_hometype);
	DDX_Control(pDX, IDC_EDIT_LOCATION, m_location);
	DDX_Control(pDX, IDC_CHECK_PRIORITY, m_priority);
	DDX_Control(pDX, IDC_CHECK_MINI1003, m_mini1003);
	DDX_Control(pDX, IDC_CHECK_ADCAMPAIGN, m_adcampaign);
	DDX_Control(pDX, IDC_EDIT_AGENTID, m_agentid);
	DDX_Control(pDX, IDC_CHECK_SUBPRIME, m_subprime);
	DDX_Control(pDX, IDC_EDIT_SQL, m_sql);
	DDX_Control(pDX, IDC_EDIT_SFR, m_SFR);
	DDX_Control(pDX, IDC_EDIT_MOBILE, m_Mobile);
	DDX_Control(pDX, IDC_EDIT_COMMERCIAL, m_Commercial);
	DDX_Control(pDX, IDC_EDIT_CITY, m_city);
	DDX_Control(pDX, IDC_EDIT_LTV, m_ltv);
	DDX_Control(pDX, IDC_CHECK_ABOVE, m_ltvAbove);
	DDX_Control(pDX, IDC_EDIT_LAST_ORDER_QUANTITY, m_lastOrderQuantity);
	DDX_Control(pDX, IDC_STATIC_COMPANY_COUNT_ORDERS, m_totalOrdersLabel);
	DDX_Control(pDX, IDC_EDIT_COMPANY_ORDERS_COUNT, m_companyTotalOrders);
	DDX_Control(pDX, IDC_EDIT_YIELD, m_yield);
	DDX_Control(pDX, IDC_EDIT_CAMPAIGNID, m_campaignid);
	DDX_Control(pDX, IDC_EDIT_LEADTYPE, m_leadtype);
	DDX_Control(pDX, IDC_EDIT_ORDERTYPE, m_ordertype);
	DDX_Control(pDX, IDC_PROGRESS_READ, m_progress);
	DDX_CBIndex(pDX, IDC_COMBO_GROUP, m_groupIndex);
	DDX_Control(pDX, IDC_COMBO_GROUP, m_group);
	DDX_Control(pDX, IDC_BUTTON_PROCESS, m_button_all);
	DDX_Control(pDX, IDC_BUTTON_STOP, m_button_stop);
	DDX_Control(pDX, IDC_BUTTON_VERIFIED, m_button_verified);
	DDX_Control(pDX, IDC_BUTTON_INTERNET, m_button_internet);
	DDX_Control(pDX, IDC_BUTTON_ONEPASS, m_button_onepass);
	DDX_Control(pDX, IDC_PROGRESS_OVERALL, m_progress_all);
	DDX_Control(pDX, IDC_BUTTON_RETRY, m_retry);
	DDX_Control(pDX, IDC_BUTTON_GEN_ORDERS, m_gen);
	DDX_Control(pDX, IDC_BUTTON_STATS, m_stats);
	DDX_Control(pDX, IDC_COMBO_STATES, m_comboStates);
	DDX_CBIndex(pDX, IDC_COMBO_STATES, m_comboStatesIndex);
}

BEGIN_MESSAGE_MAP(CBatchProcessDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON_READ, OnBnClickedButtonRead)
	ON_WM_ENTERIDLE()
	ON_WM_SIZING()
	ON_BN_CLICKED(IDC_BUTTON_PRINT_LIST, OnBnClickedButtonPrintList)
	ON_BN_CLICKED(IDC_BUTTON_PRINT_INTERNET, OnBnClickedButtonPrintInternet)
	ON_BN_CLICKED(IDC_BUTTON_TEST, OnBnClickedButtonTest)
	ON_MESSAGE(ID_LEAD_READ, OnReadLead)
	ON_MESSAGE(ID_INIT_DB_CONNECT, OnInitDB)
	ON_MESSAGE(START_LEAD_FILE_READ, OnStartLeadFileRead)
	ON_MESSAGE(ID_LEADS_COUNT, OnLeadsCount)
	ON_MESSAGE(ID_COM_ERR, OnComError)
	ON_MESSAGE(FAIL_CREATE_RECORD_INSTANCE, OnFailCreateRecordInstance)
	ON_MESSAGE(S_START_MESSAGE, OnStartMessageString)
	ON_MESSAGE(S_CHAR_MESSAGE, OnCharMessageString)
	ON_MESSAGE(S_END_MESSAGE, OnEndMessageString)
	ON_MESSAGE(S_ORDER_TRANS, OnOrderTrans)
	ON_MESSAGE(S_START_FAX_STRING, OnStartFaxString)
	ON_MESSAGE(S_CHAR_FAX, OnCharFaxString)
	ON_MESSAGE(S_END_FAX_STRING, OnEndFaxString)
	ON_MESSAGE(S_FAX_SUCCESS, OnFaxSuccess)
	//ON_MESSAGE(LOG_WRITE_DONE, OnLogWriteDone)
	ON_MESSAGE(S_FAX_FAIL, OnFaxFail)
	ON_MESSAGE(S_TIMER_START, OnStartTimer)
	ON_MESSAGE(S_TIMER_FIRE, OnFireTimer)
	ON_MESSAGE(S_TIMER_RESET, OnResetTimer)
	ON_MESSAGE(EMAIL_FAIL_CHAR, OnEmailFailChar)
	ON_MESSAGE(EMAIL_FAIL_END, OnEmailFailEnd)
	ON_MESSAGE(EMAIL_FAIL_START, OnEmailFailStart)

	ON_MESSAGE(OB_FAXJOB_START, OnFaxJobStart)
	ON_MESSAGE(OB_FAXJOB_END, OnFaxJobEnd)
	ON_MESSAGE(OB_FAXJOB_CAMID, OnFaxJobCamID)
	ON_MESSAGE(OB_FAXJOB_ORDERID, OnFaxJobOrderID)
	ON_MESSAGE(OB_FAXJOB_DIV, OnFaxJobDiv)
	ON_MESSAGE(OB_FAXJOB_SENDCOVERPAGE, OnFaxJobSendCoverpage)
	
	ON_MESSAGE(OB_FAXJOB_AGENTEMAIL_CHAR, OnFaxJobAgentEmailChar)
	ON_MESSAGE(OB_FAXJOB_AGENTNAME_CHAR, OnFaxJobAgentNameChar)
	ON_MESSAGE(OB_FAXJOB_COMPANY_CHAR, OnFaxJobCompanyChar)
	ON_MESSAGE(OB_FAXJOB_CONTACT_CHAR, OnFaxJobContactChar)
	ON_MESSAGE(OB_FAXJOB_FAXNUM_CHAR, OnFaxJobFaxNumChar)
	ON_MESSAGE(OB_FAXJOB_FILENAME_CHAR, OnFaxJobFilenameChar)
	ON_MESSAGE(OB_FAXJOB_STATES_CHAR, OnFaxJobStatesChar)

	ON_MESSAGE(PO_VERIFIED, OnProcessVerified)
	ON_MESSAGE(PO_INTERNET, OnProcessInternet)
	ON_MESSAGE(PO_ONEPASS, OnProcessOnePass)
	ON_MESSAGE(PO_PRIORITY, OnProcessPriority)
	ON_MESSAGE(PO_SENTZERO, OnProcessSentZero)
	ON_MESSAGE(PO_NEWCUSTOMER, OnProcessNewCustomer)
	ON_MESSAGE(PO_NORMAL, OnProcessNormal)
	ON_MESSAGE(PO_DONE, OnProcessDone)
	ON_MESSAGE(ID_PROGRESS, OnProgress)
	ON_MESSAGE(ID_PROGRESS_OVERALL, OnProgressAll)
	
	ON_MESSAGE(PO_RETRY_SENDING, OnRetrySending)

	ON_WM_DESTROY()
	ON_BN_CLICKED(IDCANCEL, OnBnClickedCancel)
	ON_BN_CLICKED(IDC_BUTTON_TEST_ORDER, OnBnClickedButtonTestOrder)
	ON_BN_CLICKED(IDC_CHECK_VERIFIED, OnBnClickedCheckVerified)
	ON_EN_CHANGE(IDC_EDIT_STATES, OnEnChangeEditStates)
	ON_EN_CHANGE(IDC_EDIT_COMPANYID, OnEnChangeEditCompanyid)
	ON_EN_CHANGE(IDC_EDIT_CONTACT, OnEnChangeEditContact)
	ON_EN_CHANGE(IDC_EDIT_ORDERID, OnEnChangeEditOrderid)
	ON_BN_CLICKED(IDOK, OnBnClickedOk)
	ON_EN_CHANGE(IDC_EDIT_FAXNUMBER, OnEnChangeEditFaxnumber)
	ON_EN_CHANGE(IDC_EDIT_QUANTITY, OnEnChangeEditQuantity)
	ON_EN_CHANGE(IDC_EDIT_EMAIL, OnEnChangeEditEmail)
	ON_EN_CHANGE(IDC_EDIT_AGENT, OnEnChangeEditAgent)
	ON_EN_CHANGE(IDC_EDIT_RATE, OnEnChangeEditRate)
	ON_EN_CHANGE(IDC_EDIT_ZIP_FILTER, OnEnChangeEditZipFilter)
	ON_EN_CHANGE(IDC_EDIT_DESIRED_LOAN_FILTER, OnEnChangeEditDesiredLoanFilter)
	ON_EN_CHANGE(IDC_EDIT_AREACODE_FILTER, OnEnChangeEditAreacodeFilter)
	ON_BN_CLICKED(IDC_CHECK_ONEPASS, OnBnClickedCheckOnepass)
	ON_BN_CLICKED(IDC_CHECK_EXCEL, OnBnClickedCheckExcel)
	ON_BN_CLICKED(IDC_CHECK_FAX, OnBnClickedCheckFax)
	ON_BN_CLICKED(IDC_CHECK_WORD, OnBnClickedCheckWord)
	ON_BN_CLICKED(IDC_CHECK_ONELEAD, OnBnClickedCheckOnelead)
	ON_BN_CLICKED(IDC_BUTTON_OPEN_INTERNET_ORDERS, OnBnClickedButtonOpenInternetOrders)
	ON_BN_CLICKED(IDC_BUTTON_OPEN_INTERNET_ONEPASS, OnBnClickedButtonOpenInternetOnepass)
	ON_BN_CLICKED(IDC_BUTTON_PROCESS, OnBnClickedButtonProcess)
	ON_BN_CLICKED(IDC_BUTTON_STOP, OnBnClickedButtonStop)
	ON_WM_SIZE()
	ON_BN_CLICKED(IDC_BUTTON_VERIFIED, OnBnClickedButtonVerified)
	ON_BN_CLICKED(IDC_BUTTON_INTERNET, OnBnClickedButtonInternet)
	ON_BN_CLICKED(IDC_BUTTON_ONEPASS, OnBnClickedButtonOnepass)
	ON_CBN_SELCHANGE(IDC_COMBO_GROUP, OnCbnSelchangeComboGroup)
	ON_COMMAND(ID_FILE_EXIT, OnFileExit)
	ON_COMMAND(ID_HELP_ABOUT, OnHelpAbout)
	ON_BN_CLICKED(IDC_BUTTON_RETRY, OnBnClickedButtonRetry)
	ON_BN_CLICKED(IDC_BUTTON_GEN_ORDERS, OnBnClickedButtonGenOrders)
	ON_BN_CLICKED(IDC_BUTTON_STATS, OnBnClickedButtonStats)
	ON_CBN_SELCHANGE(IDC_COMBO_STATES, &CBatchProcessDlg::OnCbnSelchangeComboStates)
END_MESSAGE_MAP()


// CBatchProcessDlg message handlers

BOOL CBatchProcessDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
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

	// TODO: Add extra initialization here
	killtimer.SetParentWindow(this->m_hWnd);
	killtimer.Start();
	leadDBMap.SetParentWindow(this->m_hWnd);
	theProcessQueue.SetParentWindow(this->m_hWnd);
	m_group.SetCurSel(3);//set Normal as the default
	m_groupIndex = 3;
	m_comboStates.SetCurSel(51);
	m_comboStatesIndex = 51;
	//theLogger.Start();
	manualProcessThread.SetParentWindow(this->m_hWnd);
	manualProcessThread.Start();
	
	CString str;	
	str = CString(theApp.m_lpCmdLine);
	PrintOut(str);
	if(str.CompareNoCase("autorun")==0)
	{
		autorun = true;
		PrintOut("Got an autorun command from commandline");
		OnBnClickedButtonProcess();
		
	}
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CBatchProcessDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CBatchProcessDlg::OnPaint() 
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
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CBatchProcessDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CBatchProcessDlg::OnBnClickedButtonRead()
{
	// TODO: Add your control notification handler code here
	/*PrintOut("test");
	Lead aLead, aLead2;
	aLead.SetLead_Id(327789645);
	aLead.SetLead_DoNotUse(true);
	aLead.SetLead_FirstName("firstname");
	aLead.SetLead_LastName("lastname");
	aLead.SetLead_CoApp("co app");
	aLead.SetDupe(false);
	COleDateTime aDate;
	aDate.SetDate(2006,5,15);
	aLead.SetImportDate(aDate);
	aLead.SetLangInt(1);
	aLead.SetLead_1stBalance(100000);
	aLead.SetLead_Address("123 nowhere");
	aLead.SetLead_City("aCity");
	aDate.SetDate(2006,7,11);
	aLead.SetLead_CreateDate(aDate);
	aLead.SetLead_Credit("good");
	aLead.SetLead_CurrentValue(150000);
	aLead.SetLead_DesiredLoan(140000);
	aLead.SetLead_Email("test@test.com");
	aLead.SetLead_FixedAdjust("Fixed");
	aLead.SetLead_HomePhone("8187851234");
	aLead.SetLead_HouseType("SFR");
	aLead.SetLead_Late("not late");
	aLead.SetLead_LookingTo("refi");
	aLead.SetLead_Payment("1200");
	aLead.SetLead_Rate(5.6);
	aLead.SetLead_Salary(5000);
	aLead.SetLead_State("CA");
	aLead.SetLead_TimePeriod("na");
	aLead.SetLead_Verified(true);
	aLead.SetLead_WorkInfo("workinfo");
	aLead.SetLead_WorkPhone("8184444868");
	aLead.SetLead_YouShouldCall("morning");
	aLead.SetLead_Zip("91405");
	aLead.SetPersonalTouch("likes football");
	aDate.SetDate(2006,1,2);
	aLead.SetSys_Cre_Date(aDate);

	
	PrintOut("aLead");
	PrintOut(aLead);
	//PrintOut(aLead.GetLead_Id_String()+ ", " + aLead.GetLead_DoNotUse_String() + ", " + aLead.GetLead_FirstName() + ", " + aLead.GetLead_LastName() + ", " + aLead.GetLead_CoApp());
	aLead2 = aLead;
	PrintOut("aLead2 = aLead");
	PrintOut(aLead2);*/
	//PrintOut(aLead2.GetLead_Id_String()+ ", " + aLead2.GetLead_DoNotUse_String() + ", " + aLead2.GetLead_FirstName() + ", " + aLead2.GetLead_LastName() + ", " + aLead2.GetLead_CoApp());
	ReadLeads();
	
}

void CBatchProcessDlg::PrintOut(CString str)
{
	try
	{
		//print a line of text to the feedback listbox
		this->m_output.AddString(str);		
		if(m_output.GetCount()> 5000) 
		{			
			m_output.DeleteString(0);			
		}
		this->m_output.SetCurSel(m_output.GetCount()-1);
		
		}
	catch(...)
	{		
		WriteLog("Error: CBatchProcessDlg::PrintOut");
		AfxMessageBox("Error: CBatchProcessDlg::PrintOut");
	}
	
}

long CBatchProcessDlg::ReadLeads(void)
{
	/*long result = -1;
	HRESULT hr = S_OK;
    try
    {
		PrintOut("Init database connection");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

       _RecordsetPtr pRstInternetLeads = NULL;

      // Call Create instance to instantiate the Record set
      hr = pRstInternetLeads.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {
            //printf("Failed creating record set instance\n");
		  PrintOut("Failed creating record set instance");
		  result = -1;
            //return 0;
      }
	  else
	  {
		//WHERE (((recentLeads.dupe)=False) AND ((recentLeads.age)<14));
		//Open the Record set for getting records from InternetLeads query
		pRstInternetLeads->Open("SELECT * FROM recentLeads WHERE dupe=false AND age<7 ORDER BY lead_id", strCnn, adOpenStatic, adLockReadOnly, adCmdText);

		

		pRstInternetLeads->MoveFirst();

		//Loop through the Record set
		if (!pRstInternetLeads->EndOfFile)
		{
			result = 0;
			PrintOut("Start Lead Read from file");
			CString lead_id;
			CString first_name;
			CString last_name;
			CString sys_cre_date;
			//leads.SetOutput(&m_output);
			//Lead aLead;
			CLead* aLead;//&& (result <= 10000)
			while((!pRstInternetLeads->EndOfFile))
			{				
				aLead = new CLead();
				aLead->ReadLead(pRstInternetLeads);
				//aLead.PrintOut(&m_output);
				leads.add(aLead);
				if(aLead->GetLead_Verified())
				{
					verifiedLeads.add(aLead);
				}
				if((result % 1000)==0)
				{
					//aLead.PrintOut(&m_output);
					PrintOut(GetLong_String(result) + " Leads Read");
					
				}
				
				result++;				
				pRstInternetLeads->MoveNext();
			}			
			
			PrintOut(GetLong_String(result) + " Leads Read");
		}
	  }
	  
	}
	catch(_com_error & ce)
	{
		CString str;
		//_bstr_t temp;
		//temp = ce.Description;
		//str = getCString(temp);
		str = "Error: " + GetCString(ce.Description());
		PrintOut(str);
		result = -1;
		//printf("Error:%s\n",ce.Description);
	}
	
  CoUninitialize();
	return result;*/
	return 0;
}

CString CBatchProcessDlg::GetCString(_bstr_t bstrStart)
{
	//CStringHelper aHelper;
	return CStringHelper::GetCString(bstrStart);
	//return aHelper.GetCString(bstrStart);
	/*
	try
	{

		TCHAR szFinal[255];
		_stprintf(szFinal, _T("%s"), (LPCTSTR)bstrStart);
		CString str;
		str = szFinal;
		return str;
		
	}
	catch(...)
	{
		//error trap template
		CString err;
		err = "Error: CBatchProcessDlg::GetCString(_bstr_t bstrStart)";
		PrintOut(err);
		WriteLog(err);
	}
	*/
}

void CBatchProcessDlg::OnEnterIdle(UINT nWhy, CWnd* pWho)
{
	CDialog::OnEnterIdle(nWhy, pWho);

	// TODO: Add your message handler code here
}

CString CBatchProcessDlg::GetLong_String(long num)
{
	
	try
	{
		char buffer[200];
		_ltoa( num, buffer, 10 );
		CString str;			
		str = buffer;
		return str;
	}
	catch(...)
	{
		//error trap template
		CString err;
		err = "Error: CBatchProcessDlg::GetLong_String(long num)";
		PrintOut(err);
		WriteLog(err);
	}
}
CString CBatchProcessDlg::GetInt_String(int num)
{
	try
	{
		char buffer[200];
		_itoa( num, buffer, 10 );
		CString str;			
		str = buffer;
		return str;
	}
	catch(...)
	{
		//error trap template
		CString err;
		err = "Error: CBatchProcessDlg::GetInt_String(int num)";
		PrintOut(err);
		WriteLog(err);
		return CString();
	}
}
CString CBatchProcessDlg::GetCString(_variant_t var)
{
	return CString();
}

CString CBatchProcessDlg::GetField(_RecordsetPtr rs, CString field)
{
	_bstr_t val;
	_variant_t vtFld;
	
	try
	{
		vtFld = rs->Fields->GetItem(field.GetBuffer())->Value;
		//if(vtFld.vt != VT_NULL && vtFld.vt != VT_EMPTY)
		if(vtFld.vt == VT_NULL)
		{
			return CString();
		}
		else if(vtFld.vt == VT_BSTR)
		{	
			val = vtFld.bstrVal;
			return GetCString(val);
			
		}
		if (vtFld.vt == VT_DATE)
		{
			COleDateTime dt(vtFld);
			CString aDateStr;
			aDateStr = GetInt_String(dt.GetMonth()) + "/";
			aDateStr = aDateStr + GetInt_String(dt.GetDay()) + "/";
			aDateStr = aDateStr + GetInt_String(dt.GetYear());
			return aDateStr;
		}
		
		return "type unkown";
	}
	catch(...)
	{
		return "";
	}
}



CString CBatchProcessDlg::GetField(_RecordsetPtr rs, CString field, int vt_type)
{
	CString result;
	if(vt_type == VT_I4)
	{
		long val;
		val = rs->Fields->GetItem(field.GetBuffer())->Value.lVal;
		result = GetLong_String(val);
	}
	return result;
}

void CBatchProcessDlg::PrintOut(CLead aLead)
{
	aLead.PrintOut(&m_output);
}

void CBatchProcessDlg::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	//int nWidth=(pRect->right)-(pRect->left);
	//int nHeight=(pRect->bottom)-(pRect->top);	
	//m_output.MoveWindow(11,290,nWidth - 29,nHeight-330,true);
	// TODO: Add your message handler code here

}

void CBatchProcessDlg::OnBnClickedButtonPrintList()
{
	// TODO: Add your control notification handler code here
	/*LeadList aList;
	CString states;
	m_states.GetWindowText(states);
	if(states.GetLength()>0)
	{
		aList = leads.GetLeads(m_verified.GetCheck(), states);
	}
	else
	{
		aList = leads.GetLeads(m_verified.GetCheck());
	}
	long count = aList.getCount(); 
	for(long x = 0;x<count;x++)
	{
		((Lead*)aList.getAt(x))->PrintOut(&m_output);
	}*/
	//long count = leads.getCount();
	//for(long x = 0;x<count;x++)
	//{
		//leads.getAt(x)->PrintOut(&m_output);
		//leads.getAt(x)->SetPersonalTouch("Test Modify");
		
	//}
	
	//count = verifiedLeads.getCount();
	//for(long x = 0;x<count;x++)
	//{
		//verifiedLeads.getAt(x)->PrintOut(&m_output);
		//verifiedLeads.getAt(x)->SetPersonalTouch("Test Modify");
		
	//}

}

void CBatchProcessDlg::OnBnClickedButtonPrintInternet()
{
	// TODO: Add your control notification handler code here
	/*LeadList aList;
	aList = leads.GetLeads(false);
	long count = aList.getCount(); 
	for(long x = 0;x<count;x++)
	{
		((Lead*)aList.getAt(x))->PrintOut(&m_output);
	}*/
}
LRESULT CBatchProcessDlg::OnReadLead(WPARAM wp, LPARAM lp)
{
	// TODO: Add your control notification handler code here
	//StateList states;
	
	//states.PrintOut(&m_output);
	//leads.PrintOut(&m_output);
	char buffer[200];
   
    _itoa( wp, buffer, 10 );

	CString str;
	str = "Read Lead: ";
	str = str + buffer;
	PrintOut(str);

	return 0;
}
void CBatchProcessDlg::OnBnClickedButtonTest()
{
	// TODO: Add your control notification handler code here
	CGenOrders aGenOrders;
	CString sql;
	sql = aGenOrders.GetSQL();
	m_sql.SetWindowText(sql);
	
	
	
	
}

LRESULT CBatchProcessDlg::OnInitDB(WPARAM wp, LPARAM lp)
{
	PrintOut("Initialize Database Connection");
	return 0;
}
LRESULT CBatchProcessDlg::OnFailCreateRecordInstance(WPARAM wp, LPARAM lp)
{
	PrintOut("Failed to create record instance");
	return 0;
}
LRESULT CBatchProcessDlg::OnComError(WPARAM wp, LPARAM lp)
{
	PrintOut("COM Error");
	return 0;
}
LRESULT CBatchProcessDlg::OnLeadsCount(WPARAM wp, LPARAM lp)
{
	char buffer[200];   
    _itoa( wp, buffer, 10 );
	CString str;
	str = buffer;
	str = str + " Leads Read";
	PrintOut(str);
	return 0;
}
LRESULT CBatchProcessDlg::OnStartLeadFileRead(WPARAM wp, LPARAM lp)
{
	//PrintOut("Start reading leads file");
	return 0;
}


void CBatchProcessDlg::OnDestroy()
{
	//DWORD dwExitCode;
	//processOrders.Stop(dwExitCode, 0);
	manualProcessThread.Kill();
	CDialog::OnDestroy();

	// TODO: Add your message handler code here
}

void CBatchProcessDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	OnCancel();
}

void CBatchProcessDlg::OnCancel()
{
	// TODO: Add your specialized code here and/or call the base class
	
	//DWORD dwExitCode;

	//processOrders.Stop(dwExitCode, 0);
	 
	//if (processOrders.GetActivityStatus() != CProcessOrders::THREAD_MSG_SESSION_CLOSED) 
	//{ 
		//	call this handler again
		// IDCANCEL is the dialog's 'Cancel' button resource ID
		//PostMessage(WM_COMMAND, MAKEWPARAM((WORD)IDCANCEL, (WORD)IDCANCEL));
	//} 
	//else 
	//{ 
		//processOrders.Stop(dwExitCode);

		// ... and close the dialog properly 
		CDialog::OnCancel();
	//}

}

void CBatchProcessDlg::OnBnClickedButtonTestOrder()
{
	// TODO: Add your control notification handler code here
	
	COrder aOrder;
	bool stop = false;
	aOrder.SetStop(&stop);
	//aOrder = testOrder;
	aOrder.SetAgent("Test Name");
	//aOrder.Set2ndMortgage(1);
	//aOrder.SetAdCampaign(true);
	aOrder.SetAgentID(-25);
	//aOrder.SetAreacodeFilter("818");
	//aOrder.SetCashout(2);
	aOrder.SetCompanyID(-20);
	aOrder.SetCompanyName("TestCompany");
	aOrder.SetContact("Test contact");
	//aOrder.SetCreditScore(3);
	//aOrder.SetDebtCon(2);
	//aOrder.SetDesiredLoanFilter(100000);
	aOrder.SetEmail("test@test.com");
	aOrder.SetFaxNumber("8181111111");
	//aOrder.SetHouseType(1);
	//aOrder.SetLang(2);
	aOrder.SetLocation(2);
	//aOrder.SetLookingTo(1);
	//aOrder.SetOnePass(true);
	//aOrder.SetOneLeadPerPage(true);
	aOrder.SetOrderID(-13);
	aOrder.SetPrint(true);
	
	aOrder.SetPriority(true);
	aOrder.SetQuantity(5);
	//aOrder.SetRateFilter(3.5);
	//aOrder.SetSendEmailNote(true);
	//aOrder.SetSendExcel(true);
	//aOrder.SetSendFax(true);
	//aOrder.SetSendMini1003(true);
	aOrder.SetSendWord(true);
	aOrder.SetStates("CA");
	//aOrder.SetSubprime(true);
	aOrder.SetVerified(false);
	aOrder.SetOneLeadPerPage(false);
	aOrder.SetOnePass(true);
	aOrder.SetSendMini1003(false);
	//aOrder.SetZipFilter("91405");
	aOrder.SetMaxNeeded(8);
	LoadScreenFromOrder(&aOrder);
	PrintOut(aOrder.ToString());
	aOrder.SetParentWindow(this->m_hWnd);
	CLeadDB leadDBMap;
	aOrder.SetLeadDBMap(&leadDBMap);
	COrder aOrder2, aOrder3;
	aOrder2 = aOrder;
	aOrder3 = aOrder;
	aOrder2.SetCompanyID(-19);
	aOrder3.SetCompanyID(-18);
	aOrder3.SetQuantity(10);
	//aOrder.AddFriend(-19);
	//aOrder.AddFriend(-18);
	//aOrder2.AddFriend(-20);
	//aOrder2.AddFriend(-18);
	//aOrder3.AddFriend(-20);
	//aOrder3.AddFriend(-19);
	aOrder.ReadLeads();
	//aOrder2.ReadLeads();
	//aOrder3.ReadLeads();

	for(int x = 0; x < 10; x++)
	{
		aOrder.AssignLead();
		//aOrder2.AssignLead();
		//aOrder3.AssignLead();
	}
	PrintOut("order 1");
	aOrder.PrintAssignedLeads();
	//PrintOut("order 2");
	//aOrder2.PrintAssignedLeads();
	//PrintOut("order 3");
	//aOrder3.PrintAssignedLeads();
	//PrintOut("Matching ordered as read");
	//aOrder.PrintMatchingLeads();
	//aOrder.SortMatching();
	//PrintOut("Matching ordered after sort");
	//aOrder.PrintMatchingLeads();
	//aOrder.AssignLead();
	//aOrder.PrintAssignedLeads();
	//COrder aOrder2;
	//aOrder2 = aOrder;
	//aOrder2.PrintAssignedLeads();
	//CString target, rtfTemplate;
	//if( _mkdir( ".\\ProcessedOrders" ) == 0 )
	//{
		//PrintOut(".\\ProcessedOrders directory created");
	//}
	//else
	//{
		//PrintOut(".\\ProcessedOrders directory already exists");
	//}
	//target = ".\\ProcessedOrders\\test.rtf";
	
	//if(aOrder.GetVerified() || aOrder.GetOneLeadPerPage())
	//{
		//rtfTemplate = "1perPage.rtf";
	//}
	//else
	//{
		//rtfTemplate = "2perPage.rtf";
	//}
	CSettings mySettings;
	mySettings.ReadSettings();
	mySettings.ReadLocations();
	CLocation myloc = mySettings.GetLocation(aOrder.GetLocationLong());
	
	//CRTFWriter aWriter(myloc.GetDesDir(), myloc.GetTemplateDir(), this->m_hWnd); 
	//aWriter.Write(aOrder);
	CCSVWriter aWriter(myloc.GetDesDir(),this->m_hWnd);
	aWriter.Write(aOrder, myloc);
}

void CBatchProcessDlg::OnBnClickedCheckVerified()
{
	// TODO: Add your control notification handler code here
	testOrder.SetVerified((bool)m_verified.GetCheck());
}

void CBatchProcessDlg::OnEnChangeEditStates()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	CString str;
	m_states.GetWindowText(str);
	testOrder.SetStates(str);
}

void CBatchProcessDlg::OnEnChangeEditCompanyid()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	//ltoa
	CString sCompid;
	m_companyid.GetWindowText(sCompid);
	testOrder.SetCompanyID(atol(sCompid.GetBuffer()));

}

void CBatchProcessDlg::OnEnChangeEditContact()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	CString str;
	m_contact.GetWindowText(str);
	testOrder.SetContact(str);
}

void CBatchProcessDlg::OnEnChangeEditOrderid()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	CString str;
	m_orderid.GetWindowText(str);
	testOrder.SetOrderID(atol(str.GetBuffer()));
}

void CBatchProcessDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	OnOK();
}

void CBatchProcessDlg::OnEnChangeEditFaxnumber()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	CString str;
	m_faxnumber.GetWindowText(str);
	testOrder.SetFaxNumber(str);
}

void CBatchProcessDlg::OnEnChangeEditQuantity()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	CString str;
	m_quantity.GetWindowText(str);
	testOrder.SetQuantity(atoi(str.GetBuffer()));
}

void CBatchProcessDlg::OnEnChangeEditEmail()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	testOrder.SetEmail(GetText(&m_email));
}

CString CBatchProcessDlg::GetText(CEdit *aBox)
{
	CString str;
	aBox->GetWindowText(str);
	return str;
}

void CBatchProcessDlg::OnEnChangeEditAgent()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	testOrder.SetAgent(GetText(&m_agent));
}

void CBatchProcessDlg::OnEnChangeEditRate()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	testOrder.SetRateFilter(atof(GetText(&m_rate).GetBuffer()));
}

void CBatchProcessDlg::OnEnChangeEditZipFilter()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	testOrder.SetZipFilter(GetText(&m_zip));
}

void CBatchProcessDlg::OnEnChangeEditDesiredLoanFilter()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	testOrder.SetDesiredLoanFilter(atol(GetText(&m_loan)));
}

void CBatchProcessDlg::OnEnChangeEditAreacodeFilter()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	testOrder.SetAreacodeFilter(GetText(&m_areacodes));
}

void CBatchProcessDlg::OnBnClickedCheckOnepass()
{
	// TODO: Add your control notification handler code here
	testOrder.SetOnePass(m_onepass.GetCheck());
}

void CBatchProcessDlg::OnBnClickedCheckExcel()
{
	// TODO: Add your control notification handler code here
	testOrder.SetSendExcel(m_excel.GetCheck());
}

void CBatchProcessDlg::OnBnClickedCheckFax()
{
	// TODO: Add your control notification handler code here
	testOrder.SetSendFax(m_fax.GetCheck());
}

void CBatchProcessDlg::OnBnClickedCheckWord()
{
	// TODO: Add your control notification handler code here
	testOrder.SetSendWord(m_word.GetCheck());
}

void CBatchProcessDlg::OnBnClickedCheckOnelead()
{
	// TODO: Add your control notification handler code here
	testOrder.SetOneLeadPerPage(m_onelead.GetCheck());
}

void CBatchProcessDlg::LoadScreenFromOrder(COrder *aOrder)
{
	this->m_agent.SetWindowText(aOrder->GetAgent());
	this->m_areacodes.SetWindowText(aOrder->GetAreacodeFilter());
	this->m_companyid.SetWindowText(aOrder->GetCompanyID());
	this->m_contact.SetWindowText(aOrder->GetContact());
	this->m_email.SetWindowText(aOrder->GetEmail());
	this->m_excel.SetCheck(aOrder->GetExcel());
	this->m_fax.SetCheck(aOrder->GetSendFax());
	this->m_faxnumber.SetWindowText(aOrder->GetFaxNumber());
	this->m_loan.SetWindowText(aOrder->GetDesiredLoanFilter());
	this->m_onelead.SetCheck(aOrder->GetOneLeadPerPage());
	this->m_onepass.SetCheck(aOrder->GetOnePass());
	this->m_orderid.SetWindowText(aOrder->GetOrderID());
	this->m_quantity.SetWindowText(aOrder->GetQuantity());
	this->m_rate.SetWindowText(aOrder->GetRateFilter());
	this->m_states.SetWindowText(aOrder->GetStates());
	this->m_verified.SetCheck(aOrder->GetVerified());
	this->m_word.SetCheck(aOrder->GetSendWord());
	this->m_zip.SetWindowText(aOrder->GetZipFilter());
	this->m_cashout.SetWindowText(aOrder->GetCashout());
	this->m_2ndmortgage.SetWindowText(aOrder->Get2ndMortgage());
	this->m_debtcon.SetWindowText(aOrder->GetDebtCon());
	this->m_emailnote.SetCheck(aOrder->GetSendEmailNote());
	this->m_print.SetCheck(aOrder->GetPrint());
	this->m_lookingto.SetWindowText(aOrder->GetLookingto());
	this->m_lang.SetWindowText(aOrder->GetLang());
	this->m_creditscore.SetWindowText(aOrder->GetCreditscore());
	this->m_hometype.SetWindowText(aOrder->GetHomeType());
	this->m_location.SetWindowText(aOrder->GetLocation());
	this->m_priority.SetCheck(aOrder->GetPriority());
	this->m_mini1003.SetCheck(aOrder->GetSendMini1003());
	this->m_adcampaign.SetCheck(aOrder->GetAdCampaign());
	this->m_agentid.SetWindowText(aOrder->GetAgentID());
	this->m_subprime.SetCheck(aOrder->GetSubPrime());
	this->m_sql.SetWindowText(aOrder->GetSQL());
	this->m_SFR.SetWindowText(aOrder->GetSFR());
	this->m_Mobile.SetWindowText(aOrder->GetMobile());
	this->m_Commercial.SetWindowText(aOrder->GetCommercial());
	this->m_city.SetWindowText(aOrder->GetCity());
	this->m_ltv.SetWindowText(aOrder->GetLTV());
	this->m_ltvAbove.SetCheck(aOrder->GetLTVAbove());
	
	
	
}

void CBatchProcessDlg::OnBnClickedButtonOpenInternetOrders()
{
	// TODO: Add your control notification handler code here
	HRESULT hr = S_OK;
    try
    {
		//PrintOut("Init database connection");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

      

      // Call Create instance to instantiate the Record set
      hr = pRstInternetOrders.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {
            //printf("Failed creating record set instance\n");
		  PrintOut("Failed creating record set instance");
		 
      }
	  else
	  {
		
		pRstInternetOrders->Open("SELECT * FROM t_orders WHERE (Ord_Complete=false) AND (Ord_OnePass=false) AND (Ord_Verified=false) ORDER BY Ord_Id", strCnn, adOpenStatic, adLockReadOnly, adCmdText);

		if(!pRstInternetOrders->EndOfFile) pRstInternetOrders->MoveFirst(); else PrintOut("No Open Orders");
		//while
		if(!pRstInternetOrders->EndOfFile)
		{	
			
			//PrintOut("Start order Read from file");			
			COrder aOrder;
			aOrder.SetParentWindow(this->m_hWnd);
			aOrder.SetLeadDBMap(&leadDBMap);
			//aOrder.leadsList = &leadsMasterList;
			aOrder.ReadOrder(pRstInternetOrders);
			LoadScreenFromOrder(&aOrder);
			this->RedrawWindow();
			aOrder.ReadLeads();
			//aOrder.PrintMatchingLeads();
			pRstInternetOrders->MoveNext();
		}
		
		//if (!pRstInternetOrders->EndOfFile)
		//{			
			//PrintOut("Start 2nd order Read from file");			
			//COrder aOrder;
			//aOrder.SetParentWindow(this->m_hWnd);
			//aOrder.SetLeadDBMap(&leadDBMap);
			
			//aOrder.ReadOrder(pRstInternetOrders);
			//LoadScreenFromOrder(&aOrder);
			//this->RedrawWindow();
			//aOrder.ReadLeads();
			//aOrder.PrintMyLeads();
		//}
	  }
	  PrintOut("Master List contains");
		leadDBMap.PrintLeads();
	  //for(int x=0; x < leadsMasterList.GetCount(); x++)
	  //{
			//PrintOut(((CLead)leadsMasterList.GetAt(x)).ToString());
	  //}
		CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;
		//_bstr_t temp;
		//temp = ce.Description;
		//str = getCString(temp);
		str = "Error: " + GetCString(ce.Description());
		PrintOut(str);
		
		//printf("Error:%s\n",ce.Description);
	}
	
  
}

LRESULT CBatchProcessDlg::OnStartMessageString(WPARAM wp, LPARAM lp)
{
	//theLogger.PostCommand(N_START_STRING);
	message = "";
	return LRESULT();
}

LRESULT CBatchProcessDlg::OnEndMessageString(WPARAM wp, LPARAM lp)
{
	//theLogger.PostCommand(N_STOP_STRING);
	PrintOut(message);
	message = "";
	return LRESULT();
}

LRESULT CBatchProcessDlg::OnCharMessageString(WPARAM wp, LPARAM lp)
{
	//theLogger.PostCommand(N_MESSAGE + wp);
	char c = (char) wp;
	//char  buffer[200];
	//sprintf( buffer,     "Got letter %c ", wp );
	//PrintOut(buffer);
	message = message + c;
	return LRESULT();
}

void CBatchProcessDlg::OnBnClickedButtonOpenInternetOnepass()
{
	// TODO: Add your control notification handler code here
	HRESULT hr = S_OK;
    try
    {
		//PrintOut("Init database connection");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

      

      // Call Create instance to instantiate the Record set
      hr = pRstInternetOrders.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {
            //printf("Failed creating record set instance\n");
		  PrintOut("Failed creating record set instance");
		 
      }
	  else
	  {
		
		pRstInternetOrders->Open("SELECT * FROM t_orders WHERE (Ord_Complete=false) AND (Ord_OnePass=true) AND (Ord_Verified=false) ORDER BY Ord_Id", strCnn, adOpenStatic, adLockReadOnly, adCmdText);

		if(!pRstInternetOrders->EndOfFile) pRstInternetOrders->MoveFirst(); else PrintOut("No Open Orders");
		//while
		 if(!pRstInternetOrders->EndOfFile)
		{			
			//PrintOut("Start order Read from file");			
			COrder aOrder;
			aOrder.SetParentWindow(this->m_hWnd);
			//aOrder.leadsList = &leadsMasterList;
			aOrder.SetLeadDBMap(&leadDBMap);
			aOrder.ReadOrder(pRstInternetOrders);
			LoadScreenFromOrder(&aOrder);
			this->RedrawWindow();
			aOrder.ReadLeads();
			//aOrder.PrintMatchingLeads();
			pRstInternetOrders->MoveNext();
		}
		
		//if (!pRstInternetOrders->EndOfFile)
		//{			
			//PrintOut("Start order Read from file");			
			//COrder aOrder;
			//aOrder.SetParentWindow(this->m_hWnd);
			
			//aOrder.SetLeadDBMap(&leadDBMap);
			//aOrder.ReadOrder(pRstInternetOrders);
			//LoadScreenFromOrder(&aOrder);
			//this->RedrawWindow();
			//aOrder.ReadLeads();
			//aOrder.PrintMyLeads();
		//}
	  }
	  //PrintOut("Master list contains");
	  //leadDBMap.PrintLeads();
	  CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;
		//_bstr_t temp;
		//temp = ce.Description;
		//str = getCString(temp);
		str = "Error: " + GetCString(ce.Description());
		PrintOut(str);
		
		//printf("Error:%s\n",ce.Description);
	}
	
  
}

void CBatchProcessDlg::OnBnClickedButtonProcess()
{
	// TODO: Add your control notification handler code here
	//PrintOut("Not coded yet. Use Manual Process Buttons");
	ProcessButtonsEnable(false);
	theProcessQueue.ReadJobs();
	if(theProcessQueue.GetCount() > 0)
	{
		processing = true;
		CProcessJob aJob;
		aJob = theProcessQueue.GetNextJob();
		if(aJob.GetiCategory() == PO_RETRY_SENDING)
		{
			manualProcessThread.PostCommand(PO_RETRY_SENDING);
		}
		else if(aJob.GetiCategory() == PO_STATS)
		{
			manualProcessThread.PostCommand(PO_STATS);
		}
		else if(aJob.GetiCategory() == PO_GEN_ORDERS)
		{
			manualProcessThread.PostCommand(PO_GEN_ORDERS);
		}
		else
		{
			manualProcessThread.PostCommand(aJob.GetiCategory());
			manualProcessThread.PostCommand(aJob.GetGroup());
			
			manualProcessThread.PostCommand(aJob.GetState());
			manualProcessThread.PostCommand(CMD_PROCESS);
		}		
		
	}
}

void CBatchProcessDlg::OnBnClickedButtonStop()
{
	// TODO: Add your control notification handler code here
	
	//processOrdersThread.SetStop(true);
	//manualProcessThread.SetStop(true);
	manualProcessThread.Kill();
	PrintOut("Process Thread Killed. You will need to restart the application before you can process any more Orders");
	
}

LRESULT CBatchProcessDlg::OnOrderTrans(WPARAM wp, LPARAM lp)
{
	
	//PrintOut("Got an Order trans message");
	switch ( wp )
      {
		case IDC_EDIT_COMPANYID:
			SetEditBox(&m_companyid, lp);			            
            break;
         
		case IDC_EDIT_STATES:
            SetEditBox(&m_states, lp);
            break;

		case IDC_EDIT_CONTACT:
            SetEditBox(&m_contact, lp);
            break;

		case IDC_EDIT_ORDERID:
            SetEditBox(&m_orderid, lp);
            break;

		case IDC_EDIT_FAXNUMBER:
            SetEditBox(&m_faxnumber, lp);
            break;
		
		case IDC_EDIT_QUANTITY:
            SetEditBox(&m_quantity, lp);
            break;
		
		case IDC_EDIT_EMAIL:
            SetEditBox(&m_email, lp);
            break;

		case IDC_EDIT_AGENT:
            SetEditBox(&m_agent, lp);
            break;

		case IDC_EDIT_RATE:
            SetEditBox(&m_rate, lp);
            break;

		case IDC_EDIT_ZIP_FILTER:
            SetEditBox(&m_zip, lp);
            break;

		case IDC_EDIT_DESIRED_LOAN_FILTER:
            SetEditBox(&m_loan, lp);
            break;

		case IDC_EDIT_AREACODE_FILTER:
            SetEditBox(&m_areacodes, lp);
            break;

		case IDC_EDIT_CASHOUT:
            SetEditBox(&this->m_cashout, lp);
            break;

		case IDC_EDIT_2NDMORTGAGE:
            SetEditBox(&this->m_2ndmortgage, lp);
            break;

		case IDC_EDIT_DEBTCON:
            SetEditBox(&this->m_debtcon, lp);
            break;

		case IDC_EDIT_AGENTID:
            SetEditBox(&this->m_agentid, lp);
            break;

		case IDC_EDIT_LOOKINGTO:
            SetEditBox(&this->m_lookingto, lp);
            break;

		case IDC_EDIT_LANG:
            SetEditBox(&this->m_lang, lp);
            break;

		case IDC_EDIT_CREDIT:
            SetEditBox(&this->m_creditscore, lp);
            break;

		case IDC_EDIT_HOMETYPE:
            SetEditBox(&this->m_hometype, lp);
            break;

		case IDC_EDIT_LOCATION:
            SetEditBox(&this->m_location, lp);
            break;

		case IDC_EDIT_SFR:
            SetEditBox(&this->m_SFR, lp);
            break;

		case IDC_EDIT_MOBILE:
            SetEditBox(&this->m_Mobile, lp);
            break;

		case IDC_EDIT_COMMERCIAL:
            SetEditBox(&this->m_Commercial, lp);
            break;

		case IDC_EDIT_CITY:
            SetEditBox(&this->m_city, lp);
            break;

		case IDC_EDIT_LTV:
            SetEditBox(&this->m_ltv, lp);
            break;

		case IDC_EDIT_SQL:
			SetSQL(lp);            
            break;

		case IDC_CHECK_VERIFIED:            
			SetCheckBox(&this->m_verified, lp);
            break;

		case IDC_CHECK_ONEPASS:            
			SetCheckBox(&this->m_onepass, lp);
            break;

		case IDC_CHECK_WORD:            
			SetCheckBox(&this->m_word, lp);
            break;

		case IDC_CHECK_EXCEL:            
			SetCheckBox(&this->m_excel, lp);
            break;

		case IDC_CHECK_FAX:            
			SetCheckBox(&this->m_fax, lp);
            break;

		case IDC_CHECK_ONELEAD:            
				SetCheckBox(&this->m_onelead, lp);
				break;

		case IDC_CHECK_EMAILNOTE:            
				SetCheckBox(&this->m_emailnote, lp);
				break;

		case IDC_CHECK_PRINT:            
				SetCheckBox(&this->m_print, lp);
				break;

		case IDC_CHECK_PRIORITY:            
				SetCheckBox(&this->m_priority, lp);
				break;

		case IDC_CHECK_MINI1003:            
				SetCheckBox(&this->m_mini1003, lp);
				break;

		case IDC_CHECK_ADCAMPAIGN:            
				SetCheckBox(&this->m_adcampaign, lp);
				break;

		case IDC_CHECK_SUBPRIME:            
				SetCheckBox(&this->m_subprime, lp);
				break;

		case IDC_CHECK_ABOVE:            
				SetCheckBox(&this->m_ltvAbove, lp);
				break;

		case IDC_EDIT_LAST_ORDER_QUANTITY:
				SetEditBox(&this->m_lastOrderQuantity, lp);
				break;

		case IDC_EDIT_COMPANY_ORDERS_COUNT:
				SetEditBox(&this->m_companyTotalOrders, lp);
				break;

		case IDC_EDIT_YIELD:
				SetEditBox(&this->m_yield, lp);
				break;

		case IDC_EDIT_CAMPAIGNID:
				SetEditBox(&this->m_campaignid, lp);
				break;

         default:
           break;
      }

	return LRESULT();
}

void CBatchProcessDlg::SetEditBox(CEdit *pEdit, int iMessage)
{	
	
	if(iMessage == 0)
	{
		pEdit->SetWindowText("");
	}
	else
	{
		CString str;
		
		pEdit->GetWindowText(str);
		str = str + (char)iMessage;
		pEdit->SetWindowText(str);
	}			
}

void CBatchProcessDlg::SetCheckBox(CButton* pCheck, int state)
{
	pCheck->SetCheck(state);
}

void CBatchProcessDlg::SetSQL(int iMessage)
{
	switch (iMessage)
	{
		case 0:
			sSQL = "";
			break;

		case -1:
			this->m_sql.SetWindowText(sSQL);
			break;

		default:
			sSQL = sSQL + (char)iMessage;
			break;
	}
	
}
LRESULT CBatchProcessDlg::OnStartFaxString(WPARAM wp, LPARAM lp)
{
	faxmessage = "";
	//PrintOut("start string message");
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnCharFaxString(WPARAM wp, LPARAM lp)
{
	//PrintOut("got a letter");
	char c = (char) wp;
	//char  buffer[200];
	//sprintf( buffer,     "Got letter %c ", wp );
	//PrintOut(buffer);
	faxmessage = faxmessage + c;
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnFaxSuccess(WPARAM wp, LPARAM lp)
{
	ResetTimer();
	sendingFax = false;
	iFailCount = 0;
	CFaxJob aJob = (CFaxJob) myfaxJobs.GetAt(0);
	//MarkOrderComplete(aJob); //marked complete inside of fax thread now
	myfaxJobs.RemoveAt(0);
	SendNextFax();
	return LRESULT();
}
//LRESULT CBatchProcessDlg::OnLogWriteDone(WPARAM wp, LPARAM lp)
//{	
	//writingLog = false;	
	//myLogEntrys.RemoveAt(0);
	//WriteNextEntry();
	//return LRESULT();
//}
LRESULT CBatchProcessDlg::OnFaxFail(WPARAM wp, LPARAM lp)
{
	sendingFax = false;
	iFailCount++;
	if(iFailCount <=5)
	{
		PrintOut("Fax Fail trying again");
		ResetTimer();
		SendNextFax();//try and send again
	}
	else
	{
		PrintOut("The fax sender has failed 5 times. Halting fax sender");
		CFaxJob aJob = (CFaxJob) myfaxJobs.GetAt(0);
		//AddToFailLog(aJob.GetDiv(), "Send Fax Fail:" + aJob.GetIDMessage());
		PrintFail("Send Fax Time Out:" + aJob.GetIDMessage());
		ResetTimer();
		sendingFax = false;
		iFailCount = 0;
		myfaxJobs.RemoveAt(0);
		SendNextFax();
	}
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnEndFaxString(WPARAM wp, LPARAM lp)
{
	//PrintOut("end string message");
	PrintOut(faxmessage);
	faxmessage = "";
	return LRESULT();
}
bool CBatchProcessDlg::SendFax(CString faxnum, CString agentname, CString agentEmail, CString company, CString contact, int div, CString filename)
{
	CFax aFax;
	aFax.SetParentWindow(this->m_hWnd);
	CSettings mySettings;
	mySettings.ReadSettings();
	mySettings.ReadLocations();
	
	CLocation myLoc;
	myLoc = mySettings.GetLocation(div);
	CString body;
	CString defaultReturnEmail;
	defaultReturnEmail = myLoc.GetReturnEmail();
	body = myLoc.GetBody();
	body.Replace("<agentname>", agentname);	
	body.Replace("<agentemail>", agentEmail);
	//body.Replace("<orderid>", orderid);
	body.Replace("<company>", company);
	body.Replace("<contact>", contact);	
	
   //PrintOut("Write temp fax body file");
   //ofstream fileout("c:\\~faxbody.txt");
   //streamsize y = body.GetLength();
   //fileout.write( body.GetBuffer(), y );
   //fileout.close();
	aFax.SetCoverPage("z:\\COVERPAGE.COV");
	aFax.SetSendCoverpage(1);
	PrintOut(body);
	aFax.SetCoverPageNote(body);
	aFax.SetBody(filename);
	if(testing)
	{		
		aFax.SetFaxNumber("18184444768");
		
	}
	else
	{
		aFax.SetFaxNumber(faxnum);
	}
	
	PrintOut("Attempting to send fax to: " + aFax.GetFaxNumber());
	return aFax.Send();
}

bool CBatchProcessDlg::SendMail(CString brokerEmail, CString agentname, CString agentEmail, CString company, CString contact, int div, CString orderid)
{
	CLocation myLoc;
	CSettings mySettings;
	mySettings.ReadSettings();
	mySettings.ReadLocations();
	myLoc = mySettings.GetLocation(div);
	CString body;
	CString defaultReturnEmail;
	defaultReturnEmail = myLoc.GetReturnEmail();
	body = myLoc.GetBody();
	body.Replace("<agentname>", agentname);	
	body.Replace("<agentemail>", agentEmail);
	body.Replace("<orderid>", orderid);
	body.Replace("<company>", company);
	body.Replace("<contact>", contact);

	bool success = false;
	CString to;
	CString agent;
	CString subject;	

	subject = myLoc.GetSubject();
	subject.Replace("<agentname>", agentname);	
	subject.Replace("<agentemail>", agentEmail);
	subject.Replace("<orderid>", orderid);
	subject.Replace("<company>", company);
	subject.Replace("<contact>", contact);

	if(testing)
	{
		to = "test@country-loan.com";
		agent = "test@country-loan.com";
		
	}
	else
	{
		to = brokerEmail;
		if(to == "")		
		to = "kenco_computers@yahoo.com";
		agent = agentEmail;
	}	
	


	ESMTP mail;
	mail.SetParentWindow(this->m_hWnd);
	//mail.SetName("Testing");
	mail.SetName(agentname);
	//mail.SetAddress("test@country-loan.com");
	if(agentEmail.GetLength()>0)
	{
		mail.SetAddress(agentEmail);
	}
	else
	{
		mail.SetAddress(defaultReturnEmail);
	}

	
	mail.SetTo(to);
	//mail.SetBCC("test@country-loan.com");
	mail.SetSubject(subject);
	
	//mail.m_sHost = "216.219.253.246";
	//mail.SetUserName("test@country-loan.com");
	//mail.SetPassword("test123");
	//PrintOut(mySettings.GetSmtp());
	mail.SetHost(mySettings.GetSmtp());
	mail.SetUserName(mySettings.GetUsername());
	mail.SetPassword(mySettings.GetPassword());
	mail.SetBody(body);
	CString filename = ".\\ProcessedOrders\\";
	filename = filename + orderid;
	filename = filename + ".rtf";
	mail.AddFilename(filename);

	PrintOut("Sending email");
	if(mail.Send())
	{
		PrintOut("Email sent to: " + contact + " At: " + company + " Email: " + mail.GetTo() + " and BCC to: " + mail.GetBCC());
		PrintOut(" ");
		success = true;
		//AfxMessageBox("Message send to: " + mail.m_sTo , MB_OK);
	}
	else
	{
		PrintOut("Error sending Email. Make sure you are connected to the internet.");
		//AfxMessageBox("Error sending Email. Make sure you are connected to the internet before hitting Send." , MB_OK);
		success = false;
	}
	return success;
}
void CBatchProcessDlg::SendNextFax(void)
{
	if((myfaxJobs.GetCount()>0) && (!sendingFax))
	{
		CFaxJob aJob = (CFaxJob) myfaxJobs.GetAt(0);
		if(pt_faxer)
		{
			delete pt_faxer;
		}
		pt_faxer = new CFaxThread();
		pt_faxer->SetParentWindow(this->m_hWnd);		
		pt_faxer->SetFaxJob(aJob);
		sendingFax = true;
		ResetTimer();
		pt_faxer->Start();		
	}
	else
	{
		CheckForDone();
	}
}

void CBatchProcessDlg::QueueFaxJob(CFaxJob aJob)
{
	myfaxJobs.Add(aJob);
	SendNextFax();//if not already sending a fax send now

}

void CBatchProcessDlg::ResetTimer(void)
{
	iCountDown = 60;
}
LRESULT CBatchProcessDlg::OnResetTimer(WPARAM wp, LPARAM lp)
{
	ResetTimer();
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnStartTimer(WPARAM wp, LPARAM lp)
{
	//PrintOut("Start Timer");
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnFireTimer(WPARAM wp, LPARAM lp)
{
	if(sendingFax)
	{
		if(iCountDown > 0)
		{
			iCountDown--;
			char buffer[200];
			sprintf(buffer, "%d", iCountDown);
			PrintOut(buffer);
		}
		else if (iCountDown == 0)
		{
			iCountDown--;
			PrintOut("Fax Process Timed Out. There may be an issue fax dll. Attempting to restart the fax process.");
			pt_faxer->SetStop(true);
			DWORD dwExitCode;

		// force the thread to stop but leave immediately (the second parameter = 0) 
			pt_faxer->Stop(dwExitCode, 0);
			iFailCount++;
			if(iFailCount < 5)
			{
				sendingFax = false;
				SendNextFax();
			}
			else
			{
				
				PrintOut("The fax process has restarted 5 times due to time out. Stopping the process and moving on to the next fax");
				CFaxJob aJob = (CFaxJob) myfaxJobs.GetAt(0);
				PrintFail("Send Fax Time Out:" + aJob.GetIDMessage());
				//AddToFailLog(aJob.GetDiv(), "Send Fax Time Out:" + aJob.GetIDMessage());
				ResetTimer();
				sendingFax = false;
				iFailCount = 0;
				myfaxJobs.RemoveAt(0);
				SendNextFax();
			}
		}
	}
	
	return LRESULT();
}

//void CBatchProcessDlg::AddToFailLog(int div, CString str)
//{
	//CLogEntry aEntry;
	//aEntry.SetEntry(str, div, true);
	//myLogEntrys.Add(aEntry);
	//WriteNextEntry();
	/*HRESULT hr = S_OK;
    try
    {
		//PrintOut("Init database connection write fail log");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BatcherSettings.mdb;Jet OLEDB:Database Password=cepher;");

       _RecordsetPtr pRst = NULL;

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {
            
		  PrintOut("Failed creating settings record set instance for fail log");		  
            
      }
	  else
	  {		
		 
		CString sql;
		sql = "SELECT * FROM FailLog"; 		
		
		pRst->Open(sql.GetBuffer(), strCnn, adOpenDynamic, adLockOptimistic, adCmdText);
		
		
			COleDateTime aTime;
			aTime = COleDateTime::GetCurrentTime();
			
			pRst->AddNew();
			pRst->Fields->Item["div"]->Value = div;
			pRst->Fields->Item["Date"]->Value = _variant_t(aTime);
			pRst->Fields->Item["Error"]->Value = err.GetBuffer();	
			pRst->Fields->Item["Message"]->Value = str.GetBuffer();	
			
			
			pRst->Update();
			
			
		
	  }
	  
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Error: " + GetCString(ce.Description());
		this->m_output.AddString(str);	
		this->m_output.SetCurSel(m_output.GetCount()-1);		
		this->m_output.RedrawWindow();
		//PrintOut(str);		
		
	}
	
  CoUninitialize();  */
//}

void CBatchProcessDlg::MarkOrderComplete(CFaxJob aJob)
{
	_RecordsetPtr pRst;
	HRESULT hr = S_OK;
    try
    {
		//PrintOut("Init database connection");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

      

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {            
		  PrintOut("Failed creating record set instance");		 
      }
	  else
	  {
		  CString sql;
			sql = "SELECT * FROM t_Orders WHERE Ord_Id=" + aJob.GetOrderID();
			

		  PrintOut("Submit SQL: " + sql);
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);

		if(!pRst->EndOfFile) pRst->MoveFirst(); 
		else 
		{
			PrintOut("No Order Found");
			//return -1;
		}
		//
		if(!pRst->EndOfFile)//&& (!stop))
		{	
			
			PrintOut("Mark Fax Job complete");
			
			pRst->Fields->Item["Ord_Complete"]->Value = true;
			
			COleDateTime today = COleDateTime::GetCurrentTime();
			
			pRst->Fields->Item["Ord_Sent_On"]->Value = _variant_t(today);
			
			pRst->Update();
						
		}		
	  }	 
	   CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Mark Fax Job Complete Error: " + GetCString(ce.Description());
		PrintOut(str);		
		
	}
	
 
}
LRESULT CBatchProcessDlg::OnFaxJobStart(WPARAM wp, LPARAM lp)
{	
	theFaxJob.ClearAll();	
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnFaxJobEnd(WPARAM wp, LPARAM lp)
{	
	QueueFaxJob(theFaxJob);
	
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnFaxJobCamID(WPARAM wp, LPARAM lp)
{	
	theFaxJob.SetCampaignID(wp);
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnFaxJobOrderID(WPARAM wp, LPARAM lp)
{	
	theFaxJob.SetOrderID(wp);
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnFaxJobDiv(WPARAM wp, LPARAM lp)
{	
	theFaxJob.SetDiv(wp);
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnFaxJobSendCoverpage(WPARAM wp, LPARAM lp)
{	
	theFaxJob.SetSendCoverpage(wp);
	return LRESULT();
}

LRESULT CBatchProcessDlg::OnFaxJobAgentEmailChar(WPARAM wp, LPARAM lp)
{	
	char c = (char) wp;
	CString str = theFaxJob.GetAgentEmail();
	str = str + c;
	theFaxJob.SetAgentEmail(str);
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnFaxJobAgentNameChar(WPARAM wp, LPARAM lp)
{	
	char c = (char) wp;	
	CString str = theFaxJob.GetAgentname();
	str = str + c;
	theFaxJob.SetAgentname(str);
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnFaxJobCompanyChar(WPARAM wp, LPARAM lp)
{	
	char c = (char) wp;
	CString str = theFaxJob.GetCompany();
	str = str + c;
	theFaxJob.SetCompany(str);
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnFaxJobContactChar(WPARAM wp, LPARAM lp)
{	
	char c = (char) wp;	
	CString str = theFaxJob.GetContact();
	str = str + c;
	theFaxJob.SetContact(str);
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnFaxJobFaxNumChar(WPARAM wp, LPARAM lp)
{	
	char c = (char) wp;	
	CString str = theFaxJob.GetFaxnum();
	str = str + c;
	theFaxJob.SetFaxnum(str);
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnFaxJobFilenameChar(WPARAM wp, LPARAM lp)
{	
	char c = (char) wp;	
	CString str = theFaxJob.GetFilename();
	str = str + c;
	theFaxJob.SetFilename(str);
	return LRESULT();
}
//
LRESULT CBatchProcessDlg::OnFaxJobStatesChar(WPARAM wp, LPARAM lp)
{	
	char c = (char) wp;	
	CString str = theFaxJob.GetStates();
	str = str + c;
	theFaxJob.SetStates(str);
	return LRESULT();
}
void CBatchProcessDlg::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	if(m_output&&(cx>0)&&(cy>0))
	{	
		this->m_output.MoveWindow(10,120,cx-20,cy-130);
	}
	// TODO: Add your message handler code here
}
LRESULT CBatchProcessDlg::OnProcessVerified(WPARAM wp, LPARAM lp)
{
	m_leadtype.SetWindowText("Verified");
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnProcessInternet(WPARAM wp, LPARAM lp)
{
	m_leadtype.SetWindowText("Internet");
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnProcessOnePass(WPARAM wp, LPARAM lp)
{
	m_leadtype.SetWindowText("One-Pass");
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnProcessPriority(WPARAM wp, LPARAM lp)
{
	m_ordertype.SetWindowText("Priority");
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnProcessSentZero(WPARAM wp, LPARAM lp)
{
	m_ordertype.SetWindowText("Sent Zero on Last Order");
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnProcessNewCustomer(WPARAM wp, LPARAM lp)
{
	m_ordertype.SetWindowText("New Customer");
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnProcessNormal(WPARAM wp, LPARAM lp)
{
	m_ordertype.SetWindowText("Normal");
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnProcessDone(WPARAM wp, LPARAM lp)
{
	//m_leadtype.SetWindowText("");
	//m_ordertype.SetWindowText("");
	if(theProcessQueue.GetCount() > 0)
	{
		processing = true;
		CProcessJob aJob;
		aJob = theProcessQueue.GetNextJob();
		if(aJob.GetiCategory() == PO_RETRY_SENDING)
		{
			manualProcessThread.PostCommand(PO_RETRY_SENDING);
		}
		else if(aJob.GetiCategory() == PO_STATS)
		{
			manualProcessThread.PostCommand(PO_STATS);
		}
		else if(aJob.GetiCategory() == PO_GEN_ORDERS)
		{
			manualProcessThread.PostCommand(PO_GEN_ORDERS);
		}
		else
		{
			manualProcessThread.PostCommand(aJob.GetiCategory());
			manualProcessThread.PostCommand(aJob.GetGroup());
			manualProcessThread.PostCommand(aJob.GetState());
			manualProcessThread.PostCommand(CMD_PROCESS);
		}		
		//manualProcessThread.PostCommand(aJob.GetiCategory());
		//manualProcessThread.PostCommand(aJob.GetGroup());
		//manualProcessThread.PostCommand(CMD_PROCESS);
	}
	else
	{
		processing = false;
		CheckForDone();		
		ProcessButtonsEnable(true);
		
	}
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnProgressAll(WPARAM wp, LPARAM lp)
{
	
	m_progress_all.SetRange(0, lp);
	m_progress_all.SetPos(wp);
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnProgress(WPARAM wp, LPARAM lp)
{
	m_progress.SetRange(0, lp);
	m_progress.SetPos(wp);
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnRetrySending(WPARAM wp, LPARAM lp)
{
	this->m_leadtype.SetWindowText("Retry Sending");
	this->m_ordertype.SetWindowText("Retry Sending");
	return LRESULT();
}
void CBatchProcessDlg::OnBnClickedButtonVerified()
{
	// TODO: Add your control notification handler code here
	ProcessButtonsEnable(false);
	//UpdateData();
	int resID = 5003 + m_groupIndex;
	//int resStateID = 5101 + m_comboStatesIndex;

	int resStateID = 5101 + m_comboStates.GetCurSel();
	manualProcessThread.PostCommand(resStateID);

	manualProcessThread.PostCommand(PO_VERIFIED);
	
	manualProcessThread.PostCommand(resID);
	manualProcessThread.PostCommand(CMD_PROCESS);
	
	//manualProcessThread.SetCategory(PO_VERIFIED, resID);
	//OnManualProcess();
}

void CBatchProcessDlg::OnBnClickedButtonInternet()
{
	// TODO: Add your control notification handler code here
	ProcessButtonsEnable(false);
	int resID = 5003 + m_groupIndex;
	//manualProcessThread.SetCategory(PO_INTERNET, resID);
	int resStateID = 5101 + m_comboStates.GetCurSel();
	manualProcessThread.PostCommand(resStateID);
	manualProcessThread.PostCommand(PO_INTERNET);
	manualProcessThread.PostCommand(resID);
	manualProcessThread.PostCommand(CMD_PROCESS);
	//OnManualProcess();
}

void CBatchProcessDlg::OnBnClickedButtonOnepass()
{
	// TODO: Add your control notification handler code here
	ProcessButtonsEnable(false);
	int resID = 5003 + m_groupIndex;
	//manualProcessThread.SetCategory(PO_ONEPASS, resID);
	int resStateID = 5101 + m_comboStates.GetCurSel();
	manualProcessThread.PostCommand(resStateID);
	manualProcessThread.PostCommand(PO_ONEPASS);
	manualProcessThread.PostCommand(resID);
	manualProcessThread.PostCommand(CMD_PROCESS);
	//OnManualProcess();
}

void CBatchProcessDlg::OnCbnSelchangeComboGroup()
{
	// TODO: Add your control notification handler code here
	UpdateData();

	if( m_groupIndex < 0 ) return;
	//CString str;
	//char buffer[200];
	//int resID = 5003 + m_groupIndex;
	
	//sprintf(buffer, "%d", resID);
	//str = buffer;
	//PrintOut(str);
}

void CBatchProcessDlg::OnManualProcess(void)
{
	manualProcessThread.SetParentWindow(this->m_hWnd);
	if(manualProcessThread.IsAlive())
	{		
		manualProcessThread.PostCommand(CMD_PROCESS);
	}
	else
	{		
		manualProcessThread.Start();
		manualProcessThread.PostCommand(CMD_PROCESS);
	}
}

void CBatchProcessDlg::ProcessButtonsEnable(bool bState)
{
	
	m_button_stop.EnableWindow(!bState);
	m_gen.EnableWindow(bState);
	m_retry.EnableWindow(bState);
	m_button_all.EnableWindow(bState);
	m_button_verified.EnableWindow(bState);
	m_button_internet.EnableWindow(bState);
	m_button_onepass.EnableWindow(bState);
	m_group.EnableWindow(bState);
	m_stats.EnableWindow(bState);
	m_comboStates.EnableWindow(bState);
}

//void CBatchProcessDlg::AddToLog(int div, CString str)
//{
	//CLogEntry aEntry;
	//aEntry.SetEntry(str, div, false);
	//myLogEntrys.Add(aEntry);
	//WriteNextEntry();
	
//}

/*void CBatchProcessDlg::WriteNextEntry(void)
{
	if((myLogEntrys.GetCount()>0) && (!writingLog))
	{
		CLogEntry aEntry = (CLogEntry) myLogEntrys.GetAt(0);
		if(pt_logger)
		{
			delete pt_logger;
		}
		pt_logger = new CLogThread();
		pt_logger->SetParentWindow(this->m_hWnd);		
		pt_logger->SetEntry(aEntry);
		writingLog = true;
		
		pt_logger->Start();		
	}
}*/
LRESULT CBatchProcessDlg::OnEmailFailChar(WPARAM wp, LPARAM lp)
{
	int iMessage = 7000 + wp;
	//theLogger.PostCommand(iMessage);
	//char c = (char) wp;
	//CString str;
	//str = str + c;
	//myEmailFailLogEntry.Append(str, lp, true);
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnEmailFailEnd(WPARAM wp, LPARAM lp)
{
	//theLogger.PostCommand(-2);
	//AddToFailLog(lp, myEmailFailLogEntry.GetEntry());
	//myEmailFailLogEntry.Clear();
	return LRESULT();
}
LRESULT CBatchProcessDlg::OnEmailFailStart(WPARAM wp, LPARAM lp)
{
	//theLogger.PostCommand(-1);
	//myEmailFailLogEntry.Clear();
	return LRESULT();
}
void CBatchProcessDlg::PrintFail(CString str)
{
	
	SendMessage(EMAIL_FAIL_START, 0, 0);
	int iLen = str.GetLength();
	for(int x = 0; x < iLen; x++)
	{
		int iLetter = (int) str.GetAt(x);
		SendMessage(EMAIL_FAIL_CHAR, iLetter, 0);
	}
	SendMessage(EMAIL_FAIL_END, 0, 0);
	
}

void CBatchProcessDlg::OnFileExit()
{
	// TODO: Add your command handler code here
	OnOK();
}

void CBatchProcessDlg::OnHelpAbout()
{
	// TODO: Add your command handler code here
	CAboutDlg dlgAbout;
	dlgAbout.DoModal();
}

bool CBatchProcessDlg::CheckForDone(void)
{
	bool done = false;
	if((myfaxJobs.GetCount()==0) && (!sendingFax) && (!processing) && (theProcessQueue.GetCount()==0))
	{
		PrintOut("Processing and Faxing Done");
		done = true;
		if(autorun)
		{
			OnBnClickedCancel();//close the app if it was autorun batch
		}
	}
	return done;
}

void CBatchProcessDlg::OnBnClickedButtonRetry()
{
	ProcessButtonsEnable(false);	
		
	manualProcessThread.PostCommand(PO_RETRY_SENDING);
	// TODO: Add your control notification handler code here
}

void CBatchProcessDlg::OnBnClickedButtonGenOrders()
{
	// TODO: Add your control notification handler code here
	ProcessButtonsEnable(false);	
		
	manualProcessThread.PostCommand(PO_GEN_ORDERS);
}

void CBatchProcessDlg::WriteLog(CString str)
{
	///////////////////////////////////////////////////////////////////
	HRESULT hr;
	hr = S_OK;	
	
	try
    {
		 CoInitialize(NULL);

		_RecordsetPtr pRst;		
		_bstr_t strCnn;
		pRst = NULL;

		strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BatcherLog.mdb;Jet OLEDB:Database Password=cepher;";
		 // Call Create instance to instantiate the Record set
		CString sql = "SELECT * FROM Feedback";
		hr = pRst.CreateInstance(__uuidof(Recordset));
		if(FAILED(hr))
		{            
			PrintOut("Failed creating batcher log record set instance");
		}
		else
		{			
			pRst->Open(sql.GetBuffer(), strCnn, adOpenForwardOnly, adLockOptimistic, adCmdText);
			COleDateTime aTime;
			aTime = COleDateTime::GetCurrentTime();
			
			pRst->AddNew();			
			pRst->Fields->Item["Date"]->Value = _variant_t(aTime);	
			pRst->Fields->Item["Message"]->Value = str.GetBuffer();	
			pRst->Update();
			
		}
		CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "Error WriteLog: " + GetCString(ce.Description());
		PrintOut(str);	
		
	}
	///////////////////////////////////////////////////////////////////
}

void CBatchProcessDlg::OnBnClickedButtonStats()
{
	ProcessButtonsEnable(false);	
		
	manualProcessThread.PostCommand(PO_STATS);
	// TODO: Add your control notification handler code here
}

void CBatchProcessDlg::OnCbnSelchangeComboStates()
{
	// TODO: Add your control notification handler code here
}
