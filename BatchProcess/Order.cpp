#include "StdAfx.h"
#include ".\order.h"

#define DONT_CARE	0

#define PURCHASE	1
#define REFINANCE	2 

#define POOR		1
#define FAIR		2
#define GOOD		3
#define EXCELLENT	4

#define ENGLISH		1
#define SPANISH		2

#define SFR			1
#define COMMERCIAL	2
#define MOBILE		3

#define COUNTRYLOAN 1
#define FASTFUNDING	2

#define I_WANT		1
#define I_DONT_WANT	2

COrder::COrder(void)
: sStates(_T(""))
, iCompanyID(0)
, sContact(_T(""))
, iOrderID(0)
, sFaxNumber(_T(""))
, iQuantity(0)
, sEmail(_T(""))
, sAgent(_T(""))
, fRateFilter(0)
, sZipFilter(_T(""))
, uliDesiredLoanFilter(0)
, sAreacodeFilter(_T(""))
, bOnePass(false)
, bSendExcel(false)
, bSendFax(false)
, bVerified(false)
, bSendWord(false)
, bOneLeadPerPage(false)
, bSendEmailNote(false)
, bPrint(false)
, iLookingTo(0)
, iLang(0)
, iCreditScore(0)
, iHomeType(0)
, iLocation(0)
, bPriority(false)
, bSendMini1003(false)
, bAdCampaign(false)
, bSubprime(false)
, iCashout(0)
, i2ndMortgage(0)
, iDebtCon(0)
, sCompanyName(_T(""))
, iAgentID(0)
, pt_LeadDBMap(NULL)
, iSFR(0)
, iMobile(0)
, iCommercial(0)
, sCity(_T(""))
, fLTV(0)
, bLTVAbove(false)
, iMaxNeeded(1)
, iLastOrderQuantity(0)
, iCompanyOrders(0)
, pt_stop(NULL)
, iCampaignID(0)
, iMyUseLimit(0)
, iDefaultUseLimit(0)
, iMyAgeLimit(0)
, iDefaultAgeLimit(0)
, oleOrderDate(COleDateTime::GetCurrentTime())
{
	srand( (unsigned)time( NULL ) );
}

COrder::~COrder(void)
{
	
}

CString COrder::ToString(void)
{
	
	CString str;
	try
	{
		char  buffer[200];
		sprintf( buffer,     "(CompanyID:%d) ", iCompanyID );
		str = buffer;
		str = str + "(Company Name:" + sCompanyName + ") ";
		str = str + "(States:" + sStates + ") ";
		str = str + "(contact:" + sContact + ") ";
		sprintf( buffer,     "(OrderID:%d) ", iOrderID );
		str = str + buffer;
		str = str + "(Fax Number:" + sFaxNumber + ") ";
		sprintf( buffer,     "(Quantity:%d) ", iQuantity );
		str = str + buffer;//
		str = str + "(Email:" + sEmail + ") ";
		str = str + "(Agent:" + sAgent + ") ";
		sprintf( buffer,     "(Rate Filter:%f) ", fRateFilter );
		str = str + buffer;//
		str = str + "(Zip Filter:" + sZipFilter + ") ";
		sprintf( buffer,     "(Desired Loan Filter:%d) ", uliDesiredLoanFilter );
		str = str + buffer;
		str = str + "(Areacode Filter:" + sAreacodeFilter + ") ";
		str = str + "(One Pass:" + BooltoStr(bOnePass) + ") ";
		str = str + "(Send Excel:" + BooltoStr(bSendExcel) + ") ";
		str = str + "(Send Fax:" + BooltoStr(bSendFax) + ") ";
		str = str + "(Verified:" + BooltoStr(bVerified) + ") ";
		str = str + "(Send Word:" + BooltoStr(bSendWord) + ") ";
		str = str + "(One lead per page:" + BooltoStr(bOneLeadPerPage) + ") ";
		str = str + "(Send email note:" + BooltoStr(bSendEmailNote) + ") ";
		str = str + "(Print:" + BooltoStr(bPrint) + ") ";
		str = str + "(Looking to:" + GetLookingtoStr(iLookingTo) + ") ";
		str = str + "(Language:" + GetLangStr(iLang) + ") ";
		str = str + "(CreditScore:" + GetCreditScoreStr(iCreditScore) + ") ";
		str = str + "(Home type:" + GetHomeTypeStr(iHomeType) + ") ";
		str = str + "(SFR:" + GetSFR() + ") ";
		str = str + "(Mobile:" + GetMobile() + ") ";
		str = str + "(Commercial:" + GetCommercial() + ") ";
		str = str + "(Location:" + GetLocationStr(iLocation) + ") "; 
		str = str + "(Priority:" + BooltoStr(bPriority) + ") ";
		str = str + "(Send Mini 1003:" + BooltoStr(bSendMini1003) + ") ";
		str = str + "(Ad Campaign:" + BooltoStr(bAdCampaign) + ") ";
		str = str + "(Sub Prime:" + BooltoStr(bSubprime) + ") ";//
		str = str + "(Cashout:" + GetThreeStateStr(iCashout) + ") ";
		str = str + "(2nd Mortgage:" + GetThreeStateStr(i2ndMortgage) + ") ";
		str = str + "(Debt Consolidate:" + GetThreeStateStr(iDebtCon) + ") ";
		str = str + "(City:" + GetCity() + ") ";
		str = str + "(LTV:" + GetLTV() + ") ";
		str = str + "(LTV Above:" + BooltoStr(GetLTVAbove()) + ") ";
		COleDateTime aDate;
		aDate = COleDateTime::GetCurrentTime();
		CString sDate = aDate.Format(_T("%m/%d/%y"));
		str = str + "(Order Date:" + sDate + ") ";

		sprintf( buffer,     "(Default Age Limit::%d) ", iDefaultAgeLimit );
		str = str + buffer;

		sprintf( buffer,     "(Default Use Limit::%d) ", iDefaultUseLimit );
		str = str + buffer;

		sprintf( buffer,     "(My Age Limit::%d) ", iMyAgeLimit );
		str = str + buffer;

		sprintf( buffer,     "(My Use Limit::%d) ", iMyUseLimit );
		str = str + buffer;
		
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrder::ToString";
		PrintOut(err);
		WriteLog(err);
	}
	return str;
}

CString COrder::BooltoStr(bool state)
{
	CString result = "FALSE";	
	try
	{
		if(state)
		{
			result = "TRUE";
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrder::BooltoStr(bool state)";
		PrintOut(err);
		WriteLog(err);
	}
	return result;
}

CString COrder::GetLookingtoStr(int iNum)
{
	CString result;
	try
	{
		switch ( iNum )
		{
			case DONT_CARE:
				result = "Not Filtered";
				break;
			case PURCHASE:
				result = "Purchase";
				break;
			case REFINANCE:
				result = "Refinance";
				break;
			default:
				result = "Error: Not a valid (iLookingTo) value";
				break;           
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrder::GetLookingtoStr(int iNum)";
		PrintOut(err);
		WriteLog(err);
	}
	return result;
}

CString COrder::GetLangStr(int iNum)
{
	CString result;
	try
	{
		switch ( iNum )
		{
			case DONT_CARE:
				result = "Not Filtered";
				break;
			case ENGLISH:
				result = "English";
				break;
			case SPANISH:
				result = "Spanish";
				break;
			default:
				result = "Error: Not a valid (iLang) value";
				break;           
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrder::GetLangStr(int iNum)";
		PrintOut(err);
		WriteLog(err);
	}
	return result;
}

CString COrder::GetCreditScoreStr(int iNum)
{
	CString result;
	try
	{
		switch ( iNum )
		{
			case DONT_CARE:
				result = "Not Filtered";
				break;
			case POOR:
				result = "Poor";
				break;
			case FAIR:
				result = "Fair";
				break;
			case GOOD:
				result = "Good";
				break;
			case EXCELLENT:
				result = "Excellent";
				break;
			default:
				result = "Error: Not a valid (iCreditScore) value";
				break;           
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrder::GetCreditScoreStr(int iNum)";
		PrintOut(err);
		WriteLog(err);
	}
	return result;
}

CString COrder::GetHomeTypeStr(int iNum)
{
	CString result;
	try
	{
		switch ( iNum )
		{
			case DONT_CARE:
				result = "Not Filtered";
				break;
			case SFR:
				result = "SFR";
				break;
			case COMMERCIAL:
				result = "Commercial";
				break;		
			case MOBILE:
				result = "Mobile";
				break;
			default:
				result = "Error: Not a valid (iHomeType) value";
				break;           
		}	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrder::GetHomeTypeStr(int iNum)";
		PrintOut(err);
		WriteLog(err);
	}
	return result;
}

CString COrder::GetLocationStr(int iNum)
{
	CString result;
	try
	{
		switch ( iNum )
		{
			case 0:
				result = "Country Loans";
				break;
			case COUNTRYLOAN:
				result = "Country Loans";
				break;
			case FASTFUNDING:
				result = "Fast Funding";
				break;		
			
			default:
				result = "Error: Not a valid (iLocation) value";
				break;           
		}	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrder::GetLocationStr(int iNum)";
		PrintOut(err);
		WriteLog(err);
	}
	return result;
}

CString COrder::GetThreeStateStr(int iNum)
{
	CString result;
	try
	{
		switch ( iNum )
		{
			case DONT_CARE:
				result = "Not Filtered";
				break;
			case I_WANT:
				result = "I Want";
				break;
			case I_DONT_WANT:
				result = "I Don't Want";
				break;		
			
			default:
				result = "Error: Not a valid (Three State Value) value";
				break;           
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrder::GetThreeStateStr(int iNum)";
		PrintOut(err);
		WriteLog(err);
	}
	return result;
}

void COrder::SetAdCampaign(bool state)
{	
	try
	{
		bAdCampaign = state;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrder::SetAdCampaign(bool state)";
		PrintOut(err);
		WriteLog(err);
	}
}

void COrder::SetOneLeadPerPage(bool state)
{
	bOneLeadPerPage = state;
}

void COrder::SetStates(CString str)
{
	sStates = str.MakeUpper();
}

void COrder::SetCompanyID(long iNum)
{
	iCompanyID = iNum;
}

void COrder::SetContact(CString str)
{
	sContact = str;
}

void COrder::SetOrderID(long iNum)
{
	iOrderID = iNum;
}

void COrder::SetFaxNumber(CString str)
{
	sFaxNumber = str;
}

void COrder::SetQuantity(int iNum)
{
	iQuantity = iNum;
}

void COrder::SetEmail(CString str)
{
	sEmail = str;
}

void COrder::SetAgent(CString str)
{
	sAgent = str;
}

void COrder::SetRateFilter(double fNum)
{
	fRateFilter = fNum;
}

void COrder::SetZipFilter(CString str)
{
	sZipFilter = str;
}

void COrder::SetDesiredLoanFilter(unsigned long ulNum)
{
	uliDesiredLoanFilter = ulNum;
}

void COrder::SetAreacodeFilter(CString str)
{
	sAreacodeFilter = str;
}

void COrder::SetOnePass(bool state)
{
	bOnePass = state;
}

void COrder::SetSendExcel(bool state)
{
	bSendExcel = state;
}

void COrder::SetSendFax(bool state)
{
	bSendFax = state;
}

void COrder::SetVerified(bool state)
{
	bVerified = state;
}

void COrder::SetSendWord(bool state)
{
	bSendWord = state;
}

void COrder::SetSendEmailNote(bool state)
{
	bSendEmailNote = state;
}

void COrder::SetPrint(bool state)
{
	bPrint = state;
}

void COrder::SetLookingTo(int iNum)
{
	iLookingTo = iNum;
}

void COrder::SetLang(int iNum)
{
	iLang = iNum;
}

void COrder::SetCreditScore(int iNum)
{
	iCreditScore = iNum;
}

void COrder::SetHouseType(int iNum)
{
	if((iNum >=0) && (iNum <= 3))
	{
		iHomeType = iNum;
	}
	else
	{
		iHomeType = 0;
	}
	
	
}

void COrder::SetLocation(int iNum)
{
	iLocation = iNum;
}

void COrder::SetPriority(bool state)
{
	bPriority = state;
}

void COrder::SetSendMini1003(bool state)
{
	bSendMini1003 = state;
}

void COrder::SetSubprime(bool state)
{
	bSubprime = state;
}

void COrder::SetCashout(int iNum)
{
	if((iNum>=0)&&(iNum<=2))
	{
		iCashout = iNum;
	}
	else
	{
		iCashout = 0;
	}
}

void COrder::Set2ndMortgage(int iNum)
{
	if((iNum>=0)&&(iNum<=2))
	{
		i2ndMortgage = iNum;
	}
	else
	{
		i2ndMortgage = 0;
	}
}

void COrder::SetDebtCon(int iNum)
{
	if((iNum>=0)&&(iNum<=2))
	{
		iDebtCon = iNum;
	}
	else
	{
		iDebtCon = 0;
	}
}

bool COrder::ReadOrder(_RecordsetPtr rs)
{
	//PrintOut("Reading Order");
	friends.RemoveAll();
	myAssignedLeadIDs.RemoveAll();
	myMatchingLeadIDs.RemoveAll();
	bool success = false;

	SetOrderID(GetFieldLong(rs, "Ord_Id"));
	SetStates(GetFieldString(rs, "Ord_State"));
	SetCompanyID(GetFieldLong(rs, "t_orders.Ord_CompanyID"));
	if(GetCompanyIDLong()==-1) SetCompanyID(GetFieldLong(rs, "Ord_CompanyID"));
	friends.Add(iCompanyID);//in case there is more than one order by the same company

	SetCompanyName(GetFieldString(rs, "Ord_CompanyName"));
	SetContact(GetFieldString(rs, "Ord_Contact"));
	SetFaxNumber(GetFieldString(rs, "Ord_Fax"));
	SetQuantity(GetFieldLong(rs, "Ord_MakeGood"));
	if((GetQuanityLong() < 0)|| (GetQuanityLong() > 5000))
	{
		//bad read, reread the quantity
		SetQuantity(ReReadQuantity());
	}
	SetEmail(GetFieldString(rs, "Ord_Email"));
	SetAgent(GetFieldString(rs, "Ord_AgentName"));
	double rate = atof(GetFieldString(rs, "Ord_Rate_Filter"));	
	SetRateFilter(rate);
	SetZipFilter(GetFieldString(rs, "Ord_ZipCode_Filter"));
	CString str = GetFieldString(rs, "Ord_DesiredLoanAmt_Filter");
	str.Remove(',');
	str.Remove('$');
	unsigned long loan = atol(str);
	SetDesiredLoanFilter(loan);
	SetAreacodeFilter(GetFieldString(rs, "Ord_AreaCode_Filter"));
	SetOnePass(GetFieldBool(rs, "Ord_OnePass"));
	SetSendExcel(GetFieldBool(rs, "Ord_Excel"));
	SetSendFax(GetFieldBool(rs, "Ord_Send_Fax"));
	SetVerified(GetFieldBool(rs, "Ord_Verified"));
	SetSendWord(GetFieldBool(rs, "Ord_Word"));
	SetOneLeadPerPage(GetFieldBool(rs, "Ord_1Lead"));
	SetSendEmailNote(GetFieldBool(rs, "Ord_EmailNote"));
	SetPrint(GetFieldBool(rs, "Ord_Print"));
	SetLookingTo(GetFieldLong(rs, "lookingToID"));
	SetLang(GetFieldLong(rs, "lang"));
	SetCreditScore(GetFieldLong(rs, "creditscore"));
	SetHouseType(GetFieldLong(rs, "homeTypeInt"));
	SetSFR(GetFieldLong(rs, "SFR"));
	SetMobile(GetFieldLong(rs, "Mobile"));
	SetCommercial(GetFieldLong(rs, "Commercial"));
	SetLocation(GetFieldLong(rs, "locationInt"));
	SetPriority(GetFieldBool(rs, "priority"));
	SetSendMini1003(GetFieldBool(rs, "Ord_sendMini1003"));
	SetAdCampaign(GetFieldBool(rs, "adCampaign"));
	SetAgentID(GetFieldLong(rs, "agentID"));
	SetCashout(GetFieldLong(rs, "cashOut"));
	Set2ndMortgage(GetFieldLong(rs, "2ndMortgage"));
	SetDebtCon(GetFieldLong(rs, "debtCon"));
	SetSubprime(GetFieldBool(rs, "subPrime"));	
	SetCity(GetFieldString(rs, "City"));
	SetLTV(GetFieldDouble(rs, "LTV"));
	SetLTVAbove(GetFieldBool(rs, "LTVAbove"));
	SetLastOrderQuantity(GetFieldLong(rs, "lastOrderSent"));//countOrders
	SetCompanyOrders(GetFieldLong(rs, "countOrders"));
	SetCampaignID(GetFieldLong(rs, "campaignID"));
	SetMyUseLimit(GetFieldLong(rs, "UseLimit"));
	SetMyAgeLimit(GetFieldLong(rs, "AgeLimit"));
	char buffer[200];
	sprintf(buffer, "%d", iMyUseLimit);
	str = "UseLimit:";
	str = str + buffer;
	//PrintOut(str);
	//PrintOut(GetUseLimit());
	
	ReadFriends();
	
	//ReadLeads();

	success = true;
	return success;
}

void COrder::SetCompanyName(CString name)
{
	sCompanyName = name;
}

CString COrder::GetOrderID(void)
{
	CString id;
	char buffer[200];
	sprintf( buffer, "%d", iOrderID );
	id = buffer;
	return id;
}

CString COrder::GetStates(void)
{
	return sStates;
}

CString COrder::GetCompanyID(void)
{
	CString id;
	char buffer[200];
	sprintf( buffer, "%d", iCompanyID );
	id = buffer;
	return id;
}

CString COrder::GetCompanyName(void)
{
	return sCompanyName;
}

CString COrder::GetContact(void)
{
	return sContact;
}

CString COrder::GetFaxNumber(void)
{
	return sFaxNumber;
}

CString COrder::GetQuantity(void)
{
	CString sNum;
	char buffer[200];
	sprintf( buffer, "%d", iQuantity );
	sNum = buffer;
	return sNum;
}

CString COrder::GetEmail(void)
{
	return sEmail;
}

CString COrder::GetAgent(void)
{
	return sAgent;
}

void COrder::SetAgentID(long iNum)
{
	iAgentID = iNum;
}

CString COrder::GetAgentID(void)
{
	CString sNum;
	char buffer[200];
	sprintf( buffer, "%d", iAgentID );
	sNum = buffer;
	return sNum;
}

CString COrder::GetRateFilter(void)
{
	CString sNum;
	char buffer[200];
	sprintf( buffer, "%f", fRateFilter );
	sNum = buffer;
	return sNum;
}

CString COrder::GetZipFilter(void)
{
	return sZipFilter;
}

CString COrder::GetDesiredLoanFilter(void)
{
	CString sNum;
	char buffer[200];
	sprintf( buffer, "%d", uliDesiredLoanFilter );
	sNum = buffer;
	return sNum;
}

CString COrder::GetAreacodeFilter(void)
{
	return sAreacodeFilter;
}

CString COrder::GetCashout(void)
{
	return GetThreeStateStr(iCashout);
}

CString COrder::Get2ndMortgage(void)
{
	return GetThreeStateStr(i2ndMortgage);
}

CString COrder::GetDebtCon(void)
{
	return GetThreeStateStr(iDebtCon);
}

bool COrder::GetOnePass(void)
{
	return bOnePass;
}

bool COrder::GetExcel(void)
{
	return bSendExcel;
}

bool COrder::GetSendFax(void)
{
	return bSendFax;
}

bool COrder::GetVerified(void)
{
	return bVerified;
}

bool COrder::GetSendWord(void)
{
	return bSendWord;
}

bool COrder::GetOneLeadPerPage(void)
{
	return bOneLeadPerPage;
}

bool COrder::GetSendEmailNote(void)
{
	return bSendEmailNote;
}

bool COrder::GetPrint(void)
{
	return bPrint;
}

CString COrder::GetLookingto(void)
{
	return GetLookingtoStr(iLookingTo);
}

CString COrder::GetLang(void)
{
	return GetLangStr(iLang);
}

CString COrder::GetCreditscore(void)
{
	return GetCreditScoreStr(iCreditScore);
}

CString COrder::GetHomeType(void)
{
	return GetHomeTypeStr(iHomeType);
}

CString COrder::GetLocation(void)
{
	return GetLocationStr(iLocation);
}

bool COrder::GetPriority(void)
{
	return bPriority;
}

bool COrder::GetSendMini1003(void)
{
	return bSendMini1003;
}

bool COrder::GetAdCampaign(void)
{
	return bAdCampaign;
}

bool COrder::GetSubPrime(void)
{
	return bSubprime;
}


CString COrder::GetSQL(void)
{
	return GetSQLNew();// new better sql
	//the following code is not used at this time
	CString sql;	

	sql = sql + " SELECT ";
	sql = sql + " avail.t_leads.lead_Id, Count(po.PO_Company_ID) AS CountOfPO_Company_ID, avail.t_Leads.lead_FirstName, avail.t_Leads.lead_LastName, ";
	sql = sql + " avail.t_Leads.lead_CoApp, avail.t_Leads.lead_City, avail.t_Leads.lead_State,avail.t_Leads.lead_Address,avail.t_Leads.lead_Zip,avail.t_Leads.lead_WorkPhone, ";
	sql = sql + " avail.t_Leads.lead_HomePhone, avail.t_Leads.lead_HouseType, avail.t_Leads.lead_CurrentValue,  avail.t_Leads.lead_DesiredLoan, avail.t_Leads.lead_1stBalance,  ";
	sql = sql + " avail.t_Leads.lead_Rate,  avail.t_Leads.lead_FixedAdjust, avail.t_Leads.lead_Payment,  avail.t_Leads.lead_Late, ";
	sql = sql + " avail.t_Leads.lead_Credit, avail.t_Leads.lead_WorkInfo, avail.t_Leads.lead_TimePeriod, avail.t_Leads.lead_Salary, avail.t_Leads.lead_YouShouldCall, ";
	sql = sql + " avail.t_Leads.lead_LookingTo, avail.t_Leads.lead_Email, avail.t_Leads.Sys_Cre_Date, ";
	sql = sql + " avail.t_Leads.lead_Verified, avail.t_Leads.lookingToID, avail.t_Leads.langInt, avail.t_Leads.VerifiedBy, avail.t_Leads.PersonalTouch  ";
	sql = sql + " FROM ";
	sql = sql + " (SELECT * ";
	sql = sql + " FROM t_leads INNER JOIN [SELECT * ";
	sql = sql + " FROM  ";
	sql = sql + " (SELECT * ";
	//sql = sql + " FROM t_Processed_Orders  ";
	sql = sql + " FROM (SELECT t_Processed_Orders.* FROM t_Processed_Orders INNER JOIN  t_leads ON t_Processed_Orders.PO_Lead_ID = t_Leads.lead_Id  WHERE  (t_Leads.Sys_Cre_Date < t_Processed_Orders.PO_Lead_SentOn)) AS po ";
	sql = sql + " INNER JOIN ";
	sql = sql + " (SELECT * ";
	sql = sql + " FROM friends ";
	sql = sql + " WHERE (((friends.cust_id)=";
	sql = sql + GetCompanyID();
	sql = sql + "))) as myFriends1 ";
	//sql = sql + " ON t_Processed_Orders.PO_Company_ID = myFriends1.friend_id) AS myLeads1 ";
	sql = sql + " ON po.PO_Company_ID = myFriends1.friend_id) AS myLeads1 ";
	sql = sql + " RIGHT JOIN t_Leads ON myLeads1.PO_Lead_ID = t_Leads.lead_Id  ";
	sql = sql + " WHERE (((myLeads1.PO_Lead_ID) Is Null))]. AS dontHaveYet ON t_leads.lead_Id = dontHaveYet.lead_Id ";
	
	//sql = " SELECT * FROM t_leads ";
	//filter area
	sql = sql + "\r\nWHERE t_leads.lead_Verified=";
	sql = sql + BooltoStr(GetVerified());
	sql = sql + "\r\n AND t_leads.lead_DoNotUse=FALSE";
	sql = sql + "\r\n AND t_leads.dupe=FALSE";
	sql = sql + "\r\n AND ";
	if(bOnePass)
	{
		sql = sql + "DateDiff('d',t_leads.Sys_Cre_Date,Now())>";
		sql = sql + GetAgeLimit();
	}
	else
	{
		
		sql = sql + "DateDiff('d',t_leads.Sys_Cre_Date,Now())<=";
		sql = sql + GetAgeLimit();
		sql = sql + "\r\n AND ";
		sql = sql + "t_leads.Sys_Cre_Date <= Now() ";
	}
	
	if(fRateFilter>0)
	{
		sql = sql + "\r\n AND t_leads.lead_Rate>=";
		sql = sql + GetRateFilter();
	}	
	if(uliDesiredLoanFilter>0)
	{
		sql = sql + "\r\n AND t_leads.lead_DesiredLoan>=";
		sql = sql + this->GetDesiredLoanFilter();
	}	
	if(sStates.GetLength()>0)
	{
		CString statesSQL = sStates;
		statesSQL.Replace(" ", ""); //remove spaces
		sql = sql + "\r\n AND (t_leads.lead_State='";
		statesSQL.Replace(",", "' OR t_leads.lead_State='");
		sql = sql + statesSQL;
		sql = sql + "')";
	}
	if(sZipFilter.GetLength()>0)	
	{
		CString zipSQL = sZipFilter;
		zipSQL.Replace(" ", ""); //remove spaces
		sql = sql + "\r\n AND (t_leads.lead_Zip LIKE '";
		zipSQL.Replace(",", "%' OR t_leads.lead_Zip LIKE '");
		sql = sql + zipSQL;
		sql = sql + "%')";
	}
	if(sAreacodeFilter.GetLength()>0)	
	{
		sAreacodeFilter.Replace(" ", ""); //remove spaces
		sql = sql + "\r\n AND (";
		CString homeSQL = "\r\n(t_leads.lead_HomePhone LIKE '" + sAreacodeFilter;
		CString workSQL =  "\r\n(t_leads.lead_WorkPhone LIKE '" + sAreacodeFilter;		
		homeSQL.Replace(",", "%') \r\n OR (t_leads.lead_HomePhone LIKE '");
		workSQL.Replace(",", "%') \r\n OR (t_leads.lead_WorkPhone LIKE '");
		homeSQL = homeSQL + "%')";
		workSQL = workSQL + "%')";
		sql = sql + homeSQL;
		sql = sql + "\r\n OR ";
		sql = sql + workSQL;
		sql = sql + ")";
	}
	if((iLookingTo > 0) && (iLookingTo <= 2))
	{
		sql = sql + "\r\n AND ";
		if(iLookingTo == REFINANCE) sql = sql + " t_leads.lead_LookingTo NOT LIKE 'purch%' ";
		if(iLookingTo == PURCHASE) sql = sql + " t_leads.lead_LookingTo LIKE 'purch%' ";

	}
	if((iLang > 0) && (iLang <= 2))
	{
		sql = sql + "\r\n AND ";
		if(iLang == ENGLISH) sql = sql + " t_leads.langInt<2 ";			
		
		if(iLang == SPANISH) sql = sql + " t_leads.langInt=2 ";
	}
	if(iCreditScore > 0)
	{
		
		sql = sql + "\r\n AND (";
		if(bSubprime)
		{
			
			switch ( iCreditScore )
			{
				case EXCELLENT:				
					sql = sql + " (t_leads.lead_Credit = 'Excellent') OR ";
					
				case GOOD:
					sql = sql + " (t_leads.lead_Credit = 'Good') OR ";
				
				case FAIR:
					sql = sql + " (t_leads.lead_Credit = 'Fair') OR ";
				
				case POOR:
					sql = sql + " (t_leads.lead_Credit = 'Poor') OR";
					sql = sql + " (t_leads.lead_Credit = 'Major_Credit_Issues') ";
				break;
				default:
					//invalid iCreditScore add nothing to the SQL
				break;
	            
			}

		}
		else
		{
			switch ( iCreditScore )
			{
				case POOR:
					sql = sql + " (t_leads.lead_Credit = 'Poor') OR";
					sql = sql + " (t_leads.lead_Credit = 'Major_Credit_Issues') OR";
					
				case FAIR:
					sql = sql + " (t_leads.lead_Credit = 'Fair') OR ";
				
				case GOOD:
					sql = sql + " (t_leads.lead_Credit = 'Good') OR ";
				
				case EXCELLENT:
					sql = sql + " (t_leads.lead_Credit = 'Excellent') ";
				break;
				default:
					//invalid iCreditScore add nothing to the SQL
				break;
			}
		}
		sql = sql + ")";
	}
	if((iCashout==1) || (i2ndMortgage==1) || (iDebtCon==1))
	{
		sql = sql + "\r\n AND (";
		
		if(iCashout==1)
		{
			sql = sql + " (t_leads.lead_LookingTo LIKE '%cash%') ";
			if((i2ndMortgage==1) || (iDebtCon==1))
			{
				sql = sql + " OR ";
			}
		}				
		if(i2ndMortgage==1)
		{
			sql = sql + " (t_leads.lead_LookingTo LIKE '%Second%') ";
			if(iDebtCon == 1)
			{
				sql = sql + " OR ";
			}
		}				
		if(iDebtCon==1)
		{
			sql = sql + " (t_leads.lead_LookingTo LIKE '%Consol%') ";
		}	

		sql = sql + ")";
	}					
	if(iCashout==2)
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (t_leads.lead_LookingTo NOT LIKE '%cash%') ";	
		sql = sql + ")";
	}
	if(i2ndMortgage==2)				
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (t_leads.lead_LookingTo NOT LIKE '%Second%') ";
		sql = sql + ")";				
	}
	if(iDebtCon==2)				
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (t_leads.lead_LookingTo NOT LIKE '%Consol%') ";
		sql = sql + ")";				
	}			
	if(iHomeType > 0)
	{
		sql = sql + "\r\n AND (";
		

		switch ( iHomeType )//not finished more needed
			{
				case SFR:
					sql = sql + " (t_leads.lead_HouseType LIKE '%Single Family%') OR";
					sql = sql + " (t_leads.lead_HouseType LIKE 'Condo%') OR";
					sql = sql + " (t_leads.lead_HouseType LIKE 'Town%') OR";
					sql = sql + " (t_leads.lead_HouseType = 'SFR') ";
					
					break;
					
				case COMMERCIAL:
					sql = sql + " (t_leads.lead_HouseType = 'Commercial') OR";
					sql = sql + " (t_leads.lead_HouseType = 'commercial') ";
					break;
				
				case MOBILE:
					sql = sql + " (t_leads.lead_HouseType LIKE '%Mobile%') OR";	
					sql = sql + " (t_leads.lead_HouseType LIKE '%wide%') OR";
					sql = sql + " (t_leads.lead_HouseType LIKE '%Manufac%') ";	
					break;

				default:
					//invalid iHomeType add nothing to the SQL
				break;
			}
		sql = sql + ")";

	}
	if((iSFR==1) || (iCommercial==1) || (iMobile==1))
	{
		sql = sql + "\r\n AND (";
		
		if(iSFR==1)
		{
			sql = sql + " (t_leads.lead_HouseType LIKE '%Single%Family%') OR";
			sql = sql + " (t_leads.lead_HouseType LIKE 'Condo%') OR";
			sql = sql + " (t_leads.lead_HouseType LIKE 'Town%') OR";
			sql = sql + " (t_leads.lead_HouseType LIKE 'SFR') ";

			if((iCommercial==1) || (iMobile==1))
			{
				sql = sql + " OR ";
			}
		}				
		if(iCommercial==1)
		{
			sql = sql + " (t_leads.lead_HouseType LIKE '%commerc%') ";
			
			if(iMobile == 1)
			{
				sql = sql + " OR ";
			}
		}				
		if(iMobile==1)
		{
			sql = sql + " (t_leads.lead_HouseType LIKE '%Mobile%') OR";	
			sql = sql + " (t_leads.lead_HouseType LIKE '%wide%') OR";
			sql = sql + " (t_leads.lead_HouseType LIKE '%Manufac%') ";	
		}	

		sql = sql + ")";
	}					
	if(iSFR==2)
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (t_leads.lead_HouseType NOT LIKE '%Single%Family%') AND";
			sql = sql + " (t_leads.lead_HouseType NOT LIKE 'Condo%') AND";
			sql = sql + " (t_leads.lead_HouseType NOT LIKE 'Town%') AND";
			sql = sql + " (t_leads.lead_HouseType NOT LIKE 'SFR') ";

		sql = sql + ")";
	}
	if(iCommercial==2)				
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (t_leads.lead_HouseType NOT LIKE '%commercial%') ";
			
		sql = sql + ")";				
	}
	if(iMobile==2)				
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (t_leads.lead_HouseType NOT LIKE '%Mobile%') AND";	
		sql = sql + " (t_leads.lead_HouseType NOT LIKE '%wide%') AND";
		sql = sql + " (t_leads.lead_HouseType NOT LIKE '%Manufac%') ";	
		sql = sql + ")";				
	}			
	if(sCity.GetLength()>0)
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (t_leads.lead_City LIKE '%";
		sql = sql + GetCity();
		sql = sql + "%') ";
		sql = sql + ")";

	}
	if(fLTV > 0)
	{
		if(bLTVAbove)
		{
			sql = sql + "\r\n AND (";
			sql = sql + " (t_leads.lead_1stBalance/t_leads.lead_CurrentValue) >= ";
			sql = sql + GetLTV();
			sql = sql + ")";
		}
		else
		{
			sql = sql + "\r\n AND (";
			sql = sql + " (t_leads.lead_1stBalance/t_leads.lead_CurrentValue) <= ";
			sql = sql + GetLTV();
			sql = sql + ")";
		}
	}
	//end filter area

	sql = sql + ")\r\n AS avail \r\n";
	sql = sql + " LEFT JOIN t_Processed_Orders ON avail.t_leads.lead_Id = t_Processed_Orders.PO_Lead_ID ";
	sql = sql + " GROUP BY  ";
	sql = sql + " avail.t_leads.lead_Id, avail.t_Leads.lead_FirstName, avail.t_Leads.lead_LastName, ";
	sql = sql + " avail.t_Leads.lead_CoApp, avail.t_Leads.lead_City, avail.t_Leads.lead_State,avail.t_Leads.lead_Address,avail.t_Leads.lead_Zip,avail.t_Leads.lead_WorkPhone,  ";
	sql = sql + " avail.t_Leads.lead_HomePhone, avail.t_Leads.lead_HouseType, avail.t_Leads.lead_CurrentValue,  avail.t_Leads.lead_DesiredLoan, avail.t_Leads.lead_1stBalance,  ";
	sql = sql + " avail.t_Leads.lead_Rate,  avail.t_Leads.lead_FixedAdjust, avail.t_Leads.lead_Payment,  avail.t_Leads.lead_Late, ";
	sql = sql + " avail.t_Leads.lead_Credit, avail.t_Leads.lead_WorkInfo, avail.t_Leads.lead_TimePeriod, avail.t_Leads.lead_Salary, avail.t_Leads.lead_YouShouldCall,  ";
	sql = sql + " avail.t_Leads.lead_LookingTo, avail.t_Leads.lead_Email,  avail.t_Leads.Sys_Cre_Date, ";
	sql = sql + " avail.t_Leads.lead_Verified, avail.t_Leads.lookingToID, avail.t_Leads.langInt, avail.t_Leads.VerifiedBy, avail.t_Leads.PersonalTouch  ";
	sql = sql + " \r\nORDER BY Count(po.PO_Company_ID) ";//DESC
	
	return sql;
}

void COrder::SetParentWindow(HWND m_hwnd)
{
	mainWin = m_hwnd;
}

void COrder::PrintOut(CString str)
{
	//myLog.WriteLog(str, GetLocationLong());
	if(mainWin)
	{
		SendMessage(mainWin, S_START_MESSAGE, 0, 0);
		int iLen = str.GetLength();
		for(int x = 0; x < iLen; x++)
		{
			int iLetter = (int) str.GetAt(x);
			SendMessage(mainWin, S_CHAR_MESSAGE, iLetter, 0);
		}

		SendMessage(mainWin, S_END_MESSAGE, 0, 0);
	}
}

void COrder::ReadLeads(void)
{
	if(pt_stop)
	{
			HRESULT hr = S_OK;
			try
			{
				//PrintOut("Init database connection");
				CoInitialize(NULL);
				// Define string variables.
				 
				_bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

			_RecordsetPtr pRst = NULL;

			// Call Create instance to instantiate the Record set
			hr = pRst.CreateInstance(__uuidof(Recordset));

			if(FAILED(hr))
			{
		            
				PrintOut("Failed creating record set instance");
				  
			}
			else
			{
				Progress(1, 100);
				CString sql;
				sql = GetSQL();
				PrintOut("Submit SQL" + sql);
				this->SendSQLString(IDC_EDIT_SQL, sql);
				pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockReadOnly, adCmdText);

				
				//Loop through the Record set
					if (!pRst->EndOfFile)
					{
						
						pRst->MoveFirst();
						
						//PrintOut("Start reading matching leads from file");
						long count = 0;
						CLead aLead;
						char buffer[200];
						//int max = iMaxNeeded;
						//if (max < iQuantity) max = iQuantity;
						int max = pRst->GetRecordCount();
						if(iMaxNeeded < max) max = iMaxNeeded;
						int x = 0;
						Progress(x, max);
						while((!pRst->EndOfFile) && (x < iMaxNeeded)&& (!*pt_stop))
						{				
							
							//long index = -1;
							aLead.SetParentWindow(mainWin);
							//PrintOut("Read Lead");
							aLead.ReadLead(pRst);
							if(AlreadyHas(aLead.GetLead_Id_String()))
							{
								CString str = "Already has lead" + aLead.GetLead_Id_String();
								str = str + " Not adding to matches";
								//PrintOut(str);
							}
							else
							{
								pt_LeadDBMap->Add(aLead);
								myMatchingLeadIDs.Add(aLead.GetLead_Id());
								count++;
								//sprintf(buffer, "%d", count);
								//sprintf(buffer, "%d", iMaxNeeded);
								//CString str = buffer;
								//SendCEditString(IDC_EDIT_YIELD, GetYield());
								//PrintOut(buffer);
								x++;
								Progress(x, max);
							}
							pRst->MoveNext();
						}			
						
						sprintf(buffer, "Finish reading %d leads", count);
						
						PrintOut(buffer);
						
					}
				}
				CoUninitialize();
			}
			catch(_com_error & ce)
			{
				CString str;
				
				str = "Error: " + GetCString(ce.Description());
				PrintOut(str);
				//myLog.WriteFail(str, GetLocationLong());
				
				
			}
			
		
	}
	else
	{
		PrintOut("COrder::ReadLeads: stop pointer not assigned");
	}
}

void COrder::PrintMatchingLeads(void)
{
	CLead aLead;
	PrintOut("Matching Leads not assigned");
	for(int x = 0; x < myMatchingLeadIDs.GetCount(); x++)
	{
		if(pt_LeadDBMap->GetLead(myMatchingLeadIDs.GetAt(x), aLead))
		{
			PrintOut(aLead.ToString());
		}		
	}	
}

void COrder::SetLeadDBMap(CLeadDB* m_LeadDBMap)
{
	pt_LeadDBMap = m_LeadDBMap;
}

void COrder::SetSFR(long iNum)
{
	if((iNum >=0) && (iNum <= 2))
	{
		iSFR = iNum;
	}
	else
	{
		iSFR = 0;
	}
}

CString COrder::GetSFR(void)
{
	return GetThreeStateStr(iSFR);
}

void COrder::SetMobile(long iNum)
{
	if((iNum >=0) && (iNum <= 2))
	{
		iMobile = iNum;
	}
	else
	{
		iMobile = 0;
	}
	
}

CString COrder::GetMobile(void)
{
	return GetThreeStateStr(iMobile);
}

void COrder::SetCommercial(long iNum)
{
	if((iNum >=0) && (iNum <= 2))
	{
		iCommercial = iNum;
	}
	else
	{
		iCommercial = 0;
	}
	
}

CString COrder::GetCommercial(void)
{
	return GetThreeStateStr(iCommercial);
}



void COrder::SetCity(CString str)
{
	sCity = str;
}

CString COrder::GetCity(void)
{
	return sCity;
}

void COrder::SetLTV(double fNum)
{
	if((fNum>=0) && (fNum <=1))
	{
		fLTV = fNum;
	}
	else
	{
		fLTV = 0;
	}
}

CString COrder::GetLTV(void)
{
	CString sNum;
	char buffer[200];
	sprintf( buffer, "%f", fLTV );
	sNum = buffer;
	return sNum;
}

void COrder::SetLTVAbove(bool state)
{
	bLTVAbove = state;
}

bool COrder::GetLTVAbove(void)
{
	return bLTVAbove;
}
COrder& COrder::operator=(const COrder& aOrder)
{
	this->bAdCampaign = aOrder.bAdCampaign;
	this->bLTVAbove = aOrder.bLTVAbove;
	this->bOneLeadPerPage = aOrder.bOneLeadPerPage;
	this->bOnePass = aOrder.bOnePass;
	this->bPrint = aOrder.bPrint;
	this->bPriority = aOrder.bPriority;
	this->bSendEmailNote = aOrder.bSendEmailNote;
	this->bSendExcel = aOrder.bSendExcel;
	this->bSendFax = aOrder.bSendFax;
	this->bSendMini1003 = aOrder.bSendMini1003;
	this->bSendWord = aOrder.bSendWord;
	this->bSubprime = aOrder.bSubprime;
	this->bVerified = aOrder.bVerified;
	this->fLTV = aOrder.fLTV;
	this->fRateFilter = aOrder.fRateFilter;
	this->i2ndMortgage = aOrder.i2ndMortgage;
	this->iAgentID = aOrder.iAgentID;
	this->iCashout = aOrder.iCashout;
	this->iCommercial = aOrder.iCommercial;
	this->iCompanyID = aOrder.iCompanyID;
	this->iCreditScore = aOrder.iCreditScore;
	this->iDebtCon = aOrder.iDebtCon;
	this->iHomeType = aOrder.iDebtCon;
	this->iLang = aOrder.iLang;
	this->iLocation = aOrder.iLocation;
	this->iLookingTo = aOrder.iLookingTo;
	this->iMobile = aOrder.iMobile;
	this->iOrderID = aOrder.iOrderID;
	this->iQuantity = aOrder.iQuantity;
	this->iSFR = aOrder.iSFR;
	this->mainWin = aOrder.mainWin;
	this->myMatchingLeadIDs.Copy(aOrder.myMatchingLeadIDs);
	this->myAssignedLeadIDs.Copy(aOrder.myAssignedLeadIDs);
	this->pt_LeadDBMap = aOrder.pt_LeadDBMap;
	this->sAgent = aOrder.sAgent;
	this->sAreacodeFilter = aOrder.sAreacodeFilter;
	this->sCity = aOrder.sCity;
	this->sCompanyName = aOrder.sCompanyName;
	this->sContact = aOrder.sContact;
	this->sEmail = aOrder.sEmail;
	this->sFaxNumber = aOrder.sFaxNumber;
	this->sStates = aOrder.sStates;
	this->sZipFilter = aOrder.sZipFilter;
	this->uliDesiredLoanFilter = aOrder.uliDesiredLoanFilter;
	this->iMaxNeeded = aOrder.iMaxNeeded;
	this->iLastOrderQuantity = aOrder.iLastOrderQuantity;
	this->iCompanyOrders = aOrder.iCompanyOrders;
	this->pt_stop = aOrder.pt_stop;
	this->iCampaignID = aOrder.iCampaignID;
	this->iMyUseLimit = aOrder.iMyUseLimit;
	this->iDefaultUseLimit = aOrder.iDefaultUseLimit;
	this->iMyAgeLimit = aOrder.iMyAgeLimit;
	this->iDefaultAgeLimit = aOrder.iDefaultAgeLimit;
	friends.Copy(aOrder.friends);
	

	return *this;
}
COrder::COrder(const COrder& aOrder)
{
	this->bAdCampaign = aOrder.bAdCampaign;
	this->bLTVAbove = aOrder.bLTVAbove;
	this->bOneLeadPerPage = aOrder.bOneLeadPerPage;
	this->bOnePass = aOrder.bOnePass;
	this->bPrint = aOrder.bPrint;
	this->bPriority = aOrder.bPriority;
	this->bSendEmailNote = aOrder.bSendEmailNote;
	this->bSendExcel = aOrder.bSendExcel;
	this->bSendFax = aOrder.bSendFax;
	this->bSendMini1003 = aOrder.bSendMini1003;
	this->bSendWord = aOrder.bSendWord;
	this->bSubprime = aOrder.bSubprime;
	this->bVerified = aOrder.bVerified;
	this->fLTV = aOrder.fLTV;
	this->fRateFilter = aOrder.fRateFilter;
	this->i2ndMortgage = aOrder.i2ndMortgage;
	this->iAgentID = aOrder.iAgentID;
	this->iCashout = aOrder.iCashout;
	this->iCommercial = aOrder.iCommercial;
	this->iCompanyID = aOrder.iCompanyID;
	this->iCreditScore = aOrder.iCreditScore;
	this->iDebtCon = aOrder.iDebtCon;
	this->iHomeType = aOrder.iDebtCon;
	this->iLang = aOrder.iLang;
	this->iLocation = aOrder.iLocation;
	this->iLookingTo = aOrder.iLookingTo;
	this->iMobile = aOrder.iMobile;
	this->iOrderID = aOrder.iOrderID;
	this->iQuantity = aOrder.iQuantity;
	this->iSFR = aOrder.iSFR;
	this->mainWin = aOrder.mainWin;
	this->myMatchingLeadIDs.Copy(aOrder.myMatchingLeadIDs);
	this->myAssignedLeadIDs.Copy(aOrder.myAssignedLeadIDs);
	this->pt_LeadDBMap = aOrder.pt_LeadDBMap;
	this->sAgent = aOrder.sAgent;
	this->sAreacodeFilter = aOrder.sAreacodeFilter;
	this->sCity = aOrder.sCity;
	this->sCompanyName = aOrder.sCompanyName;
	this->sContact = aOrder.sContact;
	this->sEmail = aOrder.sEmail;
	this->sFaxNumber = aOrder.sFaxNumber;
	this->sStates = aOrder.sStates;
	this->sZipFilter = aOrder.sZipFilter;
	this->uliDesiredLoanFilter = aOrder.uliDesiredLoanFilter;	
	this->iMaxNeeded = aOrder.iMaxNeeded;
	this->iLastOrderQuantity = aOrder.iLastOrderQuantity;
	this->iCompanyOrders = aOrder.iCompanyOrders;
	this->pt_stop = aOrder.pt_stop;
	this->iCampaignID = aOrder.iCampaignID;
	this->iMyUseLimit = aOrder.iMyUseLimit;
	this->iDefaultUseLimit = aOrder.iDefaultUseLimit;
	this->iMyAgeLimit = aOrder.iMyAgeLimit;
	this->iDefaultAgeLimit = aOrder.iDefaultAgeLimit;
	friends.Copy(aOrder.friends);

}

long COrder::GetOrderIDLong(void)
{
	return iOrderID;
}

bool COrder::AssignLead(void)
{
	if(pt_stop)
	{
		if(*pt_stop) return false;
		int randnum = rand();
		
		if(myMatchingLeadIDs.GetCount() > 0) RemoveFriendsLeads();
		bool success = false;
		if((myMatchingLeadIDs.GetCount() > 0) && (myAssignedLeadIDs.GetCount()< iQuantity))
		{
			
			//PrintOut("Rand assign a lead");
			
			
			CIDArray mins = GetMins(myMatchingLeadIDs);
			CLead minLead = pt_LeadDBMap->GetLead(mins[0]);
			long minuse = minLead.GetUses();
			if(minuse < GetUseLimit())
			{
				//CString str;
				//char buffer[65];
				long index = 0;
				if(mins.GetCount()>1)
				{
					index = randnum % (mins.GetCount()-1);
				}
				long id = mins[index];
				//_itoa(id, buffer, 10);
				//str = "Assigning Lead id: ";
				//str = str + buffer;
				//str = str + " to Order: ";
				//str = str + GetOrderID();
				//PrintOut(str);
				myAssignedLeadIDs.Add(id);
				//_itoa( myAssignedLeadIDs.GetCount(), buffer, 10 );
				//str = "Assigned so far: ";
				//str = str + buffer;
				//PrintOut(str);
				//CLead aLead = pt_LeadDBMap->GetLead(myMatchingLeadIDs.GetAt(index));
				CLead aLead = pt_LeadDBMap->GetLead(id);
				aLead.AddCompany(iCompanyID);
				pt_LeadDBMap->Update(aLead);
				//myMatchingLeadIDs.RemoveAt(index);
				success = true;	
			}
			else
			{
				char buffer[200];

				CString str, sMin, sMinLimit;
				sprintf(buffer, "%d", minuse);
				sMin = buffer;
				sprintf(buffer, "%d", GetUseLimit());
				sMinLimit = buffer;
				str = "Lead not assigned. The min use of the matching leads (" + sMin;
				str = str + ") is not less than the use limit (";
				str = str + sMinLimit;
				str = str + ") for this order";
				PrintOut(str);
			}

		}
		return success;
	}
	else
	{
		PrintOut("COrder::AssignLead: stop pointer not assigned");
		return false;
	}
}

void COrder::PrintAssignedLeads(void)
{
	if(pt_stop)
	{
		
		CLead aLead;
		char buffer[200];
		_itoa(myAssignedLeadIDs.GetCount(), buffer, 10);
		CString str = "Assigned Leads: ";
		str = str + buffer;
		PrintOut(str);	

		for(int x = 0; x < myAssignedLeadIDs.GetCount(); x++)
		{
			if(*pt_stop) break;
			//PrintOut("myAssignedLeadIDs count > 0");
			if(pt_LeadDBMap->GetLead(myAssignedLeadIDs.GetAt(x), aLead))
			{
				//PrintOut("Printing an assigned lead");
				PrintOut(aLead.ToString());
			}
			
		}
	}
	else
	{
		PrintOut("COrder::PrintAssignedLeads: stop pointer not assigned");
	}
}

long COrder::GetQuanityLong(void)
{
	return iQuantity;
}

void COrder::SetMaxNeeded(long iNum)
{
	iMaxNeeded = iNum;
}

void COrder::SortMatching(void)
{
	//QuickSort(0, myMatchingLeadIDs.GetCount()-1);
	long index = this->GetLastIndexOfMin();
	CString str = "Last index of min is: ";
	char buffer[200];
	_itoa(index, buffer, 10);
	str = str + buffer;
	PrintOut(str);

	
	
}

void COrder::QuickSort(long beg, long end)
{
	long temp;
   if (end > beg) 
   {
	   PrintOut("Sorting by lead uses");
      long piv = pt_LeadDBMap->GetLead(myMatchingLeadIDs[beg]).GetUses(), l = beg + 1, r = end;
      while (l < r) 
      {
		  
         if (pt_LeadDBMap->GetLead(myMatchingLeadIDs[l]).GetUses() <= piv) 
            l++;
         else if(pt_LeadDBMap->GetLead(myMatchingLeadIDs[r]).GetUses() >= piv)
            r--;
         else 
         {
            temp = myMatchingLeadIDs[l];
            myMatchingLeadIDs[l] = myMatchingLeadIDs[r];
            myMatchingLeadIDs[r] = temp;
         }
      }
      if(pt_LeadDBMap->GetLead(myMatchingLeadIDs[l]).GetUses() < piv)
      {
         temp = myMatchingLeadIDs[l];
         myMatchingLeadIDs[l] = myMatchingLeadIDs[beg];
         myMatchingLeadIDs[beg] = temp;
         l--;
      } 
      else
      {
         l--;
         temp = myMatchingLeadIDs[l];
         myMatchingLeadIDs[l] = myMatchingLeadIDs[beg];
         myMatchingLeadIDs[beg] = temp;
      }
      QuickSort(beg, l);
      QuickSort(r, end);
   }
}

long COrder::GetLastIndexOfMin(void)
{
	//myMatchingLeadIDs must be sorted accending by use prior to finding the index of the last min use
	
	if(pt_stop)
	{
		if(*pt_stop) return 0;
		myMatchingLeadIDs = FatQuickSort(myMatchingLeadIDs);
		PrintOut("Get the last index of min use");
		long index = -1;
		if(myMatchingLeadIDs.GetCount() > 0)
		{
			long min = pt_LeadDBMap->GetLead(myMatchingLeadIDs[0]).GetUses();
			
			for(long x = 0; x < myMatchingLeadIDs.GetCount(); x++)
			{
				if(*pt_stop) return 0;
				if(pt_LeadDBMap->GetLead(myMatchingLeadIDs[x]).GetUses() == min)
				{
					index = x;
				}
				else
				{
					break;
				}
			}
		}
		char buffer[200];
		_itoa(index, buffer, 10);
		CString str = "Last index of min: ";
		str = str + buffer;
		PrintOut(str);
		return index;
	}
	else
	{
		PrintOut("COrder::GetLastIndexOfMin: stop pointer not assigned");
		return 0;
	}
}

void COrder::AddFriend(long iCompID)
{
	if(iCompID != iCompanyID) 
	{
		//char buffer[200];
		//_itoa(iCompID, buffer, 10);
		//CString str = "Adding friend: ";
		//str = str + buffer;
		//PrintOut(str);
		friends.Add(iCompID);
	}
}

void COrder::RemoveFriendsLeads(void)
{
	//we only need to remove friends leads if we have friends
	if(friends.GetCount()>0)
	{
		//PrintOut("Remove friends leads");
		//CArray<long,long> cleaned;
		CIDArray cleaned;

		CLead aLead;
		for(int y = 0; y < myMatchingLeadIDs.GetCount(); y++)
		{
			aLead = pt_LeadDBMap->GetLead(myMatchingLeadIDs[y]);
			bool found = false;
			for(int x=0; x < friends.GetCount(); x++)
			{			
				if(aLead.HasCompany(friends[x])) found = true;							
			}
			if(!found) cleaned.Add(aLead.GetLead_Id());
		}
		myMatchingLeadIDs.RemoveAll();
		myMatchingLeadIDs.Copy(cleaned);
	}
	
}

void COrder::ReadFriends(void)
{
	HRESULT hr = S_OK;
    try
    {
		//PrintOut("Init database connection");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

       _RecordsetPtr pRstFriends = NULL;

      // Call Create instance to instantiate the Record set
      hr = pRstFriends.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {
            
		  PrintOut("Failed creating record set instance");
		  
      }
	  else
	  {
		
		CString sql;
		sql = "SELECT * FROM friends WHERE friends.cust_id=";
		sql = sql + GetCompanyID();
		//PrintOut("Submit SQL");
		pRstFriends->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockReadOnly, adCmdText);		
		
		//Loop through the Record set
		if (!pRstFriends->EndOfFile)
		{			
			pRstFriends->MoveFirst();			
			//PrintOut("Start reading friends from file");			
			while(!pRstFriends->EndOfFile)
			{	
				AddFriend(GetFieldLong(pRstFriends, "friend_id"));
				pRstFriends->MoveNext();
			}			
			
			
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
		//myLog.WriteFail(str, GetLocationLong());
		
		
	}
	
  CoUninitialize();
}

CIDArray COrder::FatQuickSort(CIDArray ids)
{
	if(pt_stop)
	{
		if(ids.GetCount()<= 1) return ids;
		else
		{
			CIDArray less, equal, greater, result;
			long pivot = ids.GetCount()/2;
			for(int x = 0; x < ids.GetCount(); x++)
			{
				if(*pt_stop) return ids;
				if((pt_LeadDBMap->GetLead(ids[x]).GetUses()) < (pt_LeadDBMap->GetLead(ids[pivot]).GetUses()))
				{
					less.Add(ids[x]);
				}
				else if ((pt_LeadDBMap->GetLead(ids[x]).GetUses()) == (pt_LeadDBMap->GetLead(ids[pivot]).GetUses()))
				{
					equal.Add(ids[x]);
				}
				else if ((pt_LeadDBMap->GetLead(ids[x]).GetUses()) > (pt_LeadDBMap->GetLead(ids[pivot]).GetUses()))
				{
					greater.Add(ids[x]);
				}
			}
			result = FatQuickSort(less);
			result.Append(equal);
			result.Append(FatQuickSort(greater));
			return result;
		}
	}
	else
	{
		PrintOut("COrder::FatQuickSort: stop pointer not assigned");
		return ids;
	}
}

long COrder::GetCountMatching(void)
{
	return myMatchingLeadIDs.GetCount();
}

void COrder::SetLastOrderQuantity(long iNum)
{
	iLastOrderQuantity = iNum;
}

CString COrder::GetLastOrderQuantity(void)
{
	CString sNum;
	char buffer[200];
	sprintf( buffer, "%d", iLastOrderQuantity );
	sNum = buffer;
	return sNum;
}

void COrder::SetCompanyOrders(long iNum)
{
	iCompanyOrders = iNum;
}

CString COrder::GetCompanyOrders(void)
{
	CString sNum;
	char buffer[200];
	sprintf( buffer, "%d", iCompanyOrders );
	sNum = buffer;
	return sNum;
}

CString COrder::GetYield(void)
{
	CString sNum;
	char buffer[200];
	double fYield = GetYieldDouble();
	sprintf( buffer, "%f", fYield );
	sNum = buffer;
	return sNum;
}

double COrder::GetYieldDouble(void)
{
	double fYield = 0;
	if(iMaxNeeded > 0)
	{
		fYield = (double)myMatchingLeadIDs.GetCount()/(double)iMaxNeeded;
		
	}
	return fYield;
	
}

void COrder::SendCEditString(int id, CString str)
{
	if(mainWin)
	{
		SendMessage(mainWin, S_ORDER_TRANS, id, 0);
		int iLen = str.GetLength();
		for(int x = 0; x < iLen; x++)
		{
			int iLetter = (int) str.GetAt(x);
			SendMessage(mainWin, S_ORDER_TRANS, id, iLetter);
		}

		
	}
}

void COrder::SetStop(bool* pt_state)
{
	pt_stop = pt_state;
	
}

CIDArray COrder::GetMins(CIDArray ids)
{
	if(pt_stop)
	{
		if(ids.GetCount()>0)
		{
			CIDArray mins;
			long min = pt_LeadDBMap->GetLead(ids[0]).GetUses();
			for(long x = 0; x< ids.GetCount(); x++)
			{
				if(*pt_stop) return ids;// user stopped process, halt all work
				if(pt_LeadDBMap->GetLead(ids[x]).GetUses()< min)
				{
					min = pt_LeadDBMap->GetLead(ids[x]).GetUses();
					mins.RemoveAll();
					mins.Add(ids[x]);
				}
				else if(pt_LeadDBMap->GetLead(ids[x]).GetUses()== min)
				{
					mins.Add(ids[x]);
				}
			}
			return mins;
		}
		else
		{
			return ids;
		}
	}
	else
	{
		PrintOut("COrder::GetMins:Error: stop pointer not assigned");
		return ids;
	}
	
}

long COrder::GetCompanyIDLong(void)
{
	return iCompanyID;
}

CIDArray COrder::GetAssignedLeadIDs(void)
{
	return this->myAssignedLeadIDs;
}

void COrder::ReadAssignedLeads()
{
	//PrintOut("Read assigned leads");
	HRESULT hr = S_OK;
    try
    {
		//PrintOut("Init database connection");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

       _RecordsetPtr pRst = NULL;

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {
            
		  PrintOut("Failed creating record set instance");
		  
      }
	  else
	  {
		
		CString sql = "";
		sql = "SELECT BPRetrySendOrder.Ord_Id, t_Leads.* ";
		sql = sql + "FROM ([SELECT t_Orders.* ";
		sql = sql + "FROM t_Orders ";
		sql = sql + "WHERE (((t_Orders.ReadyToSend)=True) ";
		sql = sql + "AND ((t_Orders.Ord_Complete)=False)) ";
		sql = sql + "]. AS BPRetrySendOrder LEFT JOIN t_Processed_Orders ON BPRetrySendOrder.Ord_Id = t_Processed_Orders.PO_Ord_ID) LEFT JOIN t_Leads ON t_Processed_Orders.PO_Lead_ID = t_Leads.lead_Id ";
		sql = sql + "WHERE (((BPRetrySendOrder.Ord_Id)=";
		sql = sql + this->GetOrderID() + ")) AND  NOT ISNULL(lead_Id)";

		SendSQLString(IDC_EDIT_SQL, sql);
		
		PrintOut(sql);
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);
		int count = 0;
		int x = 0;
		int max = pRst->GetRecordCount();
		if(pRst->GetRecordCount()>0)
		{
			pRst->MoveFirst();
			while(!pRst->EndOfFile)
		{				
			CLead aLead;
			//long index = -1;
			aLead.SetParentWindow(mainWin);
			//PrintOut("Read Lead");
			aLead.ReadLead(pRst);
			
			pt_LeadDBMap->Add(aLead);
			myAssignedLeadIDs.Add(aLead.GetLead_Id());
			count++;
			
			x++;
			Progress(x, max);
			
			pRst->MoveNext();
		}			
		}
		
		char buffer[200];
		sprintf(buffer, "Finish reading %d leads", count);
		
		PrintOut(buffer);
	  }
	  CoUninitialize();
	  
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "Error ReadAssignedLeads : " + GetCString(ce.Description());
		PrintOut(str);	
		//myLog.WriteFail(str, GetLocationLong());
	}
	
  
}

void COrder::AddAssignedLeadID(long id)
{
	myAssignedLeadIDs.Add(id);
	CIDArray temp;
	for(long x = 0; x < myMatchingLeadIDs.GetCount(); x++)
	{
		if(myMatchingLeadIDs[x] != id) temp.Add(myMatchingLeadIDs[x]);
	}
	myMatchingLeadIDs = temp;
}

void COrder::WriteRTF(void)
{
	CFileException e;
	CFile *m_File = new CFile();
	CString path = ".\\";
	path = path + GetOrderID();
	path = path + ".rtf";

	
	
	if(!m_File->Open(path.GetBuffer(path.GetLength()), CFile::modeCreate | CFile::modeReadWrite, &e))
	{
		PrintOut("Could not open the questions file for writing at path: " + path);
		
	}
	else
	{
		PrintOut("Opened the rtf file for writing at path: " + path);
		//not done yet come back to here
		
	}
}

CLeadDB* COrder::GetLeadDBMap(void)
{
	return pt_LeadDBMap;
}

long COrder::GetLocationLong(void)
{
	return iLocation;
}

long COrder::GetCountAssigned(void)
{
	return myAssignedLeadIDs.GetCount();
}

CFaxJob COrder::GetFaxJob(CLocation myloc)
{
	CFaxJob aJob;
	aJob.SetAgentEmail(GetAgentEmail());
	aJob.SetAgentname(GetAgent());
	aJob.SetCampaignID(this->GetCampaignID());
	aJob.SetCompany(this->GetCompanyName());
	aJob.SetContact(this->GetContact());
	aJob.SetDiv(this->GetLocationLong());
	aJob.SetFaxnum(this->GetFaxNumber());
	aJob.SetStates(this->GetStates());
	CString path = myloc.GetDesDir() + "\\";
	path = path + GetOrderID();
	
	
	aJob.SetOrderID(this->GetOrderIDLong());
	if(this->GetCountAssigned() > 0)
	{
		aJob.SetSendCoverpage(1);
		path = path + ".rtf";
	}
	else
	{
		aJob.SetSendCoverpage(0);// no leads get a zero letter with no cover
		path = path + ".txt";
	}
	aJob.SetFilename(path);
	

	return aJob;
}

CString COrder::GetAgentEmail(void)
{
	CString email;
	if(pt_stop)
	{
			HRESULT hr = S_OK;
			try
			{
				//PrintOut("Init database connection");
				CoInitialize(NULL);
				// Define string variables.
				 
				_bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

			_RecordsetPtr pRs = NULL;

			// Call Create instance to instantiate the Record set
			hr = pRs.CreateInstance(__uuidof(Recordset));

			if(FAILED(hr))
			{
		            
				PrintOut("Failed creating record set instance");
				  
			}
			else
			{
				
				CString sql;
				sql = "SELECT * FROM t_Agents WHERE t_Agents.Agent_Num=" + this->GetAgentID();
				//PrintOut("Submit SQL:" + sql);
				pRs->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockReadOnly, adCmdText);

				int x = 0;
				//Loop through the Record set
				if (!pRs->EndOfFile)
				{					
					pRs->MoveFirst();					
					email = GetFieldString(pRs, "email");
					
				}
			}
			  
			}
			catch(_com_error & ce)
			{
				CString str;
				
				str = "COrder::GetAgentEmail: " + GetCString(ce.Description());
				PrintOut(str);
				//myLog.WriteFail(str, GetLocationLong());
				
				
			}
			
		CoUninitialize();
	}
	else
	{
		PrintOut("COrder::GetAgentEmail: stop pointer not assigned");
	}
	return email;
}

void COrder::SetCampaignID(long iID)
{
	iCampaignID = iID;
}

long COrder::GetCampaignID(void)
{
	return iCampaignID;
}

CString COrder::GetCampaignIDString(void)
{
	CString id;
	char buffer[200];
	sprintf( buffer, "%d", iCampaignID );
	id = buffer;
	return id;
}

int COrder::GetFilterCount(void)
{ 
	int count = 0;
	if(fRateFilter>0) count++;		
	if(uliDesiredLoanFilter>0) count++;	
	if(sZipFilter.GetLength()>0) count++;	
	if(sAreacodeFilter.GetLength()>0) count++;	
	if((iLookingTo > 0) && (iLookingTo <= 2)) count++;	
	if((iLang > 0) && (iLang <= 2)) count++;	
	if(iCreditScore > 0) count++;	
	if(iCashout > 0) count++;	
	if(i2ndMortgage > 0) count++;		
	if(iDebtCon > 0) count++;				
	if(iHomeType > 0) count++;	
	if(iSFR > 0) count++;	
	if(iCommercial > 0)	count++;	
	if(iMobile > 0) count++;		
	if(sCity.GetLength() > 0) count++;	
	if(fLTV > 0) count++;	

	return count;
}
CString COrder::GetEmailNote(CLocation myloc)
{
	CString sResult;
	CIDArray leadIDs = GetAssignedLeadIDs();
	CLeadDB* pt_leadDB = GetLeadDBMap();
	int pagecount = 1;
	for(int x = 0; x < leadIDs.GetCount(); x++)
	{
		CLead aLead = pt_leadDB->GetLead(leadIDs[x]);		
		CString aPage;
		
		if(GetVerified())
		{
			aPage = myloc.GetEmailNoteVerifiedPage();
		}
		else
		{
			aPage = myloc.GetEmailNoteNonVerifiedPage();
		}
		
		COleDateTime aDate;
		aDate = COleDateTime::GetCurrentTime();
		CString sDate = aDate.Format(_T("%m/%d/%y"));
		aPage.Replace("<date>", sDate); 
		if(GetVerified())
		{
			aPage.Replace("<title>", "Verified Application"); 
		}
		else if(GetOnePass())
		{
			aPage.Replace("<title>", "One-Pass Application"); 
		}
		else
		{
			aPage.Replace("<title>", "Internet Application"); 
		}
		if(aLead.GetLead_Id()> 0)
		{
			aPage.Replace("<id>", aLead.GetLead_Id_String()); 
		}
		else
		{
			aPage.Replace("<id>", ""); 
		}
		CString name = aLead.GetLead_FirstName() + " ";
		name = name + aLead.GetLead_LastName();
		aPage.Replace("<name>", name);
		aPage.Replace("<coapp>", aLead.GetLead_CoApp());
		aPage.Replace("<city>", aLead.GetLead_City());
		aPage.Replace("<state>", aLead.GetLead_State());
		aPage.Replace("<address>", aLead.GetLead_Address());
		aPage.Replace("<zip>", aLead.GetLead_Zip());
		
		if(aLead.IsPurchase())
		{
			aPage.Replace("<subcity>", "");
			aPage.Replace("<substate>", "");
			aPage.Replace("<subaddress>", "Purchase");
			aPage.Replace("<subzip>", "");
		}
		else
		{
			aPage.Replace("<subcity>", aLead.GetLead_City() + ",");
			aPage.Replace("<substate>", aLead.GetLead_State());
			aPage.Replace("<subaddress>", aLead.GetLead_Address() + ",");
			aPage.Replace("<subzip>", aLead.GetLead_Zip());

		}

		aPage.Replace("<workphone>", CFormatter::FormatPhone(aLead.GetLead_WorkPhone()));
		aPage.Replace("<homephone>", CFormatter::FormatPhone(aLead.GetLead_HomePhone()));
		aPage.Replace("<housetype>", aLead.GetLead_HouseType());
		if(aLead.GetLead_CurrentValue()>0)
		{
			if(aLead.IsPurchase())
			{
				PrintOut("is a purchase");
				aPage.Replace("<currentvalue>", "");
			}
			else
			{
				aPage.Replace("<currentvalue>", CFormatter::FormatMoney(aLead.GetLead_CurrentValue_String()));
			}
		}
		else
		{
			aPage.Replace("<currentvalue>", "");
		}
		if(aLead.GetLead_DesiredLoan()>0)
		{
			aPage.Replace("<desiredloan>", CFormatter::FormatMoney(aLead.GetLead_DesiredLoan_String()));
		}
		else
		{
			aPage.Replace("<desiredloan>", "");
		}
		if(aLead.GetLead_1stBalance()>0)
		{
			if(aLead.IsPurchase())
			{
				aPage.Replace("<1stbal>", "");
			}
			else
			{
				aPage.Replace("<1stbal>", CFormatter::FormatMoney(aLead.GetLead_1stBalance_String()));
			}
		}
		else
		{
			aPage.Replace("<1stbal>", "");
		}
		aPage.Replace("<rate>", aLead.GetLead_Rate_String());
		aPage.Replace("<fixadj>", aLead.GetLead_FixedAdjust());			
		
		aPage.Replace("<payment>", CFormatter::FormatMoney(aLead.GetLead_Payment()));			
		
		aPage.Replace("<late>", aLead.GetLead_Late());
		aPage.Replace("<credit>", aLead.GetLead_Credit());
		aPage.Replace("<workinfo>", aLead.GetLead_WorkInfo());
		aPage.Replace("<timeperiod>", aLead.GetLead_TimePeriod());
		//if(GetVerified())
		//{
			//aPage.Replace("<verifyby>", aLead.GetVerifiedBy());
		//}
		//else
		//{
			//aPage.Replace("<verifyby>", "");
		//}
		if(aLead.GetLead_Salary()>0)
		{
			aPage.Replace("<salary>", CFormatter::FormatMoney(aLead.GetLead_Salary_String()));
		}
		else
		{
			aPage.Replace("<salary>", "");
		}
		aPage.Replace("<youcall>", aLead.GetLead_YouShouldCall());
		aPage.Replace("<lookingto>", aLead.GetLead_LookingTo());
		aPage.Replace("<email>", aLead.GetLead_Email());
		aPage.Replace("<personaltouch>", aLead.GetPersonalTouch());		
		aPage.Replace("<orderid>", GetOrderID());
		char buffer[200];			
		sprintf(buffer, "%d", pagecount);
		CString pagenum = buffer;
		aPage.Replace("<pg#>", pagenum);
		aPage.Replace("<contact>", GetContact());
		aPage.Replace("<company>", GetCompanyName());
		pagecount++;

		sResult = sResult + aPage;		

	}
	return sResult;
}



CString COrder::GetZeroBody(CLocation myLoc)
{
	CString body = myLoc.GetZeroBody();
	body.Replace("<agentname>", GetAgent());	
	body.Replace("<agentemail>", GetAgentEmail());
	body.Replace("<orderid>", GetOrderID());
	body.Replace("<company>", GetCompanyName());
	body.Replace("<contact>", GetContact());
	body.Replace("<state>", GetStates());
	if(GetFilterCount() > 0)
	{
		body.Replace("<filterstatement>", myLoc.GetFilterStatement());
	}
	else
	{
		body.Replace("<filterstatement>", "");
	}
	return body;
}

void COrder::SetMyUseLimit(long iNum)
{
	iMyUseLimit = iNum;
}

void COrder::SetDefaultUseLimit(long iNum)
{
	iDefaultUseLimit = iNum;
}

long COrder::GetUseLimit(void)
{
	if(iMyUseLimit>0)
	{
		return iMyUseLimit;
	}
	else
	{
		return iDefaultUseLimit;
	}	
}

void COrder::SetMyAgeLimit(long iNum)
{
	iMyAgeLimit = iNum;
}

void COrder::SetDefaultAgeLimit(long iNum)
{
	iDefaultAgeLimit = iNum;
}

CString COrder::GetAgeLimit(void)
{
	CString sNum;
	char buffer[200];
	
	if(iMyAgeLimit > 0)
	{
		sprintf(buffer, "%d", iMyAgeLimit);
		
	}
	else
	{
		sprintf(buffer, "%d", iDefaultAgeLimit);
	}
	sNum = buffer;

	return sNum;
}
void COrder::Progress(int iPos, int iMax)
{
	if(mainWin)
	{
		SendMessage(mainWin, ID_PROGRESS, iPos, iMax);
	}
}
bool COrder::AlreadyHas(CString sLeadID)
{
	return false;//disable the dupe check for testing of new GetSQLNew()
	bool foundit = true;
	HRESULT hr = S_OK;
	try
	{
		//PrintOut("Init database connection");
		CoInitialize(NULL);
		// Define string variables.
			
		_bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

		_RecordsetPtr pRst = NULL;

		// Call Create instance to instantiate the Record set
		hr = pRst.CreateInstance(__uuidof(Recordset));

		if(FAILED(hr))
		{
			    
			PrintOut("Failed creating record set instance");
				
		}
		else
		{
			
			CString sql;
			sql = "SELECT * FROM t_Processed_Orders WHERE PO_Lead_ID=";
			sql = sql + sLeadID;
			sql = sql + " AND (";
			sql = sql + "PO_Company_ID=";
			sql = sql + GetCompanyID();
			for(int x = 0; x < friends.GetCount();x++)
			{
				char buffer[200];
				long id = friends[x];
				sprintf(buffer, "%d", id);
				CString sID = buffer;
				sql = sql + " OR PO_Company_ID=";
				sql = sql + sID;				
				
			}
			
			sql = sql + ")";
			//PrintOut("Dupe check:" + sql);
			
			pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockReadOnly, adCmdText);
			if(pRst->GetRecordCount()>0)
			{
				foundit = true;
				//PrintOut("Dupe");
			}
			else
			{
				foundit = false;
				//PrintOut("Not Dupe");
			}			
			
		}
		
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Error: " + GetCString(ce.Description());
		PrintOut(str);
		//myLog.WriteFail(str, GetLocationLong());
		
		
	}
	
	CoUninitialize();
	return foundit;
}

CString COrder::GetSQLNew(void)
{
	CString sql;
	sql = "SELECT BPCountValidUses.Used, BPLeadsDontHave.* \r\n";
	sql = sql + "FROM \r\n";
	sql = sql + "((\r\n";
	sql = sql + "SELECT t_Leads.*, BPfriendsLeads.PO_Lead_ID \r\n";
	sql = sql + "FROM t_Leads LEFT JOIN \r\n";
	sql = sql + "( \r\n";
	sql = sql + "SELECT t_Processed_Orders.PO_Lead_ID, t_Processed_Orders.PO_Company_ID, friends.cust_id \r\n";
	sql = sql + "FROM friends INNER JOIN t_Processed_Orders ON friends.friend_id = t_Processed_Orders.PO_Company_ID \r\n";
	sql = sql + "WHERE (((friends.cust_id)=";
	sql = sql + GetCompanyID();
	sql = sql + "))\r\n";
	sql = sql + ") AS  BPfriendsLeads ";
	sql = sql + "ON t_Leads.lead_Id = BPfriendsLeads.PO_Lead_ID \r\n";
	sql = sql + "WHERE (((BPfriendsLeads.PO_Lead_ID) Is Null)) \r\n";
	sql = sql + ") AS BPLeadsDontHave) \r\n";
	sql = sql + "INNER JOIN \r\n";
	sql = sql + "(\r\n";
	sql = sql + "SELECT t_Leads.lead_Id, Count(BPvaliduse.PO_Company_ID) AS Used \r\n";
	sql = sql + "FROM t_Leads LEFT JOIN \r\n";
	sql = sql + "(\r\n";
	sql = sql + "SELECT t_Processed_Orders.PO_Lead_ID, t_Processed_Orders.PO_Company_ID \r\n";
	sql = sql + "FROM t_Leads LEFT JOIN t_Processed_Orders ON t_Leads.lead_Id = t_Processed_Orders.PO_Lead_ID \r\n";
	sql = sql + "WHERE (((t_Processed_Orders.PO_Lead_SentOn)>=t_Leads.Sys_Cre_Date)) \r\n";
	sql = sql + ") AS BPvaliduse \r\n"; 
	sql = sql + "ON t_Leads.lead_Id = BPvaliduse.PO_Lead_ID \r\n";
	sql = sql + "GROUP BY t_Leads.lead_Id \r\n";
	sql = sql + ") AS BPCountValidUses \r\n";
	//BPCountValidUses \r\n";
	sql = sql + "ON BPLeadsDontHave.lead_Id=BPCountValidUses.lead_Id \r\n";
	sql = sql + "WHERE (((BPLeadsDontHave.dupe)=False) \r\n"; 
	sql = sql + "AND ((BPLeadsDontHave.lead_DoNotUse)=False) \r\n";
	//start filter area
	sql = sql + "AND (BPLeadsDontHave.lead_Verified=";
	sql = sql + BooltoStr(GetVerified());	
	sql = sql + ")\r\n AND ";
	if(bOnePass)
	{
		sql = sql + "DateDiff('d',BPLeadsDontHave.Sys_Cre_Date,Now())>";
		sql = sql + GetAgeLimit();
	}
	else
	{
		
		sql = sql + "DateDiff('d',BPLeadsDontHave.Sys_Cre_Date,Now())<=";
		sql = sql + GetAgeLimit();
		sql = sql + "\r\n AND ";
		sql = sql + "BPLeadsDontHave.Sys_Cre_Date <= Now() ";
	}
	
	if(fRateFilter>0)
	{
		sql = sql + "\r\n AND BPLeadsDontHave.lead_Rate>=";
		sql = sql + GetRateFilter();
	}	
	if(uliDesiredLoanFilter>0)
	{
		sql = sql + "\r\n AND BPLeadsDontHave.lead_DesiredLoan>=";
		sql = sql + this->GetDesiredLoanFilter();
	}	
	if(sStates.GetLength()>0)
	{
		CString statesSQL = sStates;
		statesSQL.Replace(" ", ""); //remove spaces
		sql = sql + "\r\n AND (BPLeadsDontHave.lead_State='";
		statesSQL.Replace(",", "' OR BPLeadsDontHave.lead_State='");
		sql = sql + statesSQL;
		sql = sql + "')";
	}
	if(sZipFilter.GetLength()>0)	
	{
		CString zipSQL = sZipFilter;
		zipSQL.Replace(" ", ""); //remove spaces
		sql = sql + "\r\n AND (BPLeadsDontHave.lead_Zip LIKE '";
		zipSQL.Replace(",", "%' OR BPLeadsDontHave.lead_Zip LIKE '");
		sql = sql + zipSQL;
		sql = sql + "%')";
	}
	if(sAreacodeFilter.GetLength()>0)	
	{
		sAreacodeFilter.Replace(" ", ""); //remove spaces
		sql = sql + "\r\n AND (";
		CString homeSQL = "\r\n(BPLeadsDontHave.lead_HomePhone LIKE '" + sAreacodeFilter;
		CString workSQL =  "\r\n(BPLeadsDontHave.lead_WorkPhone LIKE '" + sAreacodeFilter;		
		homeSQL.Replace(",", "%') \r\n OR (BPLeadsDontHave.lead_HomePhone LIKE '");
		workSQL.Replace(",", "%') \r\n OR (BPLeadsDontHave.lead_WorkPhone LIKE '");
		homeSQL = homeSQL + "%')";
		workSQL = workSQL + "%')";
		sql = sql + homeSQL;
		sql = sql + "\r\n OR ";
		sql = sql + workSQL;
		sql = sql + ")";
	}
	if((iLookingTo > 0) && (iLookingTo <= 2))
	{
		sql = sql + "\r\n AND ";
		if(iLookingTo == REFINANCE) sql = sql + " BPLeadsDontHave.lead_LookingTo NOT LIKE 'purch%' ";
		if(iLookingTo == PURCHASE) sql = sql + " BPLeadsDontHave.lead_LookingTo LIKE 'purch%' ";

	}
	if((iLang > 0) && (iLang <= 2))
	{
		sql = sql + "\r\n AND ";
		if(iLang == ENGLISH) sql = sql + " BPLeadsDontHave.langInt<2 ";			
		
		if(iLang == SPANISH) sql = sql + " BPLeadsDontHave.langInt=2 ";
	}
	if(iCreditScore > 0)
	{
		
		sql = sql + "\r\n AND (";
		if(bSubprime)
		{
			
			switch ( iCreditScore )
			{
				case EXCELLENT:				
					sql = sql + " (BPLeadsDontHave.lead_Credit = 'Excellent') OR ";
					
				case GOOD:
					sql = sql + " (BPLeadsDontHave.lead_Credit = 'Good') OR ";
				
				case FAIR:
					sql = sql + " (BPLeadsDontHave.lead_Credit = 'Fair') OR ";
				
				case POOR:
					sql = sql + " (BPLeadsDontHave.lead_Credit = 'Poor') OR";
					sql = sql + " (BPLeadsDontHave.lead_Credit = 'Major_Credit_Issues') ";
				break;
				default:
					//invalid iCreditScore add nothing to the SQL
				break;
	            
			}

		}
		else
		{
			switch ( iCreditScore )
			{
				case POOR:
					sql = sql + " (BPLeadsDontHave.lead_Credit = 'Poor') OR";
					sql = sql + " (BPLeadsDontHave.lead_Credit = 'Major_Credit_Issues') OR";
					
				case FAIR:
					sql = sql + " (BPLeadsDontHave.lead_Credit = 'Fair') OR ";
				
				case GOOD:
					sql = sql + " (BPLeadsDontHave.lead_Credit = 'Good') OR ";
				
				case EXCELLENT:
					sql = sql + " (BPLeadsDontHave.lead_Credit = 'Excellent') ";
				break;
				default:
					//invalid iCreditScore add nothing to the SQL
				break;
			}
		}
		sql = sql + ")";
	}
	if((iCashout==1) || (i2ndMortgage==1) || (iDebtCon==1))
	{
		sql = sql + "\r\n AND (";
		
		if(iCashout==1)
		{
			sql = sql + " (BPLeadsDontHave.lead_LookingTo LIKE '%cash%') ";
			if((i2ndMortgage==1) || (iDebtCon==1))
			{
				sql = sql + " OR ";
			}
		}				
		if(i2ndMortgage==1)
		{
			sql = sql + " (BPLeadsDontHave.lead_LookingTo LIKE '%Second%') ";
			if(iDebtCon == 1)
			{
				sql = sql + " OR ";
			}
		}				
		if(iDebtCon==1)
		{
			sql = sql + " (BPLeadsDontHave.lead_LookingTo LIKE '%Consol%') ";
		}	

		sql = sql + ")";
	}					
	if(iCashout==2)
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (BPLeadsDontHave.lead_LookingTo NOT LIKE '%cash%') ";	
		sql = sql + ")";
	}
	if(i2ndMortgage==2)				
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (BPLeadsDontHave.lead_LookingTo NOT LIKE '%Second%') ";
		sql = sql + ")";				
	}
	if(iDebtCon==2)				
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (BPLeadsDontHave.lead_LookingTo NOT LIKE '%Consol%') ";
		sql = sql + ")";				
	}			
	if(iHomeType > 0)
	{
		sql = sql + "\r\n AND (";
		

		switch ( iHomeType )//not finished more needed
			{
				case SFR:
					sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE '%Single Family%') OR";
					sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE 'Condo%') OR";
					sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE 'Town%') OR";
					sql = sql + " (BPLeadsDontHave.lead_HouseType = 'SFR') ";
					
					break;
					
				case COMMERCIAL:
					sql = sql + " (BPLeadsDontHave.lead_HouseType = 'Commercial') OR";
					sql = sql + " (BPLeadsDontHave.lead_HouseType = 'commercial') ";
					break;
				
				case MOBILE:
					sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE '%Mobile%') OR";	
					sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE '%wide%') OR";
					sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE '%Manufac%') ";	
					break;

				default:
					//invalid iHomeType add nothing to the SQL
				break;
			}
		sql = sql + ")";

	}
	if((iSFR==1) || (iCommercial==1) || (iMobile==1))
	{
		sql = sql + "\r\n AND (";
		
		if(iSFR==1)
		{
			sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE '%Single%Family%') OR";
			sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE 'Condo%') OR";
			sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE 'Town%') OR";
			sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE 'SFR') ";

			if((iCommercial==1) || (iMobile==1))
			{
				sql = sql + " OR ";
			}
		}				
		if(iCommercial==1)
		{
			sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE '%commerc%') ";
			
			if(iMobile == 1)
			{
				sql = sql + " OR ";
			}
		}				
		if(iMobile==1)
		{
			sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE '%Mobile%') OR";	
			sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE '%wide%') OR";
			sql = sql + " (BPLeadsDontHave.lead_HouseType LIKE '%Manufac%') ";	
		}	

		sql = sql + ")";
	}					
	if(iSFR==2)
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (BPLeadsDontHave.lead_HouseType NOT LIKE '%Single%Family%') AND";
			sql = sql + " (BPLeadsDontHave.lead_HouseType NOT LIKE 'Condo%') AND";
			sql = sql + " (BPLeadsDontHave.lead_HouseType NOT LIKE 'Town%') AND";
			sql = sql + " (BPLeadsDontHave.lead_HouseType NOT LIKE 'SFR') ";

		sql = sql + ")";
	}
	if(iCommercial==2)				
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (BPLeadsDontHave.lead_HouseType NOT LIKE '%commercial%') ";
			
		sql = sql + ")";				
	}
	if(iMobile==2)				
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (BPLeadsDontHave.lead_HouseType NOT LIKE '%Mobile%') AND";	
		sql = sql + " (BPLeadsDontHave.lead_HouseType NOT LIKE '%wide%') AND";
		sql = sql + " (BPLeadsDontHave.lead_HouseType NOT LIKE '%Manufac%') ";	
		sql = sql + ")";				
	}			
	if(sCity.GetLength()>0)
	{
		sql = sql + "\r\n AND (";
		sql = sql + " (BPLeadsDontHave.lead_City LIKE '%";
		sql = sql + GetCity();
		sql = sql + "%') ";
		sql = sql + ")";

	}
	if(fLTV > 0)
	{
		if(bLTVAbove)
		{
			sql = sql + "\r\n AND (";
			sql = sql + " (BPLeadsDontHave.lead_1stBalance/BPLeadsDontHave.lead_CurrentValue) >= ";
			sql = sql + GetLTV();
			sql = sql + ")";
		}
		else
		{
			sql = sql + "\r\n AND (";
			sql = sql + " (BPLeadsDontHave.lead_1stBalance/BPLeadsDontHave.lead_CurrentValue) <= ";
			sql = sql + GetLTV();
			sql = sql + ")";
		}
	}
	//end filter area
	//sql = sql + "AND ((BPLeadsDontHave.lead_Verified)=True)
	sql = sql + ") \r\n";
	sql = sql + "ORDER BY BPCountValidUses.Used \r\n";
	sql = sql + ", BPLeadsDontHave.Sys_Cre_Date DESC";
	return sql;
}

void COrder::DeleteAssignedLeads(void)
{
	//PrintOut("Read assigned leads");
	HRESULT hr = S_OK;
    try
    {
		//PrintOut("Init database connection");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

       _RecordsetPtr pRst = NULL;

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {
            
		  PrintOut("Failed creating record set instance");
		  
      }
	  else
	  {
		
		CString sql = "DELETE t_Processed_Orders.PO_Ord_ID FROM t_Processed_Orders WHERE (((t_Processed_Orders.PO_Ord_ID)=";
				
		sql = sql + this->GetOrderID();
		sql = sql + "))";
		PrintOut("Revoving any partially processed leads from order. Submit SQL:" + sql);
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);		
		
	  }
	  CoUninitialize();
	  
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "Error: " + GetCString(ce.Description());
		PrintOut(str);	
		
	}
}



long COrder::ReReadQuantity(void)
{
	long quantity = 0;
	HRESULT hr = S_OK;
    try
    {
		//PrintOut("Init database connection");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

       _RecordsetPtr pRst = NULL;

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {
            
		  PrintOut("Failed creating record set instance");
		  
      }
	  else
	  {
		
		CString sql = "SELECT Ord_MakeGood ";
		sql = sql + "FROM t_Orders ";
		sql = sql + "WHERE (((Ord_Id)=";				
		sql = sql + this->GetOrderID();
		sql = sql + "))";
		PrintOut("Reread Quantity SQL:" + sql);
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockReadOnly, adCmdText);
		quantity = GetFieldLong(pRst, "Ord_MakeGood");
		
	  }
	  CoUninitialize();
	  
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "Error Reread quantity: " + GetCString(ce.Description());
		PrintOut(str);	
		
	}
	return quantity;
}

void COrder::SetOrderDate(COleDateTime aDate)
{
	oleOrderDate = aDate;
}

COleDateTime COrder::GetOrderDate(void)
{
	return oleOrderDate;
}

void COrder::SendSQLString(int id, CString str)
{
	if(mainWin)
	{
		SendMessage(mainWin, S_ORDER_TRANS, id, 0);//start string message
		int iLen = str.GetLength();
		for(int x = 0; x < iLen; x++)
		{
			int iLetter = (int) str.GetAt(x);
			SendMessage(mainWin, S_ORDER_TRANS, id, iLetter);//send a letter
		}
		SendMessage(mainWin, S_ORDER_TRANS, id, -1);//end string message
		
	}
}

void COrder::WriteLog(CString str)
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
		str = "Error COrder::WriteLog(CString str): " + GetCString(ce.Description());
		PrintOut(str);
		AfxMessageBox(str);
		
	}
	///////////////////////////////////////////////////////////////////
}
