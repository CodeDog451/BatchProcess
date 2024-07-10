#pragma once
#include "dbitem.h"

#include "resource.h"
#include "leaddb.h"
#include <time.h>
#include ".\idarray.h"
#include ".\faxjob.h"
#include ".\location.h"
#include ".\formatter.h"
#include "logdb.h"
#include "atlcomtime.h"

class COrder :
	public DBItem
{
public:
	COrder(void);
	~COrder(void);
protected:
	CString sStates;
	long iCompanyID;
	CString sContact;
	long iOrderID;
	CString sFaxNumber;
	int iQuantity;
	CString sEmail;
	CString sAgent;
	double fRateFilter;
	CString sZipFilter;
	unsigned long uliDesiredLoanFilter;
	CString sAreacodeFilter;
	bool bOnePass;
	bool bSendExcel;
	bool bSendFax;
	bool bVerified;
	bool bSendWord;
	bool bOneLeadPerPage;
	bool bSendEmailNote;
	bool bPrint;
	int iLookingTo;
	int iLang;
	int iCreditScore;
	int iHomeType;
	int iLocation;
	bool bPriority;
	bool bSendMini1003;
	bool bAdCampaign;
	
	bool bSubprime;
	int iCashout;
	int i2ndMortgage;
	int iDebtCon;
	CString sCompanyName;
	CArray<long,long&> friends;
public:
	CString ToString(void);	
	CString BooltoStr(bool state);
	CString GetLookingtoStr(int iNum);
	CString GetLangStr(int iNum);
	CString GetCreditScoreStr(int iNum);
	CString GetHomeTypeStr(int iNum);
	CString GetLocationStr(int iNum);
	CString GetThreeStateStr(int iNum);
	void SetAdCampaign(bool state);
	void SetOneLeadPerPage(bool state);
	void SetStates(CString str);
	void SetCompanyID(long iNum);
	void SetContact(CString str);
	void SetOrderID(long iNum);
	void SetFaxNumber(CString str);
	void SetQuantity(int iNum);
	void SetEmail(CString str);
	void SetAgent(CString str);
	void SetRateFilter(double fNum);
	void SetZipFilter(CString str);
	void SetDesiredLoanFilter(unsigned long ulNum);
	void SetAreacodeFilter(CString str);
	void SetOnePass(bool state);
	void SetSendExcel(bool state);
	void SetSendFax(bool state);
	void SetVerified(bool state);
	void SetSendWord(bool state);
	void SetSendEmailNote(bool state);
	void SetPrint(bool state);
	void SetLookingTo(int iNum);
	void SetLang(int iNum);
	void SetCreditScore(int iNum);
	void SetHouseType(int iNum);
	void SetLocation(int iNum);
	void SetPriority(bool state);
	void SetSendMini1003(bool state);
	void SetSubprime(bool state);
	void SetCashout(int iNum);
	void Set2ndMortgage(int iNum);
	void SetDebtCon(int iNum);
	
	void SetCompanyName(CString name);

	bool ReadOrder(_RecordsetPtr rs);
	CString GetOrderID(void);
	CString GetStates(void);
	CString GetCompanyID(void);
	CString GetCompanyName(void);
	CString GetContact(void);
	CString GetFaxNumber(void);
	CString GetQuantity(void);
	CString GetEmail(void);
	CString GetAgent(void);
protected:
	long iAgentID;
public:
	void SetAgentID(long iNum);
	CString GetAgentID(void);
	CString GetRateFilter(void);
	CString GetZipFilter(void);
	CString GetDesiredLoanFilter(void);
	CString GetAreacodeFilter(void);
	CString GetCashout(void);
	CString Get2ndMortgage(void);
	CString GetDebtCon(void);
	bool GetOnePass(void);
	bool GetExcel(void);
	bool GetSendFax(void);
	bool GetVerified(void);
	bool GetSendWord(void);
	bool GetOneLeadPerPage(void);
	bool GetSendEmailNote(void);
	bool GetPrint(void);
	CString GetLookingto(void);
	CString GetLang(void);
	CString GetCreditscore(void);
	CString GetHomeType(void);
	CString GetLocation(void);
	bool GetPriority(void);
	bool GetSendMini1003(void);
	bool GetAdCampaign(void);
	bool GetSubPrime(void);
	
	CString GetSQL(void);
protected:
	HWND mainWin;
public:
	void SetParentWindow(HWND m_hwnd);
	void PrintOut(CString str);
	void ReadLeads(void);
	
	CIDArray myAssignedLeadIDs;
	CIDArray myMatchingLeadIDs;

	void PrintMatchingLeads(void);
	CLeadDB* pt_LeadDBMap;
	void SetLeadDBMap(CLeadDB* m_LeadDBMap);
protected:
	long iSFR;
	long iMobile;
	long iCommercial;
public:
	void SetSFR(long iNum);
	CString GetSFR(void);
	void SetMobile(long iNum);
	CString GetMobile(void);
	void SetCommercial(long iNum);
	CString GetCommercial(void);
	
protected:
	CString sCity;
	double fLTV;
	bool bLTVAbove;
public:
	
	void SetCity(CString str);
	CString GetCity(void);
	void SetLTV(double fNum);
	CString GetLTV(void);
	void SetLTVAbove(bool state);
	bool GetLTVAbove(void);
	COrder& operator=(const COrder& aOrder);
	COrder(const COrder& aOrder);
	long GetOrderIDLong(void);
	bool AssignLead(void);
	void PrintAssignedLeads(void);
	long GetQuanityLong(void);
	long iMaxNeeded;
	void SetMaxNeeded(long iNum);
	void SortMatching(void);
protected:
	void QuickSort(long beg, long end);

	
	
	long GetLastIndexOfMin(void);
public:
	void AddFriend(long iCompID);
	void RemoveFriendsLeads(void);
	void ReadFriends(void);
	CIDArray FatQuickSort(CIDArray ids);
	long GetCountMatching(void);
protected:
	long iLastOrderQuantity;
public:
	void SetLastOrderQuantity(long iNum);
	CString GetLastOrderQuantity(void);
protected:
	long iCompanyOrders;
public:
	void SetCompanyOrders(long iNum);
	CString GetCompanyOrders(void);
	CString GetYield(void);
	double GetYieldDouble(void);
	void SendCEditString(int id, CString str);
protected:
	bool* pt_stop;
	
public:
	void SetStop(bool* pt_state);
	CIDArray GetMins(CIDArray ids);
	long GetCompanyIDLong(void);
	CIDArray GetAssignedLeadIDs(void);
	
	void AddAssignedLeadID(long id);
	void WriteRTF(void);
	CLeadDB* GetLeadDBMap(void);
	long GetLocationLong(void);
	long GetCountAssigned(void);
	CFaxJob GetFaxJob(CLocation myloc);
	CString GetAgentEmail(void);
protected:
	long iCampaignID;
public:
	void SetCampaignID(long iID);
	long GetCampaignID(void);
	CString GetCampaignIDString(void);
	int GetFilterCount(void);
	CString GetEmailNote(CLocation myloc);
	
	CString GetZeroBody(CLocation myLoc);
protected:
	long iMyUseLimit;
	long iDefaultUseLimit;
public:
	void SetMyUseLimit(long iNum);
	void SetDefaultUseLimit(long iNum);
	long GetUseLimit(void);
protected:
	long iMyAgeLimit;
	long iDefaultAgeLimit;
	void Progress(int iPos, int iMax);
public:
	void SetMyAgeLimit(long iNum);
	void SetDefaultAgeLimit(long iNum);
	CString GetAgeLimit(void);
	bool AlreadyHas(CString sLeadID);
protected:
	//CLogDB myLog;
public:
	CString GetSQLNew(void);
	void DeleteAssignedLeads(void);
	void ReadAssignedLeads(void);
	
	long ReReadQuantity(void);
protected:
	COleDateTime oleOrderDate;
public:
	void SetOrderDate(COleDateTime aDate);
	COleDateTime GetOrderDate(void);
	void SendSQLString(int id, CString str);
	void WriteLog(CString str);
};
