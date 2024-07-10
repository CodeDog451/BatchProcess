#pragma once
#include "dbitem.h"
#include "afxtempl.h"
#include ".\order.h"
#include "leaddb.h"
#include "idarray.h"
#include ".\rtfwriter.h"
#include ".\location.h"
#include ".\settings.h"
#include "ESMTP.h"
#include ".\txtwriter.h"
#include "settings.h"
#include ".\csvwriter.h"
#include "logdb.h"

class COrderDB :
	public DBItem
{
public:
	COrderDB(void);
	~COrderDB(void);
protected:
	CMap<long, long, COrder, COrder&> myOrderMap;
	CMap<long, long, COrder, COrder&> myCompleteOrderMap;
public:
	void Add(COrder aOrder);
	bool GetOrder(long iID, COrder &aOrder);
	void RemoveAll(void);
	void PrintOrders(void);
protected:
	bool stop;
public:
	void PrintOut(CString str, int div);
protected:
	HWND mainWin;
public:
	void SetParentWindow(HWND m_wnd);
	void SetStop(bool state);
	bool AssignLeads(void);
	long GetMaxNeeded(void);
	bool ReadLeads(void);
	void LoadScreenFromOrder(COrder* aOrder);
	void SendCEditString(int id, CString str);
protected:
	void SendCheckBoxState(int id, bool state);
public:
	void SendSQLString(int id, CString str);
protected:
	bool bOnePass;
	bool bVerified;
	
public:
	void SetOnePass(bool state);
	void SetVerified(bool state);
	CString GetOnePass(void);
	CString GetVerified(void);
protected:
	bool bPriority;
public:
	void SetPriority(bool state);
	CString GetPriority(void);
	CString BoolToCString(bool state);
	CString GetSQL(void);
	
protected:
	CLeadDB leadDBMap;
public:
	CString GetCString(_bstr_t bstrStart);
	void Reset(void);
	bool Process(void);
	
	bool AssignOrders(void);

public:
	CIDArray FatQuickSort(CIDArray ids);
	COrder GetOrder(long id);
	COrderDB& operator=(const COrderDB& aCOrderDB);
	COrderDB(const COrderDB& aCOrderDB);
	COrderDB GetOrdersWithPastZeroSent(void);
	COrderDB GetOrdersWithLowPartialsSent(void);
	double GetLastPercentSent(long iCompID);
	bool bZeros;
	void SetZeros(bool state);
protected:
	bool bNewCustomer;
public:
	void SetNewCustomer(bool state);
	void SplitYield(CIDArray ids, CIDArray *low, CIDArray *high);
	
	void GenerateRTFs(void);
	bool SendMail(CString brokerEmail, CString agentname, CString agentEmail, CString company, CString contact, int div, CString orderid);
protected:
	bool testing;
public:
	void SendOrders(void);
	bool SendEmail(COrder aOrder);
	void MarkOrderComplete(COrder aOrder);
	void MarkOrderReadyToSend(COrder aOrder);
	void SendFax(COrder aOrder);
protected:
	CSettings mySettings;

	CString GetIntheBody(COrder aOrder);
public:
	void Print(COrder aOrder);
	void Progress(int iPos, int iMax);
	long GetCountAssigned(void);
	CString GetString(long iNum);
protected:
	//CLogDB myLog;
public:
	void PrintOrderTotals(void);
	void PrintFail(CString str);
	bool ReadOrders(void);
	bool ReadOrders(bool bRetry);
	CString GetSQLRetry(void);
	
	
	void DeleteAssignedLeads(void);
	void WriteLog(CString str);
	bool ReadReadytoSend(bool bGenDoc);
protected:
	void ProgressAll(int iPos, int iMax);
public:
	void PrintOut(CString str);
	

protected:
	CString sState;
public:
	void SetState(CString sStateAbr);
	CString GetState();
	bool DoesFileExist(CString sFilename);
};
