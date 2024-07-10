/////////////////////////////////////////////////////////////////////////////////////
// 
// ProcessThread.h: Interface for the CProcessThread Class.
/////////////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_PROCESSTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_4446D1A0__INCLUDED_)
#define AFX_PROCESSTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_4446D1A0__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Thread.h"
#include "resource.h"
#include ".\order.h"
#include "orderdb.h"
//#include "logdb.h"
#include ".\genorders.h"
#include ".\pop3socket.h"
#include ".\stats.h"

class CProcessThread : public CThread  
{
public:
	DECLARE_DYNAMIC(CProcessThread)

// Construction & Destruction
	CProcessThread(void* pOwnerObject = NULL, LPARAM lParam = 0L);
	virtual ~CProcessThread();

// Operations & Overridables
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:
// Overridables
	virtual void OnKill();
	/////////////////////////////////////////////////////////////////////////////////
	// Main Thread Handler
	// The only method that must be implemented
	virtual	DWORD ThreadHandler();
	/////////////////////////////////////////////////////////////////////////////////
	HWND mainWin;
public:
	void SetParentWindow(HWND m_hwnd);
	void PrintOut(CString str, int div);
protected:
	bool stop;
	//CLeadDB leadDBMap;
	//COrderDB orderDBMap;
	void ProcessOrders(int iLeadType, int iOrderType);
	//CMap<long, long, COrder, COrder&> myOrderMap;
public:
	void SetStop(bool state);
	
	void LoadScreenFromOrder(COrder* aOrder);
	void SendCEditString(int id, CString str);
	void SendCheckBoxState(int id, bool state);
	void SendSQLString(int id, CString str);
	
	CString GetCString(_bstr_t bstrStart);
	
	
	virtual void ProcessBatch(void);
	void Progress(int iPos, int iMax);
protected:
	int iCategory;
	int iGroup;
	//int iState;
	//CLogDB myLog;
	int test;
	CString sCategory;
	CString sGroup;
	CString sState;
public:
	void RetrySending(void);
	
	void ResendNoDocGen(void);
	void GenerateOrders(void);
	void WriteLog(CString str);
protected:
	bool CheckPop3Status(void);
public:
	void PrintOut(CString str);
protected:
	void SendStatsEmail(void);
	int MaxLength(CString str1, CString str2, CString str3, CString str4);
	CString SpaceFill(CString str, int iLength);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PROCESSTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_4446D1A0__INCLUDED_)
