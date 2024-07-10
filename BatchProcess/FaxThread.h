/////////////////////////////////////////////////////////////////////////////////////
// 
// FaxThread.h: Interface for the CFaxThread Class.
/////////////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_FAXTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_4460B767__INCLUDED_)
#define AFX_FAXTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_4460B767__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Thread.h"
#include "resource.h"
#include ".\fax.h"
#include ".\settings.h"
#include "faxjob.h"
#include ".\stringhelper.h"
//#include "logdb.h"
//#include "logentry.h"

class CFaxThread : public CThread  
{
public:
	DECLARE_DYNAMIC(CFaxThread)

// Construction & Destruction
	CFaxThread(void* pOwnerObject = NULL, LPARAM lParam = 0L);
	virtual ~CFaxThread();

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
public:
	bool SendFax(CString faxnum, CString agentname, CString agentEmail, CString company, CString contact, int div, CString filename, CString states);
protected:
	HWND mainWin;
public:
	void SetParentWindow(HWND m_hwnd);
	void PrintOut(CString str, int div);
protected:
	bool testing;
public:
	void SetTesting(bool state);
protected:
	bool stop;
public:
	void SetStop(bool state);

protected:
	CFaxJob theFaxJob;
public:
	void SetFaxJob(CFaxJob aJob);
	bool SendFax(CFaxJob aJob);
	void SendSuccess(bool state);
protected:
	void MarkOrderComplete(CFaxJob aJob);
	//CLogDB myLog;
	
public:
	void WriteFailLog(CString str);
	CString GetCString(_bstr_t bstrStart);
	void PrintOut(CString str);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FAXTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_4460B767__INCLUDED_)
