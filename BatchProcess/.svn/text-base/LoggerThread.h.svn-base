/////////////////////////////////////////////////////////////////////////////////////
// 
// LoggerThread.h: Interface for the CLoggerThread Class.
/////////////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_LOGGERTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_4474EF27__INCLUDED_)
#define AFX_LOGGERTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_4474EF27__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Thread.h"
#include ".\logdb.h"

class CLoggerThread : public CThread  
{
public:
	DECLARE_DYNAMIC(CLoggerThread)

// Construction & Destruction
	CLoggerThread(void* pOwnerObject = NULL, LPARAM lParam = 0L);
	virtual ~CLoggerThread();

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
	CString sMessage;
	CString sFailMessage;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_LOGGERTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_4474EF27__INCLUDED_)
