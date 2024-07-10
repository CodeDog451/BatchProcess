/////////////////////////////////////////////////////////////////////////////////////
// 
// LogThread.h: Interface for the CLogThread Class.
/////////////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_LOGTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_4474C705__INCLUDED_)
#define AFX_LOGTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_4474C705__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Thread.h"
#include "logentry.h"
#include "logdb.h"
#include "resource.h"

class CLogThread : public CThread  
{
public:
	DECLARE_DYNAMIC(CLogThread)

// Construction & Destruction
	CLogThread(void* pOwnerObject = NULL, LPARAM lParam = 0L);
	virtual ~CLogThread();

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
	CLogEntry myEntry;
	CLogDB myLog;
public:
	void SetEntry(CLogEntry aEntry);
protected:
	HWND mainWin;
public:
	void SetParentWindow(HWND m_hwnd);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_LOGTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_4474C705__INCLUDED_)
