/////////////////////////////////////////////////////////////////////////////////////
// 
// TimerThread.h: Interface for the CTimerThread Class.
/////////////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_TIMERTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_442D9BA2__INCLUDED_)
#define AFX_TIMERTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_442D9BA2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Thread.h"

class CTimerThread : public CThread  
{
public:
	DECLARE_DYNAMIC(CTimerThread)

// Construction & Destruction
	CTimerThread(void* pOwnerObject = NULL, LPARAM lParam = 0L);
	virtual ~CTimerThread();

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
	void Send(int iMsg, int wp, int lp);
	void Send(int iMsg);
	void ProcessTimer(void);
	int countDown;
protected:
	bool stop;
public:
	void SetStop(bool state);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TIMERTHREAD_H__A4C9C0B8_CD6D_11D2_BB7E_442D9BA2__INCLUDED_)
