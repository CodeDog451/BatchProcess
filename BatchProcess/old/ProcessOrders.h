/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD DERIVED CLASS - produced by Worker Thread Class Generator Rel 1.0
// ProcessOrders.h: Interface for the CProcessOrders Class.
/////////////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_PROCESSORDERS_H__A4C9C0B8_CD6D_11D2_BB7E_44075ECD__INCLUDED_)
#define AFX_PROCESSORDERS_H__A4C9C0B8_CD6D_11D2_BB7E_44075ECD__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Thread.h"
#include "leadlist.h"

class CProcessOrders : public CThread  
{
public:
	DECLARE_DYNAMIC(CProcessOrders)

// Construction & Destruction
	CProcessOrders(void* pOwnerObject = NULL, LPARAM lParam = 0L);
	virtual ~CProcessOrders();

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
	
	void SetParentWindow(HWND m_hwnd);
private:
	HWND mainWin;
	LeadList leads;
public:
	long ReadLeads(void);
private:
	void PrintOut(int msg_id);
	void PrintOut(int msg_id, int count);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PROCESSORDERS_H__A4C9C0B8_CD6D_11D2_BB7E_44075ECD__INCLUDED_)
