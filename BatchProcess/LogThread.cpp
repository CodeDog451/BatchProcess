/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD DERIVED CLASS - produced by Worker Thread Class Generator Rel. 1.0
// LogThread.cpp: Implementation of the CLogThread Class.
/////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "LogThread.h"
#include ".\logthread.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNAMIC(CLogThread, CThread)

/////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
/////////////////////////////////////////////////////////////////////////////////////
CLogThread::CLogThread(void* pOwnerObject /*= NULL*/, LPARAM lParam /*= 0L*/)
	: CThread(pOwnerObject, lParam)
{
	// User Initialization here...
}

CLogThread::~CLogThread()
{}

/////////////////////////////////////////////////////////////////////////////////////
// CLogThread diagnostics
#ifdef _DEBUG
void CLogThread::AssertValid() const
{
	CThread::AssertValid();
}

void CLogThread::Dump(CDumpContext& dc) const
{
	CThread::Dump(dc);
}
#endif //_DEBUG
/////////////////////////////////////////////////////////////////////////////////////

// Unallocate all thread specific extra resources if needed while killing this thread
void CLogThread::OnKill()
{
	// Code here...

	// This line may be removed if not necessary...
	CThread::OnKill();
}

/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD CLASS GENERATOR - Do not remove this method!
// MAIN THREAD HANDLER - The only method that must be implemented.
/////////////////////////////////////////////////////////////////////////////////////
DWORD CLogThread::ThreadHandler()
{
	// Trivial Thread specific code here...
	if(myEntry.GetFail())
	{
		myLog.WriteFail(myEntry.GetEntry(), myEntry.GetDiv());
	}
	else
	{
		myLog.WriteLog(myEntry.GetEntry(), myEntry.GetDiv());
	}
	if(mainWin)
	{
		SendMessage(mainWin, LOG_WRITE_DONE, 0, 0);
	}
	return (DWORD)CThread::DW_OK;	//... if the thread task completion OK
}

void CLogThread::SetEntry(CLogEntry aEntry)
{
	myEntry = aEntry;
}

void CLogThread::SetParentWindow(HWND m_hwnd)
{
	mainWin = m_hwnd;
}
