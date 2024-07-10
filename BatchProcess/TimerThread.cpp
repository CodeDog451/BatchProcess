/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD DERIVED CLASS - produced by Worker Thread Class Generator Rel. 1.0
// TimerThread.cpp: Implementation of the CTimerThread Class.
/////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "TimerThread.h"
#include ".\timerthread.h"
#include "resource.h"
#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNAMIC(CTimerThread, CThread)

/////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
/////////////////////////////////////////////////////////////////////////////////////
CTimerThread::CTimerThread(void* pOwnerObject /*= NULL*/, LPARAM lParam /*= 0L*/)
	: CThread(pOwnerObject, lParam)
	, countDown(0)
	, stop(false)
{
	// WORKER THREAD CLASS GENERATOR - Support Thread Notification //////////////////
	SUPPORT_THREAD_NOTIFICATION
	// User Initialization here...
}

CTimerThread::~CTimerThread()
{}

/////////////////////////////////////////////////////////////////////////////////////
// CTimerThread diagnostics
#ifdef _DEBUG
void CTimerThread::AssertValid() const
{
	CThread::AssertValid();
}

void CTimerThread::Dump(CDumpContext& dc) const
{
	CThread::Dump(dc);
}
#endif //_DEBUG
/////////////////////////////////////////////////////////////////////////////////////

// Unallocate all thread specific extra resources if needed while killing this thread
void CTimerThread::OnKill()
{
	// Code here...

	// This line may be removed if not necessary...
	CThread::OnKill();
}

/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD CLASS GENERATOR - Do not remove this method!
// MAIN THREAD HANDLER - The only method that must be implemented.
/////////////////////////////////////////////////////////////////////////////////////
DWORD CTimerThread::ThreadHandler()
{
	BOOL bContinue = TRUE;
	int nIncomingCommand;

	do
	{
		WaitForNotification(nIncomingCommand/*, CTimerThread::DEFAULT_TIMEOUT*/);

		/////////////////////////////////////////////////////////////////////////////
		//	Main Incoming Command Handling
		//
		//	This switch statement is just a basic skeleton showing the mechanism
		//	how to handle incoming commands. Developer may remove or rewrite
		//	necessary parts of this switch to fit his needs.
		/////////////////////////////////////////////////////////////////////////////
		switch (nIncomingCommand)
		{

		case CThread::CMD_TIMEOUT_ELAPSED:	// timeout-elapsing handling
			if (GetActivityStatus() != CThread::THREAD_PAUSED)
			{
				// UserSpecificTimeoutElapsedHandler();
				// HandleCommandImmediately(CMD_RUN);	// fire CThread::CMD_RUN command if needed
			};
			break;

		case CThread::CMD_INITIALIZE:		// initialize the thread task
			// this command should be handled; it is automatically fired when the thread starts
			// UserSpecificOnInitializeHandler();
			HandleCommandImmediately(CMD_RUN);	// fire CThread::CMD_RUN command after proper initialization
			break;

		case CThread::CMD_RUN:				// handle 'OnRun' command
			if (GetActivityStatus() != CThread::THREAD_PAUSED)
			{
				SetActivityStatus(CThread::THREAD_RUNNING);
				// UserSpecificOnRunHandler();
				stop = false;
				ProcessTimer();
			}
			break;

		case CThread::CMD_PAUSE:			// handle 'OnPause' command
			if (GetActivityStatus() != CThread::THREAD_PAUSED)
			{
				// UserSpecificOnPauseHandler();
				SetActivityStatus(CThread::THREAD_PAUSED);
			};
			break;

		case CThread::CMD_CONTINUE:			// handle 'OnContinue' command
			if (GetActivityStatus() == CThread::THREAD_PAUSED)
			{
				SetActivityStatus(CThread::THREAD_CONTINUING);
				// UserSpecificOnContinueHandler();
				HandleCommandImmediately(CMD_RUN);	// fire CThread::CMD_RUN command
			};
			break;		
		
		

		case CThread::CMD_STOP:				// handle 'OnStop' command
			// UserSpecificOnStopHandler();
			bContinue = FALSE;				// ... and leave the thread function finally
			stop = true;
			break;

		default:							// handle unknown commands...
			break;

		};

	} while (bContinue);

	return (DWORD)CThread::DW_OK;			//... if the thread task completion OK
}

void CTimerThread::SetParentWindow(HWND m_hwnd)
{
	mainWin = m_hwnd;
}

void CTimerThread::Send(int iMsg, int wp, int lp)
{
	if(mainWin)
	{
		SendMessage(mainWin, iMsg, 0, 0);
	}
}

void CTimerThread::Send(int iMsg)
{
	Send(iMsg, 0, 0);
}

void CTimerThread::ProcessTimer(void)
{
	Send(S_TIMER_START);
	while((true)&& (!stop))
	{
		Sleep(1000);
		Send(S_TIMER_FIRE);
		//countDown--;
	}	
	//Send(S_TIMER_FIRE);
}

void CTimerThread::SetStop(bool state)
{
	stop = state;
}
