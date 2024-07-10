/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD DERIVED CLASS - produced by Worker Thread Class Generator Rel. 1.0
// LoggerThread.cpp: Implementation of the CLoggerThread Class.
/////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "LoggerThread.h"
#include "resource.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNAMIC(CLoggerThread, CThread)

/////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
/////////////////////////////////////////////////////////////////////////////////////
CLoggerThread::CLoggerThread(void* pOwnerObject /*= NULL*/, LPARAM lParam /*= 0L*/)
	: CThread(pOwnerObject, lParam)
	, sMessage(_T(""))
	, sFailMessage(_T(""))
{
	// WORKER THREAD CLASS GENERATOR - Support Thread Notification //////////////////
	SUPPORT_THREAD_NOTIFICATION
	// User Initialization here...
}

CLoggerThread::~CLoggerThread()
{}

/////////////////////////////////////////////////////////////////////////////////////
// CLoggerThread diagnostics
#ifdef _DEBUG
void CLoggerThread::AssertValid() const
{
	CThread::AssertValid();
}

void CLoggerThread::Dump(CDumpContext& dc) const
{
	CThread::Dump(dc);
}
#endif //_DEBUG
/////////////////////////////////////////////////////////////////////////////////////

// Unallocate all thread specific extra resources if needed while killing this thread
void CLoggerThread::OnKill()
{
	// Code here...

	// This line may be removed if not necessary...
	CThread::OnKill();
}

/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD CLASS GENERATOR - Do not remove this method!
// MAIN THREAD HANDLER - The only method that must be implemented.
/////////////////////////////////////////////////////////////////////////////////////
DWORD CLoggerThread::ThreadHandler()
{
	BOOL bContinue = TRUE;
	int nIncomingCommand;

	do
	{
		WaitForNotification(nIncomingCommand/*, CLoggerThread::DEFAULT_TIMEOUT*/);

		/////////////////////////////////////////////////////////////////////////////
		//	Main Incoming Command Handling
		//
		//	This switch statement is just a basic skeleton showing the mechanism
		//	how to handle incoming commands. Developer may remove or rewrite
		//	necessary parts of this switch to fit his needs.
		/////////////////////////////////////////////////////////////////////////////
		
		
		
		if((nIncomingCommand >= N_MESSAGE) && (nIncomingCommand <= (N_MESSAGE + 256)))
		{
			short iMessage = 0;
			iMessage = nIncomingCommand - N_MESSAGE;
			//building string
			char c = (char) iMessage;
			sMessage = sMessage + c;
		}
		if(nIncomingCommand == N_START_STRING)
		{
			sMessage = "";
		}
		if(nIncomingCommand == N_STOP_STRING)
		{
			CLogDB aLog;
			aLog.WriteLog(sMessage, 0);
			sMessage = "";
		}




		short iFailMessage = 0;
		iFailMessage = nIncomingCommand - 7000;
		if((iFailMessage >= 0) && (iFailMessage <= 256))
		{
			//building string
			char c = (char) iFailMessage;
			sFailMessage = sFailMessage + c;
		}
		if(nIncomingCommand == -1)
		{
			sFailMessage = "";
		}
		if(nIncomingCommand == -2)
		{
			CLogDB aLog;
			aLog.WriteFail(sFailMessage, 0);
			sFailMessage = "";
		}
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

		
		//case -1:	// handle the user-specific command
			//sMessage = "";
			//break;

		//case -2:	// handle the user-specific command
			//CLogDB aLog;
			//aLog.WriteFail(sMessage, 0);
			//sMessage = "";
			
			//break;
		

		case CThread::CMD_STOP:				// handle 'OnStop' command
			// UserSpecificOnStopHandler();
			bContinue = FALSE;				// ... and leave the thread function finally
			break;

		default:							// handle unknown commands...
			break;

		};

	} while (bContinue);

	return (DWORD)CThread::DW_OK;			//... if the thread task completion OK
}
