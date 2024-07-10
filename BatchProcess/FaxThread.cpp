/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD DERIVED CLASS - produced by Worker Thread Class Generator Rel. 1.0
// FaxThread.cpp: Implementation of the CFaxThread Class.
/////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "FaxThread.h"
#include ".\faxthread.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNAMIC(CFaxThread, CThread)

/////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
/////////////////////////////////////////////////////////////////////////////////////
CFaxThread::CFaxThread(void* pOwnerObject /*= NULL*/, LPARAM lParam /*= 0L*/)
	: CThread(pOwnerObject, lParam)
	, testing(false)
	, stop(false)	
{
	// WORKER THREAD CLASS GENERATOR - Support Thread Notification //////////////////
	SUPPORT_THREAD_NOTIFICATION
	// User Initialization here...
}

CFaxThread::~CFaxThread()
{}

/////////////////////////////////////////////////////////////////////////////////////
// CFaxThread diagnostics
#ifdef _DEBUG
void CFaxThread::AssertValid() const
{
	CThread::AssertValid();
}

void CFaxThread::Dump(CDumpContext& dc) const
{
	CThread::Dump(dc);
}
#endif //_DEBUG
/////////////////////////////////////////////////////////////////////////////////////

// Unallocate all thread specific extra resources if needed while killing this thread
void CFaxThread::OnKill()
{
	// Code here...

	// This line may be removed if not necessary...
	CThread::OnKill();
}

/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD CLASS GENERATOR - Do not remove this method!
// MAIN THREAD HANDLER - The only method that must be implemented.
/////////////////////////////////////////////////////////////////////////////////////
DWORD CFaxThread::ThreadHandler()
{
	BOOL bContinue = TRUE;
	int nIncomingCommand;

	do
	{
		WaitForNotification(nIncomingCommand/*, CFaxThread::DEFAULT_TIMEOUT*/);

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
				SendSuccess(SendFax(theFaxJob));			

				
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

		/*
		case CFaxThread::CMD_USER_SPECIFIC:	// handle the user-specific command
			UserSpecificOnUserCommandHandler();
			break;
		*/

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

bool CFaxThread::SendFax(CString faxnum, CString agentname, CString agentEmail, CString company, CString contact, int div, CString filename, CString states)
{
	try
	{
		CFax aFax;
		aFax.SetParentWindow(mainWin);
		CSettings mySettings;
		mySettings.ReadSettings();
		mySettings.ReadLocations();
		CLocation myLoc;
		char buffer[200];
		sprintf(buffer, "%d", div);
		
		PrintOut(buffer, div);
		myLoc = mySettings.GetLocation(div);
		CString body;
		CString defaultReturnEmail;
		defaultReturnEmail = myLoc.GetReturnEmail();
		body = myLoc.GetBody();
		body.Replace("<agentname>", agentname);	
		body.Replace("<agentemail>", agentEmail);
		//body.Replace("<orderid>", orderid);
		body.Replace("<company>", company);
		body.Replace("<contact>", contact);	
		body.Replace("<state>", states);
		
	//PrintOut("Write temp fax body file");
	//ofstream fileout("c:\\~faxbody.txt");
	//streamsize y = body.GetLength();
	//fileout.write( body.GetBuffer(), y );
	//fileout.close();
		aFax.SetCoverPage("z:\\COVERPAGE.COV");
		aFax.SetSendCoverpage(1);
		PrintOut("Cover Note:" + body, div );
		aFax.SetCoverPageNote(body);
		aFax.SetBody(filename);
		if(testing)
		{		
			aFax.SetFaxNumber("18184444768");
			
		}
		else
		{
			aFax.SetFaxNumber(faxnum);
		}
		
		PrintOut("Attempting to send fax to: " + aFax.GetFaxNumber(), div);
		return aFax.Send();
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxThread::SendFax(CString faxnum, CString agentname, CString agentEmail, CString company, CString contact, int div, CString filename, CString states)";
		PrintOut(err, 0);
		CWriteLog::WriteFail(err);
	}
}

void CFaxThread::SetParentWindow(HWND m_hwnd)
{
		
	try
	{
		mainWin = m_hwnd;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxThread::SetParentWindow(HWND m_hwnd)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFaxThread::PrintOut(CString str, int div)
{
	try
	{	
		if(mainWin)
		{
			SendMessage(mainWin, S_START_FAX_STRING, 0, 0);
			int iLen = str.GetLength();
			for(int x = 0; x < iLen; x++)
			{
				int iLetter = (int) str.GetAt(x);
				SendMessage(mainWin, S_CHAR_FAX, iLetter, 0);
			}

			SendMessage(mainWin, S_END_FAX_STRING, 0, 0);
		}
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxThread::PrintOut(CString str, int div)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFaxThread::SetTesting(bool state)
{
	//testing = state;
}

void CFaxThread::SetStop(bool state)
{
	//stop = state;
}

void CFaxThread::SetFaxJob(CFaxJob aJob)
{
	
	try
	{
		theFaxJob = aJob;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxThread::SetFaxJob(CFaxJob aJob)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

bool CFaxThread::SendFax(CFaxJob aJob)
{
	
	try
	{
		return SendFax(aJob.GetFaxnum(), aJob.GetAgentname(), aJob.GetAgentEmail(), aJob.GetCompany(), aJob.GetContact(), aJob.GetDiv(), aJob.GetFilename(), aJob.GetStates());
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxThread::SendFax(CFaxJob aJob)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFaxThread::SendSuccess(bool state)
{
	try
	{
		if(state) MarkOrderComplete(theFaxJob);
		else 
		{
			CString msg = "Fax Fail: CampaignID(" + theFaxJob.GetCampaignID() + ") ";
			msg = msg + " Company:(" + theFaxJob.GetCompany() + ") ";
			msg = msg + " Fax Number:(" + theFaxJob.GetFaxnum() + ") ";
			msg = msg + " File Name:(" + theFaxJob.GetFilename() + ") ";
			WriteFailLog(msg);
		}
		if(mainWin)
		{
			if(state)
			{
				SendMessage(mainWin, S_FAX_SUCCESS, 0, 0);
			}
			else
			{
				SendMessage(mainWin, S_FAX_FAIL, 0, 0);
			}
		}
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxThread::SendSuccess(bool state)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFaxThread::MarkOrderComplete(CFaxJob aJob)
{
	
    try
    {
		_RecordsetPtr pRst;
		HRESULT hr = S_OK;
		//PrintOut("Init database connection");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

      

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {            
		  PrintOut("Failed creating record set instance", aJob.GetDiv());		 
      }
	  else
	  {
		  CString sql;
			sql = "SELECT * FROM t_Orders WHERE Ord_Id=" + aJob.GetOrderID();
			

		  PrintOut("Submit SQL: " + sql, aJob.GetDiv());
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);

		if(!pRst->EndOfFile) pRst->MoveFirst(); 
		else 
		{
			PrintOut("No Order Found", aJob.GetDiv());
			
		}
		
		if(!pRst->EndOfFile)//&& (!stop))
		{	
			
			PrintOut("Mark Fax Job complete", aJob.GetDiv());
			
			pRst->Fields->Item["Ord_Complete"]->Value = true;
			
			COleDateTime today = COleDateTime::GetCurrentTime();
			
			pRst->Fields->Item["Ord_Sent_On"]->Value = _variant_t(today);
			
			pRst->Update();
						
		}		
	  }	
	  CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str; 
		
		str = "Mark Fax Job Complete Error: ";
		str = str + (LPCTSTR)ce.Description();
		PrintOut(str, aJob.GetDiv());	
		
		
	}
	
  
}

void CFaxThread::WriteFailLog(CString str)
{
	CWriteLog::WriteFail(str);
	///////////////////////////////////////////////////////////////////
	
	
	/*try
    {
		HRESULT hr;
		hr = S_OK;	
		 CoInitialize(NULL);

		_RecordsetPtr pRst;		
		_bstr_t strCnn;
		pRst = NULL;

		strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BatcherLog.mdb;Jet OLEDB:Database Password=cepher;";
		 // Call Create instance to instantiate the Record set
		CString sql = "SELECT * FROM FailLog";
		hr = pRst.CreateInstance(__uuidof(Recordset));
		if(FAILED(hr))
		{            
			PrintOut("Failed creating batcher Fail log record set instance",0);
		}
		else
		{			
			pRst->Open(sql.GetBuffer(), strCnn, adOpenForwardOnly, adLockOptimistic, adCmdText);
			COleDateTime aTime;
			aTime = COleDateTime::GetCurrentTime();
			
			pRst->AddNew();			
			pRst->Fields->Item["Date"]->Value = _variant_t(aTime);	
			pRst->Fields->Item["Message"]->Value = str.GetBuffer();	
			pRst->Update();
			
		}
		CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "Error WriteFailLog: " + GetCString(ce.Description());
		PrintOut(str,0);	
		
	}*/
	///////////////////////////////////////////////////////////////////
}

CString CFaxThread::GetCString(_bstr_t bstrStart)
{
	return CStringHelper::GetCString(bstrStart);
	/*TCHAR szFinal[255];
	_stprintf(szFinal, _T("%s"), (LPCTSTR)bstrStart);
	CString str;
	str = szFinal;
	return str;*/
}

void CFaxThread::PrintOut(CString str)
{
	PrintOut(str, 0);
}
