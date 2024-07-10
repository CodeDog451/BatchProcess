/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD DERIVED CLASS - produced by Worker Thread Class Generator Rel. 1.0
// ProcessThread.cpp: Implementation of the CProcessThread Class.
/////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ProcessThread.h"
#include ".\processthread.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNAMIC(CProcessThread, CThread)

/////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
/////////////////////////////////////////////////////////////////////////////////////
CProcessThread::CProcessThread(void* pOwnerObject /*= NULL*/, LPARAM lParam /*= 0L*/)
	: CThread(pOwnerObject, lParam)
	, stop(true)
	, iCategory(5000)//Verified 
	, iGroup(5006)// normal are the default
	, sCategory(_T(""))
	, sGroup(_T(""))
	, sState(_T(""))
{
	// WORKER THREAD CLASS GENERATOR - Support Thread Notification //////////////////
	SUPPORT_THREAD_NOTIFICATION
	// User Initialization here...
}

CProcessThread::~CProcessThread()
{}

/////////////////////////////////////////////////////////////////////////////////////
// CProcessThread diagnostics
#ifdef _DEBUG
void CProcessThread::AssertValid() const
{
	CThread::AssertValid();
}

void CProcessThread::Dump(CDumpContext& dc) const
{
	CThread::Dump(dc);
}
#endif //_DEBUG
/////////////////////////////////////////////////////////////////////////////////////

// Unallocate all thread specific extra resources if needed while killing this thread
void CProcessThread::OnKill()
{
	// Code here...

	// This line may be removed if not necessary...
	CThread::OnKill();
}

/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD CLASS GENERATOR - Do not remove this method!
// MAIN THREAD HANDLER - The only method that must be implemented.
/////////////////////////////////////////////////////////////////////////////////////
DWORD CProcessThread::ThreadHandler()
{
	BOOL bContinue = TRUE;
	int nIncomingCommand;

	do
	{
		WaitForNotification(nIncomingCommand/*, CProcessThread::DEFAULT_TIMEOUT*/);

		/////////////////////////////////////////////////////////////////////////////
		//	Main Incoming Command Handling
		//
		//	This switch statement is just a basic skeleton showing the mechanism
		//	how to handle incoming commands. Developer may remove or rewrite
		//	necessary parts of this switch to fit his needs.
		/////////////////////////////////////////////////////////////////////////////
		if((nIncomingCommand >=5000) && (nIncomingCommand <=5002))
		{
			iCategory = nIncomingCommand;
		}
		if((nIncomingCommand >=5003) && (nIncomingCommand <=5006))
		{
			iGroup = nIncomingCommand;
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
				//stop = false;				
				//ProcessBatch();
				
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

		
		case PO_RETRY_SENDING:	
			this->RetrySending();
			break;
		
		case PO_RETRY_SEND_NODOCGEN:	
			this->ResendNoDocGen();
			break;

		case PO_GEN_ORDERS:	
			GenerateOrders();
			break;

		case PO_POP_CHECK:	
			CheckPop3Status();
			break;

		case PO_STATS:	
			SendStatsEmail();
			break;
			
			
		case CMD_PROCESS:	// handle the user-specific command
			stop = false;			
			ProcessBatch();
			
			break;

		case CThread::CMD_STOP:				// handle 'OnStop' command
			// UserSpecificOnStopHandler();
			bContinue = FALSE;				// ... and leave the thread function finally
			break;
		case PO_AK: 
			sState = "AK";
			break; 
		case PO_AL: 
			sState = "AL";
			break; 
		case PO_AR: 
			sState = "AR";
			break;                           
		case PO_AZ:
			sState = "AZ";
			break;                           
		case PO_CA:
			sState = "CA";
			break;                          
		case PO_CO: 
			sState = "CO";
			break;                         
		case PO_CT: 
			sState = "CT";
			break;                        
		case PO_DC: 
			sState = "DC";
			break;                       
		case PO_DE: 
			sState = "DE";
			break;                            
		case PO_FL: 
			sState = "FL";
			break;                            
		case PO_GA: 
			sState = "GA";
			break;                             
		case PO_HI: 
			sState = "HI";
			break;                            
		case PO_IA: 
			sState = "IA";
			break;                            
		case PO_ID: 
			sState = "ID";
			break;                            
		case PO_IL: 
			sState = "IL";
			break;                            
		case PO_IN: 
			sState = "IN";
			break;                            
		case PO_KS: 
			sState = "KS";
			break;                            
		case PO_KY: 
			sState = "KY";
			break;                            
		case PO_LA: 
			sState = "LA";
			break;                            
		case PO_MA: 
			sState = "MA";
			break;                            
		case PO_MD: 
			sState = "MD";
			break;                            
		case PO_ME: 
			sState = "ME";
			break;                            
		case PO_MI: 
			sState = "MI";
			break;                            
		case PO_MN: 
			sState = "MN";
			break;                            
		case PO_MO: 
			sState = "MO";
			break;                            
		case PO_MS: 
			sState = "MS";
			break;                            
		case PO_MT:  
			sState = "MT";
			break;                           
		case PO_NC:  
			sState = "NC";
			break;                           
		case PO_ND:  
			sState = "ND";
			break;                           
		case PO_NE:
			sState = "NE";
			break;
		case PO_NH:   
			sState = "NH";
			break;                        
		case PO_NJ: 
			sState = "NJ";
			break;                        
		case PO_NM:     
			sState = "NM";
			break;                        
		case PO_NV:     
			sState = "NV";
			break;                        
		case PO_NY:     
			sState = "NY";
			break;                        
		case PO_OH:     
			sState = "OH";
			break;                        
		case PO_OK:     
			sState = "OK";
			break;                        
		case PO_OR:     
			sState = "OR";
			break;                        
		case PO_PA:     
			sState = "PA";
			break;                        
		case PO_RI:     
			sState = "RI";
			break;                        
		case PO_SC:     
			sState = "SC";
			break;                        
		case PO_SD:  
			sState = "SD";
			break;                        
		case PO_TN:     
			sState = "TN";
			break;                        
		case PO_TX:     
			sState = "TX";
			break;                        
		case PO_UT:     
			sState = "UT";
			break;                        
		case PO_VA:     
			sState = "VA";
			break;                        
		case PO_VT:     
			sState = "VT";
			break;                        
		case PO_WA:     
			sState = "WA";
			break;                        
		case PO_WI:     
			sState = "WI";
			break;                        
		case PO_WV:     
			sState = "WV";
			break;                        
		case PO_WY:     
			sState = "WY";
			break;  
		case PO_ALL:     
			sState = "";
			break;     


		default:							// handle unknown commands...
			break;

		};

	} while (bContinue);

	return (DWORD)CThread::DW_OK;			//... if the thread task completion OK
}

void CProcessThread::SetParentWindow(HWND m_hwnd)
{	
	try
	{
		mainWin = m_hwnd;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CProcessThread::SetParentWindow(HWND m_hwnd)";
		PrintOut(err);
		WriteLog(err);
	}
}

void CProcessThread::PrintOut(CString str, int div)
{	
	
	try
	{
		if(mainWin)
		{
			
			SendMessage(mainWin, S_START_MESSAGE, 0, 0);
			int iLen = str.GetLength();
			for(int x = 0; x < iLen; x++)
			{
				int iLetter = (int) str.GetAt(x);
				SendMessage(mainWin, S_CHAR_MESSAGE, iLetter, 0);
			}

			SendMessage(mainWin, S_END_MESSAGE, 0, 0);
		}

	}
	catch(...)
	{
		//error trap template
		CString err;
		err = "Error: PrintOut(CString str, int div)";
		PrintOut(err);
		WriteLog(err);
	}
}

void CProcessThread::SetStop(bool state)
{
	//stop = state;
	//orderDBMap.SetStop(state);
	//leadDBMapSetStop(state);
	//PrintOut("Stop process", 0);
}

void CProcessThread::ProcessOrders(int iLeadType, int iOrderType)
{
	//COrderDB aOrderDB;
	//orderDBMap = aOrderDB;
	try
	{
		COrderDB orderDBMap;
		if(mainWin)
		{		
			SendMessage(mainWin, iLeadType, 0, 0);
			SendMessage(mainWin, iOrderType, 0, 0);
		}

		orderDBMap.Reset();
		orderDBMap.SetParentWindow(mainWin);
		orderDBMap.SetState(sState);
		switch (iLeadType)
		{
			case PO_VERIFIED:
				sCategory = "Verifed";
				orderDBMap.SetVerified(true);
				orderDBMap.SetOnePass(false);
				break;
			case PO_INTERNET: 
				sCategory = "Internet";
				orderDBMap.SetVerified(false);
				orderDBMap.SetOnePass(false);
				break;
			case PO_ONEPASS:
				sCategory = "One-Pass";
				orderDBMap.SetVerified(false);
				orderDBMap.SetOnePass(true);
				break;
			default:
				//error should not get here
				PrintOut("Invalid Leads Type", 0);
				break;
		}
		switch (iOrderType)
		{
			case PO_PRIORITY:
				sGroup = "Priority";
				orderDBMap.SetPriority(true);
				orderDBMap.SetZeros(false);
				orderDBMap.SetNewCustomer(false);
				break;
			case PO_SENTZERO: 
				sGroup = "Sent Zero on last order";
				orderDBMap.SetPriority(false);
				orderDBMap.SetZeros(true);
				orderDBMap.SetNewCustomer(false);
				break;
			case PO_NEWCUSTOMER:
				sGroup = "New Customer";
				orderDBMap.SetPriority(false);
				orderDBMap.SetZeros(false);
				orderDBMap.SetNewCustomer(true);
				break;
			case PO_NORMAL:
				sGroup = "Normal";
				orderDBMap.SetPriority(false);
				orderDBMap.SetZeros(false);
				orderDBMap.SetNewCustomer(false);
				break;
			default:
				//error should not get here
				PrintOut("Invalid Orders Type", 0);
				break;
		}	
		CString str;
		str = "Process (" + sCategory;
		str = str + ") (" + sGroup;
		str = str + ") State:" + sState;
		PrintOut(str, 0);
		if(orderDBMap.Process())
		{
			//orderDBMap.PrintOrders();
			orderDBMap.GenerateRTFs();
			if(CheckPop3Status())
			{
				orderDBMap.SendOrders();
			}
			else
			{
				PrintOut("There was a problem connecting to the email server");
			}
			
		}
		//orderDBMap.Reset();
	}	
	catch(...)
	{
		PrintOut("Error: CProcessThread::ProcessOrders");
		WriteLog("Error: CProcessThread::ProcessOrders");
	}
	
	
}

void CProcessThread::LoadScreenFromOrder(COrder* aOrder)
{
	try
	{
		SendCEditString(IDC_EDIT_COMPANYID, aOrder->GetCompanyID());
		SendCEditString(IDC_EDIT_CAMPAIGNID, aOrder->GetCampaignIDString());
		SendCEditString(IDC_EDIT_STATES, aOrder->GetStates());//
		SendCEditString(IDC_EDIT_CONTACT, aOrder->GetContact());
		SendCEditString(IDC_EDIT_ORDERID, aOrder->GetOrderID());
		SendCEditString(IDC_EDIT_FAXNUMBER, aOrder->GetFaxNumber());
		SendCEditString(IDC_EDIT_QUANTITY, aOrder->GetQuantity());
		SendCEditString(IDC_EDIT_EMAIL, aOrder->GetEmail());
		SendCEditString(IDC_EDIT_AGENT, aOrder->GetAgent());
		SendCEditString(IDC_EDIT_RATE, aOrder->GetRateFilter());
		SendCEditString(IDC_EDIT_ZIP_FILTER, aOrder->GetZipFilter());
		SendCEditString(IDC_EDIT_DESIRED_LOAN_FILTER, aOrder->GetDesiredLoanFilter());
		SendCEditString(IDC_EDIT_AREACODE_FILTER, aOrder->GetAreacodeFilter());
		SendCEditString(IDC_EDIT_CASHOUT, aOrder->GetCashout());
		SendCEditString(IDC_EDIT_2NDMORTGAGE, aOrder->Get2ndMortgage());
		SendCEditString(IDC_EDIT_DEBTCON, aOrder->GetDebtCon());
		SendCEditString(IDC_EDIT_AGENTID, aOrder->GetAgentID());
		SendCEditString(IDC_EDIT_LOOKINGTO, aOrder->GetLookingto());
		SendCEditString(IDC_EDIT_LANG, aOrder->GetLang());
		SendCEditString(IDC_EDIT_CREDIT, aOrder->GetCreditscore());
		SendCEditString(IDC_EDIT_HOMETYPE, aOrder->GetHomeType());
		SendCEditString(IDC_EDIT_LOCATION, aOrder->GetLocation());
		SendCEditString(IDC_EDIT_SFR, aOrder->GetSFR());
		SendCEditString(IDC_EDIT_MOBILE, aOrder->GetMobile());
		SendCEditString(IDC_EDIT_COMMERCIAL, aOrder->GetCommercial());
		SendCEditString(IDC_EDIT_CITY, aOrder->GetCity());
		SendCEditString(IDC_EDIT_LTV, aOrder->GetLTV());
		SendSQLString(IDC_EDIT_SQL, aOrder->GetSQL());//
		SendCheckBoxState(IDC_CHECK_VERIFIED, aOrder->GetVerified());
		SendCheckBoxState(IDC_CHECK_ONEPASS, aOrder->GetOnePass());
		SendCheckBoxState(IDC_CHECK_WORD, aOrder->GetSendWord());
		SendCheckBoxState(IDC_CHECK_EXCEL, aOrder->GetExcel());
		SendCheckBoxState(IDC_CHECK_FAX, aOrder->GetSendFax());
		SendCheckBoxState(IDC_CHECK_ONELEAD, aOrder->GetOneLeadPerPage());
		SendCheckBoxState(IDC_CHECK_EMAILNOTE, aOrder->GetSendEmailNote());
		SendCheckBoxState(IDC_CHECK_PRINT, aOrder->GetPrint());//
		SendCheckBoxState(IDC_CHECK_PRIORITY, aOrder->GetPriority());
		SendCheckBoxState(IDC_CHECK_MINI1003, aOrder->GetSendMini1003());
		SendCheckBoxState(IDC_CHECK_ADCAMPAIGN, aOrder->GetAdCampaign());
		SendCheckBoxState(IDC_CHECK_SUBPRIME, aOrder->GetSubPrime());
		SendCheckBoxState(IDC_CHECK_ABOVE, aOrder->GetLTVAbove());
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CProcessThread::LoadScreenFromOrder(COrder* aOrder)";
		PrintOut(err);
		WriteLog(err);
	}


}

void CProcessThread::SendCEditString(int id, CString str)
{
	try
	{
		if(mainWin)
		{
			SendMessage(mainWin, S_ORDER_TRANS, id, 0);
			int iLen = str.GetLength();
			for(int x = 0; x < iLen; x++)
			{
				int iLetter = (int) str.GetAt(x);
				SendMessage(mainWin, S_ORDER_TRANS, id, iLetter);
			}		
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CProcessThread::SendCEditString(int id, CString str)";
		PrintOut(err);
		WriteLog(err);
	}
}

void CProcessThread::SendCheckBoxState(int id, bool state)
{
	try
	{
		if(mainWin)
		{
			SendMessage(mainWin, S_ORDER_TRANS, id, (int)state);		
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CProcessThread::SendCheckBoxState(int id, bool state)";
		PrintOut(err);
		WriteLog(err);
	}
}

void CProcessThread::SendSQLString(int id, CString str)
{
	if(mainWin)
	{
		SendMessage(mainWin, S_ORDER_TRANS, id, 0);//start string message
		int iLen = str.GetLength();
		for(int x = 0; x < iLen; x++)
		{
			int iLetter = (int) str.GetAt(x);
			SendMessage(mainWin, S_ORDER_TRANS, id, iLetter);//send a letter
		}
		SendMessage(mainWin, S_ORDER_TRANS, id, -1);//end string message

		
	}
}


CString CProcessThread::GetCString(_bstr_t bstrStart)
{	
	return CStringHelper::GetCString(bstrStart);
	/*try
	{
		TCHAR szFinal[255];
		_stprintf(szFinal, _T("%s"), (LPCTSTR)bstrStart);
		CString str;
		str = szFinal;
		return str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CProcessThread::GetCString(_bstr_t bstrStart";
		PrintOut(err);
		WriteLog(err);
		return "";
	}*/
}


void CProcessThread::ProcessBatch(void)
{
	//ProcessOrders(PO_ONEPASS, PO_NORMAL);
	//ProcessOrders(PO_VERIFIED, PO_NORMAL);
	//ProcessOrders(PO_INTERNET, PO_NORMAL);
	try
	{
		Progress(0, 5100);
		char buffer[200];
		for(int x = 0; x <=5100; x++)
		{
			Progress(x, 5100);
			
			sprintf(buffer, "%d", x);
			PrintOut(buffer, 0);
			//Sleep(10);
		}
		PrintOut("Done", 0);
		if(mainWin)
		{		
			SendMessage(mainWin, PO_DONE, 0, 0);		
		}	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CProcessThread::ProcessBatch(void)";
		PrintOut(err);
		WriteLog(err);
	}
}

void CProcessThread::Progress(int iPos, int iMax)
{	
	try
	{
		if(mainWin)
		{
			SendMessage(mainWin, ID_PROGRESS, iPos, iMax);
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CProcessThread::Progress(int iPos, int iMax)";
		PrintOut(err);
		WriteLog(err);
	}
}

void CProcessThread::RetrySending(void)
{
	try
	{	
		COrderDB orderDBMap;
		if(mainWin)
		{		
			SendMessage(mainWin, PO_RETRY_SENDING, 0, 0);
			
		}	
		orderDBMap.SetParentWindow(mainWin);	
		CString str;
		str = "Retry Sending ";
		PrintOut(str, 0);
		bool bRetry = true;
		if(CheckPop3Status())
		{
			if(orderDBMap.ReadReadytoSend(true))
			{		
				PrintOut("Done", 0);
							
			}
			if(mainWin)
			{		
				SendMessage(mainWin, PO_DONE, 0, 0);		
			}
		}
		else
		{
			PrintOut("There was a problem connecting to the email server");
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CProcessThread::RetrySending";
		PrintOut(err);
		WriteLog(err);
	}
}


void CProcessThread::ResendNoDocGen(void)
{
	try
	{
		COrderDB orderDBMap;
		if(mainWin)
		{		
			SendMessage(mainWin, PO_RETRY_SEND_NODOCGEN, 0, 0);
			
		}	
		orderDBMap.SetParentWindow(mainWin);	
		CString str;
		str = "Retry Sending No Doc Gen";
		PrintOut(str, 0);
		bool bRetry = true;
		if(orderDBMap.ReadReadytoSend(false))
		{		
			PrintOut("Done", 0);
			if(mainWin)
			{		
				SendMessage(mainWin, PO_DONE, 0, 0);		
			}		
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CProcessThread::ResendNoDocGen";
		PrintOut(err);
		WriteLog(err);
	}
}

void CProcessThread::GenerateOrders(void)
{
	try
	{
		CGenOrders aGenOrders;
		aGenOrders.SetParentWindow(mainWin);
		WriteLog("Generate Orders");
		aGenOrders.ReadCampaigns();
		
		PrintOut("Generate Orders Done", 0);
		if(mainWin)
		{		
			SendMessage(mainWin, PO_DONE, 0, 0);		
		}
	}
	catch(...)
	{
		//error trap template
		CString err;
		err = "Error: CProcessThread::GenerateOrders";
		PrintOut(err);
		WriteLog(err);
	}
}

void CProcessThread::WriteLog(CString str)
{
	///////////////////////////////////////////////////////////////////
	HRESULT hr;
	hr = S_OK;	
	
	try
    {
		 CoInitialize(NULL);

		_RecordsetPtr pRst;		
		_bstr_t strCnn;
		pRst = NULL;

		strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BatcherLog.mdb;Jet OLEDB:Database Password=cepher;";
		 // Call Create instance to instantiate the Record set
		CString sql = "SELECT * FROM Feedback";
		hr = pRst.CreateInstance(__uuidof(Recordset));
		if(FAILED(hr))
		{            
			PrintOut("Failed creating batcher log record set instance",0);
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
		str = "Error WriteLog: " + GetCString(ce.Description());
		PrintOut(str,0);	
		
	}
	///////////////////////////////////////////////////////////////////
}

bool CProcessThread::CheckPop3Status(void)
{
	bool success = false;
	try
	{
		WriteLog("Check Pop3 Status");
		CSettings mySettings;
		mySettings.ReadSettings();
		//PrintOut(mySettings.GetPop3());
		//Pop3Socket *pSocket = new Pop3Socket( "216.219.254.138", 110);
		Pop3Socket *pSocket = new Pop3Socket( mySettings.GetPop3().GetBuffer(), 110);
		CString response = pSocket->GetResponse();
		response.Replace("\r\n","");
		//PrintOut(response);
		if(response.Left(3).CompareNoCase("+ok")!=0)
		{
			//connection problem
			PrintOut("Connection error");
			pSocket->Close();
		}
		else
		{
			PrintOut("Pop3 Connected");
			//CString username = "user test@country-loan.com\r\n";
			CString username = "user ";
			username = username + mySettings.GetUsername();
			username = username + "\r\n";
			//PrintOut(username);
			int iSent = pSocket->Write(username.GetBuffer(), username.GetLength());
			response = pSocket->GetResponse();
			response.Replace("\r\n","");
			//PrintOut(response);
			if(response.Left(3).CompareNoCase("+ok")!=0)
			{
				//user problem
				PrintOut("Username error, closing connection");
				pSocket->Close();
			}
			else
			{
				//PrintOut("User OK");
				//CString password = "pass test123\r\n";
				CString password = "pass ";
				password = password + mySettings.GetPassword();
				password =password + "\r\n";
				int iSent = pSocket->Write(password.GetBuffer(), password.GetLength());
				response = pSocket->GetResponse();
				response.Replace("\r\n","");
				//PrintOut(response);
				if(response.Left(3).CompareNoCase("+ok")!=0)
				{
					//pass problem
					PrintOut("Password problem, closing connection");
					pSocket->Close();
				}
				else
				{
					PrintOut("Pop3 Logged in");
					CString status = "stat\r\n";
					int iSent = pSocket->Write(status.GetBuffer(), status.GetLength());
					response = pSocket->GetResponse();
					response.Replace("\r\n","");
					
					//PrintOut(response);
					if(response.Left(3).CompareNoCase("+ok")!=0)
					{
						//stat problem
						PrintOut("status problem");
						pSocket->Close();
					}
					else
					{
						CArray<CString,CString> ptArray;

						//PrintOut("Status OK");
						CString resToken;
						int curPos= 0;
						response.Replace("\r\n","");
						resToken= response.Tokenize(" ",curPos);
						while (resToken != "")
						{
							ptArray.Add(resToken);
							//PrintOut(resToken);
							resToken= response.Tokenize(" ",curPos);
						};
						if(ptArray.GetCount()==3)
						{
							PrintOut("There are [" + ptArray[1] + "] messages for a total of [" + ptArray[2] + "] bytes...");
							success = true;
							WriteLog("Pop3 Status OK");							
						}
						pSocket->Close();

					}
				}
			}
		}
	}
	catch( CSocketException se)
	{
		success = false;
		CString str = " : Exception => ";
		str = str + se.getText();
		PrintOut(str);
		
		return success;
	}
	return success;
}

void CProcessThread::PrintOut(CString str)
{	
	try
	{
		PrintOut(str, 0);
	}
	catch(...)
	{
		//error trap template
		CString err;
		err = "Error: CProcessThread::PrintOut(CString str)";
		PrintOut(err);
		WriteLog(err);
	}
}

void CProcessThread::SendStatsEmail(void)
{
	try
	{
		CStats stats;
		stats.SetParentWindow(mainWin);			
		stats.ReadTodaysLeadsSent();
		stats.ReadTodaysLeadsOrdered();
		stats.ReadTodaysZerosSent();
		int maxLength = MaxLength(stats.GetCountVerified(), stats.GetCountInternet(), stats.GetCountOnePass(), stats.GetCountTotal());
		if(maxLength < 4) maxLength = 4;
		int lengthOrdered = MaxLength(stats.GetCountVerifiedOrdered(), stats.GetCountInternetOrdered(), stats.GetCountOnePassOrdered(), stats.GetCountTotalOrdered());	
		if(lengthOrdered < 7) lengthOrdered = 7;
		
		CString body;
		body = "Todays Stats\r\n\r\nTotal Leads by type:\r\n";
		body = body + "          " + SpaceFill("Sent", maxLength) + " / " + SpaceFill("Ordered", lengthOrdered) + " = Yield\r\n";
		body = body + "*******************************************\r\n";
		
		
		body = body + "Verified :";
		body = body + SpaceFill(stats.GetCountVerified(), maxLength);		
		body = body + " / " + SpaceFill(stats.GetCountVerifiedOrdered(), lengthOrdered)  + " = " + stats.GetYieldVerified() + "\r\n";
		
		body = body + "Internet :";
		body = body + SpaceFill(stats.GetCountInternet(), maxLength);		
		body = body + " / " + SpaceFill(stats.GetCountInternetOrdered(), lengthOrdered) + " = " + stats.GetYieldInternet() + "\r\n";
		
		body = body + "One Pass :";
		body = body + SpaceFill(stats.GetCountOnePass(), maxLength);		
		body = body + " / " + SpaceFill(stats.GetCountOnePassOrdered(), lengthOrdered) + " = " + stats.GetYieldOnePass() + "\r\n";
		
		body = body + "Total    :";
		body = body + SpaceFill(stats.GetCountTotal(), maxLength);
		
		body = body + " / " + SpaceFill(stats.GetCountTotalOrdered(), lengthOrdered) + " = " + stats.GetYieldTotal() + "\r\n";
		
		body = body + "*******************************************\r\n\r\n\r\n";
		int camLength = MaxLength(stats.GetSentZero(), stats.GetSentPartials(), stats.GetSentAll(), "Count");
		if(camLength < 4) camLength = 4;
		body = body + "Today’s Verified Campaign Stats\r\n";
		body = body + "                         " + SpaceFill("Count", camLength) + " / Total = Percent\r\n";
		body = body + "************************************************\r\n";
		body = body + "Campaigns Sent zero    : " + SpaceFill(stats.GetSentZero(), camLength) + " / " + stats.GetTotalCampaigns() + "   = " + stats.GetPercentZero() + "\r\n";
		body = body + "Campaigns Sent Partials: " + SpaceFill(stats.GetSentPartials(), camLength) + " / " + stats.GetTotalCampaigns() + "   = " + stats.GetPercentPartials() + "\r\n";
		body = body + "Campaigns Sent All     : " + SpaceFill(stats.GetSentAll(), camLength) + " / " + stats.GetTotalCampaigns() + "   = " + stats.GetPercentAll() + "\r\n";
		body = body + "The Median Delivery: " + stats.GetMean() + "\r\n";
		body = body + "************************************************\r\n";

		PrintOut("Sent / Ordered = Yield");
		CString str;
		str = "Verified    :" + stats.GetCountVerified();
		str = str + " / " + stats.GetCountVerifiedOrdered() + " = " + stats.GetYieldVerified();
		PrintOut(str);
		PrintOut("Internet   :" + stats.GetCountInternet() + " / " + stats.GetCountInternetOrdered() + " = " + stats.GetYieldInternet());
		PrintOut("One Pass :" + stats.GetCountOnePass() + " / " + stats.GetCountOnePassOrdered() + " = " + stats.GetYieldOnePass());
		PrintOut("Total        :" + stats.GetCountTotal() + " / " + stats.GetCountTotalOrdered() + " = " + stats.GetYieldTotal());
		PrintOut("*******************************************");
		PrintOut("Today’s Campaign Stats");
		PrintOut("Count / Total Verified Campaigns = Percent"); 
		
		PrintOut("Campaigns Sent zero: " + stats.GetSentZero() + " / " + stats.GetTotalCampaigns() + " = " + stats.GetPercentZero());
		PrintOut("Campaigns Sent Partials: " + stats.GetSentPartials() + " / " + stats.GetTotalCampaigns() + "   = " + stats.GetPercentPartials());
		PrintOut("Campaigns Sent All     : " + stats.GetSentAll() + " / " + stats.GetTotalCampaigns() + "   = " + stats.GetPercentAll());
		PrintOut("The Median Delivery: " + stats.GetMean());
		PrintOut("*******************************************");

		ESMTP mail;
		mail.SetParentWindow(mainWin);
		//mail.SetName("Testing");
		mail.SetName("Batch Process");
		mail.SetAddress("test@country-loan.com");
		//CString to = "kenco_computers@yahoo.com";
		//CString to = "test@country-loan.com";
		///*money@e-realtyloans.com*/
		CString to = "tomer9876@yahoo.com;kenco_computers@yahoo.com;avi@country-loan.com;";
		mail.SetTo(to);
		mail.SetBCC("test@country-loan.com");
		mail.SetSubject("Todays Leads Sent Stats");	
		mail.SetHost("216.219.253.246");
		mail.SetUserName("test@country-loan.com");
		mail.SetPassword("test123");
		//PrintOut(mySettings.GetSmtp());
		//mail.SetHost(mySettings.GetSmtp());
		//mail.SetUserName(mySettings.GetUsername());
		//mail.SetPassword(mySettings.GetPassword());
		mail.SetBody(body);
		//CString filename = myLoc.GetDesDir() + "\\";
		//filename = filename + orderid;
		//filename = filename + ".rtf";
		//mail.AddFilename(filename);

		PrintOut("Sending Stats Email");
		//mail.SetDiv(0);
		if(mail.Send())
		{
			PrintOut("Email sent to: " + mail.GetTo() + " and BCC to: " + mail.GetBCC());
			
			//success = true;
			//AfxMessageBox("Message send to: " + mail.m_sTo , MB_OK);
		}
		else
		{
			PrintOut("Error sending Email. Make sure you are connected to the internet.");
			//AfxMessageBox("Error sending Email. Make sure you are connected to the internet before hitting Send." , MB_OK);
			//success = false;
		}
		if(mainWin)
		{		
			SendMessage(mainWin, PO_DONE, 0, 0);		
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CProcessThread::SendStatsEmail";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

int CProcessThread::MaxLength(CString str1, CString str2, CString str3, CString str4)
{
	try
	{
		int maxLength = str1.GetLength();
		
		if(str2.GetLength() > maxLength)
		{
			maxLength = str2.GetLength();
		}
		if(str3.GetLength() > maxLength)
		{
			maxLength = str3.GetLength();
		}
		if(str4.GetLength() > maxLength)
		{
			maxLength = str4.GetLength();
		}
		return maxLength;

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CProcessThread::MaxLength(CString str1, CString str2, CString str3, CString str4)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return 0;
	}
}

CString CProcessThread::SpaceFill(CString str, int iLength)
{
	try
	{
		CString filledStr = str;
		int x = iLength - str.GetLength();
		while(x > 0)
		{
			filledStr = filledStr + " ";
			x--;
		}
		return filledStr;

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CProcessThread::SpaceFill(CString str, int iLength)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}
