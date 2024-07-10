/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD DERIVED CLASS - produced by Worker Thread Class Generator Rel. 1.0
// ProcessOrders.cpp: Implementation of the CProcessOrders Class.
/////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ProcessOrders.h"
#include ".\processorders.h"
#include "resource.h"	
#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNAMIC(CProcessOrders, CThread)

/////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
/////////////////////////////////////////////////////////////////////////////////////
CProcessOrders::CProcessOrders(void* pOwnerObject /*= NULL*/, LPARAM lParam /*= 0L*/)
	: CThread(pOwnerObject, lParam)
{
	// User Initialization here...
}

CProcessOrders::~CProcessOrders()
{}

/////////////////////////////////////////////////////////////////////////////////////
// CProcessOrders diagnostics
#ifdef _DEBUG
void CProcessOrders::AssertValid() const
{
	CThread::AssertValid();
}

void CProcessOrders::Dump(CDumpContext& dc) const
{
	CThread::Dump(dc);
}
#endif //_DEBUG
/////////////////////////////////////////////////////////////////////////////////////

// Unallocate all thread specific extra resources if needed while killing this thread
void CProcessOrders::OnKill()
{
	// Code here...

	// This line may be removed if not necessary...
	CThread::OnKill();
}

/////////////////////////////////////////////////////////////////////////////////////
// WORKER THREAD CLASS GENERATOR - Do not remove this method!
// MAIN THREAD HANDLER - The only method that must be implemented.
/////////////////////////////////////////////////////////////////////////////////////
DWORD CProcessOrders::ThreadHandler()
{
	// Trivial Thread specific code here...
	for(int x = 0; x < 5; x++)
	{
		//test thread message sending
		
		Sleep(500);
		if(mainWin)
		{
			PostMessage(mainWin, ID_LEAD_READ, x, 0);
		}
		ReadLeads();
	}
	return (DWORD)CThread::DW_OK;	//... if the thread task completion OK
}

void CProcessOrders::SetParentWindow(HWND m_hwnd)
{
	mainWin = m_hwnd;
}

long CProcessOrders::ReadLeads(void)
{
	long result = -1;
	HRESULT hr = S_OK;
    try
    {
		PrintOut(ID_INIT_DB_CONNECT);
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=whompa;");

       _RecordsetPtr pRstInternetLeads = NULL;

      // Call Create instance to instantiate the Record set
      hr = pRstInternetLeads.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {
            //printf("Failed creating record set instance\n");
		  PrintOut(FAIL_CREATE_RECORD_INSTANCE);
		  result = -1;
            //return 0;
      }
	  else
	  {
		//WHERE (((recentLeads.dupe)=False) AND ((recentLeads.age)<14));
		//Open the Record set for getting records from InternetLeads query
		pRstInternetLeads->Open("SELECT * FROM recentLeads WHERE dupe=false AND age<7 ORDER BY lead_id", strCnn, adOpenStatic, adLockReadOnly, adCmdText);

		

		pRstInternetLeads->MoveFirst();

		//Loop through the Record set
		if (!pRstInternetLeads->EndOfFile)
		{
			result = 0;
			PrintOut(START_LEAD_FILE_READ);
			CString lead_id;
			CString first_name;
			CString last_name;
			CString sys_cre_date;
			//leads.SetOutput(&m_output);
			//Lead aLead;
			CLead* aLead;//&& (result <= 10000)
			while((!pRstInternetLeads->EndOfFile))
			{				
				aLead = new CLead();
				aLead->ReadLead(pRstInternetLeads);
				//aLead.PrintOut(&m_output);
				leads.add(aLead);
				//if(aLead->GetLead_Verified())
				//{
					//verifiedLeads.add(aLead);
				//}
				if((result % 1000)==0)
				{
					//aLead.PrintOut(&m_output);
					PrintOut(ID_LEADS_COUNT, result);
					
				}
				
				result++;				
				pRstInternetLeads->MoveNext();
			}			
			
			PrintOut(ID_LEADS_COUNT, result);
		}
	  }
	  
	}
	catch(_com_error & ce)
	{
		//CString str;
		//_bstr_t temp;
		//temp = ce.Description;
		//str = getCString(temp);
		//str = "Error: " + GetCString(ce.Description());
		PrintOut(ID_COM_ERR);
		result = -1;
		//printf("Error:%s\n",ce.Description);
	}
	
  CoUninitialize();
	return result;
}

void CProcessOrders::PrintOut(int msg_id)
{
	if(mainWin)
	{
		PostMessage(mainWin, msg_id, 0, 0);
	}
}

void CProcessOrders::PrintOut(int msg_id, int count)
{
	if(mainWin)
	{
		PostMessage(mainWin, msg_id, count, 0);
	}
}
