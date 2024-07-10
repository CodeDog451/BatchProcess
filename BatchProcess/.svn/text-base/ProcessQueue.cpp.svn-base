#include "StdAfx.h"
#include ".\processqueue.h"

CProcessQueue::CProcessQueue(void)
{
}

CProcessQueue::~CProcessQueue(void)
{
}
CProcessQueue& CProcessQueue::operator=(const CProcessQueue& aCProcessQueue)
{
	this->myJobs.Copy(aCProcessQueue.myJobs);	

	return *this;
}
CProcessQueue::CProcessQueue(const CProcessQueue& aCProcessQueue)
{
	
	this->myJobs.Copy(aCProcessQueue.myJobs);	

}
bool CProcessQueue::ReadJobs(void)
{

	bool success = false;
	HRESULT hr = S_OK;
    try
    {
		myJobs.RemoveAll();
		//PrintOut("Init database connection read locations");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BatcherSettings.mdb;Jet OLEDB:Database Password=cepher;");

       _RecordsetPtr pRst = NULL;

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {
            
		  PrintOut("Failed creating locations record set instance");
		 
            
      }
	  else
	  {
		
		
		CString sql;
		sql = "SELECT * FROM ProcessOrder"; 
		
		
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);
		
		if (!pRst->EndOfFile)
		{
			pRst->MoveFirst();

			while(!pRst->EndOfFile)
			{
				CProcessJob aJob;
				aJob.SetCategory(GetFieldLong(pRst, "LeadType"));
				aJob.SetGroup(GetFieldLong(pRst, "OrderType"));
				aJob.SetState(GetFieldLong(pRst, "iStateID_fk"));
				//PrintOut(CStringHelper::GetCString(GetFieldLong(pRst, "iStateID_fk")));
				myJobs.Add(aJob);

				pRst->MoveNext();
			}		
			
			
		}
	  }
	  //CoUninitialize();  
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Location Error: " + GetCString(ce.Description());
		PrintOut(str);
		
		
		
	}
	
  
	return success;
}

void CProcessQueue::SetParentWindow(HWND m_hwnd)
{
	mainWin = m_hwnd;
}

void CProcessQueue::PrintOut(CString str)
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

CProcessJob CProcessQueue::GetJobAt(int index)
{
	return myJobs.GetAt(index);
}

int CProcessQueue::GetCount(void)
{
	return myJobs.GetCount();
}

void CProcessQueue::DeleteAt(int index)
{
	myJobs.RemoveAt(index);
}

CProcessJob CProcessQueue::GetNextJob(void)
{
	if(myJobs.GetCount()>0)
	{
		CProcessJob aJob;
		aJob = GetJobAt(0);
		DeleteAt(0);
		return aJob;
	}
	else
	{
		return CProcessJob();
	}
}
