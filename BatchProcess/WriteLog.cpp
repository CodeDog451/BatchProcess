#include "StdAfx.h"
#include ".\writelog.h"

CWriteLog::CWriteLog(void)
{
}

CWriteLog::~CWriteLog(void)
{
}

void CWriteLog::WriteFail(CString str)
{
	///////////////////////////////////////////////////////////////////
	
	
	try
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
			PrintOut("Failed creating batcher log record set instance");
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
		
		str = "Error CWriteLog::WriteFail(CString str)): " + CStringHelper::GetCString(ce.Description());
		PrintOut(str);	
		
	}
	///////////////////////////////////////////////////////////////////
}

void CWriteLog::Write(CString str)
{
	///////////////////////////////////////////////////////////////////
	
	
	try
    {
		HRESULT hr;
		hr = S_OK;	
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
			PrintOut("Failed creating batcher log record set instance");
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
		
		str = "Error CWriteLog::Write(CString str): " + CStringHelper::GetCString(ce.Description());
		PrintOut(str);	
		
	}
	///////////////////////////////////////////////////////////////////
}

void CWriteLog::PrintOut(CString str)
{
	AfxMessageBox(str);
}
