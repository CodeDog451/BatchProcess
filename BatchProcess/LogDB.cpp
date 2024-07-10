#include "StdAfx.h"
#include ".\logdb.h"

CLogDB::CLogDB(void)
{
	
	try
    {
		CoInitialize(NULL);
		strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BatcherLog.mdb;Jet OLEDB:Database Password=cepher;";
		
		pRst = NULL;
		Connect(pRst, strCnn, "SELECT * FROM Feedback");
		
		pRstFail = NULL;
		Connect(pRstFail, strCnn, "SELECT * FROM FailLog");

		//pRstFax = NULL;
		//Connect(pRstFax, strCnn, "SELECT * FROM FaxFeedback");
		
		//pRstFaxFail = NULL;
		//Connect(pRstFaxFail, strCnn, "SELECT * FROM FaxFailLog");
		      
      
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Error CLogDB init: " + GetCString(ce.Description());
		AfxMessageBox(str);		
		
	}
}

CLogDB::~CLogDB(void)
{
	CoUninitialize(); 
}

void CLogDB::WriteLog(CString str, int div)
{
	Write(pRst, str, div);
	/*try
    {
		COleDateTime aTime;
		aTime = COleDateTime::GetCurrentTime();
		
		pRst->AddNew();
		pRst->Fields->Item["div"]->Value = div;
		pRst->Fields->Item["Date"]->Value = _variant_t(aTime);	
		pRst->Fields->Item["Message"]->Value = str.GetBuffer();	
		pRst->Update();
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Error Write Log: " + GetCString(ce.Description());
		AfxMessageBox(str);		
		
	}*/
}

void CLogDB::WriteFail(CString str, int div)
{
	Write(pRstFail, str, div);
	/*try
    {
		COleDateTime aTime;
		aTime = COleDateTime::GetCurrentTime();
		
		pRstFail->AddNew();
		pRstFail->Fields->Item["div"]->Value = div;
		pRstFail->Fields->Item["Date"]->Value = _variant_t(aTime);	
		pRstFail->Fields->Item["Message"]->Value = str.GetBuffer();	
		pRstFail->Update();
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Error Write Fail: " + GetCString(ce.Description());
		AfxMessageBox(str);		
		
	}*/
}

bool CLogDB::Connect(_RecordsetPtr &pt_Rst, _bstr_t strCon, CString sql)
{
	bool success = false;
	///////////////////////////////////////////////////////////////////
	HRESULT hr;
	hr = S_OK;	
	
	try
    {		
		pt_Rst = NULL;
		 // Call Create instance to instantiate the Record set
		hr = pt_Rst.CreateInstance(__uuidof(Recordset));
		if(FAILED(hr))
		{            
			AfxMessageBox("Failed creating batcher log record set instance");
		}
		else
		{			
			pt_Rst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);
			success = true;
		}	      
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "Error Connect: " + GetCString(ce.Description());
		AfxMessageBox(str);	
		AfxMessageBox(sql);	
	}
	///////////////////////////////////////////////////////////////////
	return success;
}

//void CLogDB::WriteFaxLog(CString str, int div)
//{
	//Write(pRstFax, str, div);
//}

//void CLogDB::WriteFaxFail(CString str, int div)
//{
	//Write(pRstFaxFail, str, div);
//}

void CLogDB::Write(_RecordsetPtr &pt_Rst, CString str, int div)
{
	try
    {
		COleDateTime aTime;
		aTime = COleDateTime::GetCurrentTime();
		
		pt_Rst->AddNew();
		pt_Rst->Fields->Item["div"]->Value = div;
		pt_Rst->Fields->Item["Date"]->Value = _variant_t(aTime);	
		pt_Rst->Fields->Item["Message"]->Value = str.GetBuffer();	
		pt_Rst->Update();
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Error Write: " + GetCString(ce.Description());
		AfxMessageBox(str);		
		
	}
}
