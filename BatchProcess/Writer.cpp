#include "StdAfx.h"
#include ".\writer.h"

CWriter::CWriter(CString target, HWND m_hwnd)
: desDir(_T(""))
{
	mainWin = m_hwnd;
	desDir = target;
	
	if( _mkdir( desDir ) == 0 )
	{
		//PrintOut(desDir + " directory created");
	}
	else
	{
		//PrintOut(desDir + " directory already exists");
	}
}

CWriter::~CWriter(void)
{
}

void CWriter::PrintOut(CString str)
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
		//error trap
		CString err;
		err = "Error: CWriter::PrintOut(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CWriter::Write(COrder aOrder)
{
}
long CWriter::WriteString(CFile* pt_target, CString str)
{
	try
	{
		char letter;
		long byteswrite = 0;
		for(long x = 0; x < str.GetLength(); x++)
		{
			letter = str.GetAt(x);
			pt_target->Write(&letter, 1);
			byteswrite++;
		}
		return byteswrite;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CWriter::WriteString(CFile* pt_target, CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	
}
void CWriter::Progress(int iPos, int iMax)
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
		err = "Error: CWriter::Progress(int iPos, int iMax)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CWriter::WriteLog(CString str)
{
	CWriteLog::Write(str);
	///////////////////////////////////////////////////////////////////
	/*HRESULT hr;
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
		
		str = "Error WriteLog: " + CStringHelper::GetCString(ce.Description());
		PrintOut(str);	
		
	}*/
	///////////////////////////////////////////////////////////////////
}
