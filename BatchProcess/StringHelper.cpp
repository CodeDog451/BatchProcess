#include "StdAfx.h"
#include ".\stringhelper.h"

CStringHelper::CStringHelper(void)
{
}

CStringHelper::~CStringHelper(void)
{
}

CString CStringHelper::GetCString(_bstr_t bstrStart)
{
	//convert bstring to cstring
	try
	{	
		if(bstrStart.length()>0)
		{
			int length = bstrStart.length();
			TCHAR *szFinal = new TCHAR[length+1];			
			_stprintf(szFinal, _T("%s"), (LPCTSTR)bstrStart);
			CString str;
			str = szFinal;
			delete szFinal;
			return str;
		}
		else
		{
			return "";
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStringHelper::GetCString(_bstr_t bstrStart)";
		CWriteLog::WriteFail(err);
		AfxMessageBox(err);
		return CString();
		
	}
}

CString CStringHelper::GetCString(long iNum)
{
	//conver long to cstring
	try
	{
		char buffer[200];
		sprintf(buffer, "%d", iNum);
		return CString(buffer);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStringHelper::GetCString(long iNum)";
		
		CWriteLog::WriteFail(err);
		AfxMessageBox(err);
	}
}

CString CStringHelper::GetCString(double fNum)
{
	//convert double to cstring
	try
	{
		char buffer[200];
		sprintf(buffer, "%f", fNum);
		return CString(buffer);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStringHelper::GetCString(double fNum)";
		
		CWriteLog::WriteFail(err);
		AfxMessageBox(err);
		return CString();
	}
}


CString CStringHelper::GetCString(bool state)
{
	// convert bool to cstring
	try
	{
		CString str;
		if(state)
		{
			str = "true";
		}
		else
		{
			str = "false";
		}
		return str;	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStringHelper::GetCString(bool state)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

CString CStringHelper::GetCString(double fNum, int iPer)
{
	try
	{
		CString str;
		char buffer[200];
		_gcvt(fNum, iPer, buffer);
		str = buffer;
		return str;

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: GetCString";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}
