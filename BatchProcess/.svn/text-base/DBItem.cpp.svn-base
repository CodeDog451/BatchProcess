#include "StdAfx.h"
#include ".\dbitem.h"

DBItem::DBItem(void)
{
}

DBItem::~DBItem(void)
{
}

bool DBItem::GetFieldBool(_RecordsetPtr rs, CString field)
{
	try
	{
		bool val;
		_variant_t vtFld;
		if(vtFld.vt == VT_NULL)
		{
			return false;
		}
		
		vtFld = rs->Fields->GetItem(field.GetBuffer())->Value;
		val = vtFld.boolVal;
		return val;
	}
	catch(...)
	{
		return false;
	}
}
COleDateTime DBItem::GetFieldDate(_RecordsetPtr rs, CString field)
{
	try
	{
		COleDateTime val;
		_variant_t vtFld;
		if(vtFld.vt == VT_NULL)
		{
			return COleDateTime();
		}
		vtFld = rs->Fields->GetItem(field.GetBuffer())->Value;
		val = vtFld.date;
		return val;
	}
	catch(...)
	{
		return COleDateTime();
	}
}
double DBItem::GetFieldDouble(_RecordsetPtr rs, CString field)
{
	try
	{
		double val;
		_variant_t vtFld;
		
		vtFld = rs->Fields->GetItem(field.GetBuffer())->Value;
		

		if(vtFld.vt == VT_NULL)
		{
			return -1;
		}
		else
		{
			val = vtFld.dblVal;
			return val;
		}
	}
	catch(...)
	{
		return -1.00;
	}
}
long DBItem::GetFieldLong(_RecordsetPtr rs, CString field)
{
	try
	{
		long val;
		_variant_t vtFld;
		vtFld = rs->Fields->GetItem(field.GetBuffer())->Value;
		
		if(vtFld.vt == VT_NULL)
		{
			return -1;
		}
		else
		{
			val = vtFld.lVal;
			return val;
		}
		
	}
	catch(...)
	{
		return -1;
	}
}
CString DBItem::GetFieldString(_RecordsetPtr rs, CString field)
{
	try
	{
		_bstr_t val;
		_variant_t vtFld;
		vtFld = rs->Fields->GetItem(field.GetBuffer())->Value;
		
		
		if(vtFld.vt == VT_NULL)
		{
			return CString();
		}
		else
		{
			val = vtFld.bstrVal;
			return GetCString(val);
		}
	}
	catch(...)
	{
		return "";
	}
}
CString DBItem::GetCString(_bstr_t bstrStart)
{
	return CStringHelper::GetCString(bstrStart);
	/*try
	{	
		if(bstrStart.length()>0)
		{
			int length = bstrStart.length();
			TCHAR *szFinal = new TCHAR[length+1];
			//TCHAR szFinal[length+1];
			//TCHAR szFinal[1024];
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
		err = "Error: DBItem::GetCString(_bstr_t bstrStart)";
		AfxMessageBox(err);
		//PrintOut(err);
		//WriteLog(err);
	}*/
}