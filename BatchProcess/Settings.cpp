#include "StdAfx.h"
#include ".\settings.h"
#include "resource.h"
CSettings::CSettings(void)
: sendfax(false)
, timeout(0)
, sendEmail(false)
, pop3(_T(""))
, smtp(_T(""))
, username(_T(""))
, password(_T(""))
{	
	
}

CSettings::~CSettings(void)
{
}



long CSettings::GetFieldLong(_RecordsetPtr rs, CString field)
{
	try
	{
		long val;
		val = rs->Fields->GetItem(field.GetBuffer())->Value.lVal;
		return val;
	}
	catch(...)
	{
		return -1;
	}
}

void CSettings::SetParentWindow(HWND m_hwnd)
{
	mainWin = m_hwnd;
}

void CSettings::PrintOut(CString str)
{
	//myLog.WriteLog(str, 0);
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

CString CSettings::GetCString(_bstr_t bstrStart)
{
	return CStringHelper::GetCString(bstrStart);
	/*int x = bstrStart.length()+1;
	TCHAR* pt_szFinal = new TCHAR[x];
	//TCHAR szFinal[x];
	_stprintf(pt_szFinal, _T("%s"), (LPCTSTR)bstrStart);
	CString str;
	str = pt_szFinal;
	return str;*/
}

void CSettings::SetSendFax(bool state)
{
	sendfax = state;
	//WriteSettings();
}

bool CSettings::GetSendFax(void)
{
	//ReadSettings();
	return sendfax;
}

bool CSettings::GetFieldBool(_RecordsetPtr rs, CString field)
{
	try
	{
		bool val;
		_variant_t vtFld;
		
		vtFld = rs->Fields->GetItem(field.GetBuffer())->Value;
		if(vtFld.vt == VT_NULL)
		{
			return false;
		}
		val = vtFld.boolVal;
		return val;
	}
	catch(...)
	{
		return false;
	}
}

void CSettings::ReadSettings(void)
{
	
	
    try
    {
		HRESULT hr = S_OK;
		//PrintOut("Init database connection read  setting");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BatcherSettings.mdb;Jet OLEDB:Database Password=cepher;");

       _RecordsetPtr pRst = NULL;

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {
            
		  PrintOut("Failed creating settings record set instance");
		 
            
      }
	  else
	  {
		
		
		CString sql;
		sql = "SELECT * FROM main"; 
		
		
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);
		
		if (!pRst->EndOfFile)
		{
			pRst->MoveFirst();
			
			sendfax = GetFieldBool(pRst, "sendfax");
			sendEmail = GetFieldBool(pRst, "sendemail");
			timeout = GetFieldLong(pRst, "timeout");
			password = GetFieldString(pRst, "password");
			pop3 = GetFieldString(pRst, "pop3");
			smtp = GetFieldString(pRst, "smtp");
			username = GetFieldString(pRst, "username");
			
			
		}
	  }
	  CoUninitialize();  
	  
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Error: " + GetCString(ce.Description());
		PrintOut(str);
		//myLog.WriteFail(str, 0);
		
		
	}
	
  
	
	
}



void CSettings::WriteSettings(void)
{
	HRESULT hr = S_OK;
    try
    {
		//PrintOut("Init database connection write setting");
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=settings.mdb;Jet OLEDB:Database Password=cepher;");

       _RecordsetPtr pRst = NULL;

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {
            
		  PrintOut("Failed creating settings record set instance");		  
            
      }
	  else
	  {		
		 
		CString sql;
		sql = "SELECT * FROM main"; 		
		
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);
		
		if (!pRst->EndOfFile)
		{
			pRst->MoveFirst();				
			pRst->Fields->Item["sendfax"]->Value = sendfax;
			pRst->Fields->Item["sendemail"]->Value = sendEmail;
			//pRst->Fields->Item["trigger"]->Value = trigger;
			pRst->Fields->Item["timeout"]->Value = timeout;
			pRst->Fields->Item["password"]->Value = password.GetBuffer();
			
			pRst->Fields->Item["username"]->Value = username.GetBuffer();
			pRst->Fields->Item["pop3"]->Value = pop3.GetBuffer();
			pRst->Fields->Item["smtp"]->Value = smtp.GetBuffer();
			
			pRst->Update();
			
			
		}
	  }
	  CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Error: " + GetCString(ce.Description());
		PrintOut(str);
		//myLog.WriteFail(str, 0);
		
	}
	
    
}

long CSettings::GetTimeout(void)
{
	//ReadSettings();
	return timeout;
}

void CSettings::SetTimeout(long time)
{
	timeout = time;
	//WriteSettings();
}

CString CSettings::GetFieldString(_RecordsetPtr rs, CString field)
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

bool CSettings::GetSendEmail(void)
{
	//ReadSettings();
	return sendEmail;
}

void CSettings::SetSendEmail(bool state)
{
	
	sendEmail = state;
	//WriteSettings();
}

void CSettings::SetPop3(CString str)
{
	pop3 = str;
	//WriteSettings();
}

void CSettings::SetSmtp(CString str)
{
	smtp = str;
	//WriteSettings();
}

CString CSettings::GetPop3(void)
{
	//ReadSettings();
	return pop3;
}

CString CSettings::GetSmtp(void)
{
	//ReadSettings();
	return smtp;
}

void CSettings::SetUsername(CString str)
{
	//PrintOut(str);
	username = str;
	//WriteSettings();
}

CString CSettings::GetUsername(void)
{
	//ReadSettings();
	return username;
}

void CSettings::SetPassword(CString str)
{
	password = str;
	//WriteSettings();
}

CString CSettings::GetPassword(void)
{
	//ReadSettings();
	return password;
}

void CSettings::ReadLocations(void)
{
	HRESULT hr = S_OK;
    try
    {
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
		sql = "SELECT * FROM location"; 
		
		
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);
		
		if (!pRst->EndOfFile)
		{
			pRst->MoveFirst();

			while(!pRst->EndOfFile)
			{
				CLocation loc;
				loc.SetDiv(GetFieldLong(pRst, "div"));
				loc.SetReturnMail(GetFieldString(pRst, "defaultEmail"));				
				loc.SetSubject(GetFieldString(pRst, "subject"));				
				loc.SetBody(GetFieldString(pRst, "body"));
				loc.SetZeroBody(GetFieldString(pRst, "zeroBody"));
				loc.SetDesDir(GetFieldString(pRst, "desDir"));
				loc.SetTemplateDir(GetFieldString(pRst, "templateDir"));
				loc.SetFilterStatement(GetFieldString(pRst, "filterStatement"));
				//PrintOut("note: " + GetFieldString(pRst, "emailNoteVerified"));
				loc.SetEmailNoteVerifiedPage(GetFieldString(pRst, "emailNoteVerified"));
				loc.SetEmailNoteNonVerifiedPage(GetFieldString(pRst, "emailNoteNotVerified"));
				loc.SetAgeLimitVerified(GetFieldLong(pRst, "ageLimitVerified"));
				loc.SetUseLimitVerified(GetFieldLong(pRst, "useLimitVerified"));

				loc.SetAgeLimitInternet(GetFieldLong(pRst, "ageLimitInternet"));
				loc.SetUseLimitInternet(GetFieldLong(pRst, "useLimitInternet"));

				loc.SetAgeLimitOnePass(GetFieldLong(pRst, "ageLimitOnePass"));
				loc.SetUseLimitOnePass(GetFieldLong(pRst, "useLimitOnePass"));
				
				locations.Add(loc);
				pRst->MoveNext();
			}		
			
			
		}
	  }
	   CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Location Error: " + GetCString(ce.Description());
		PrintOut(str);
		//myLog.WriteFail(str, 0);
		
		
	}
	
   
	
}



CLocation CSettings::GetLocation(int loc)
{
	CLocation aLoc;
	int index = loc-1;
	if(locations.GetCount()>0)
	{
		if(index < 0) index = 0;
		if(index > (locations.GetCount()-1)) index = locations.GetCount()-1;
		aLoc = locations.GetAt(index);
	}
	
	return aLoc;
}
