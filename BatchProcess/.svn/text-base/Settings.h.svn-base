#pragma once
#include "stdafx.h"
#include ".\location.h"
#include "logdb.h"
#import "C:\Program Files\Common Files\System\ADO\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")
class CSettings
{
public:
	CSettings(void);
	~CSettings(void);
	
protected:
	long GetFieldLong(_RecordsetPtr rs, CString field);
	HWND mainWin;
public:
	void SetParentWindow(HWND m_hwnd);
	void PrintOut(CString str);
	CString GetCString(_bstr_t bstrStart);

	void SetSendFax(bool state);
	bool GetSendFax(void);
	bool GetFieldBool(_RecordsetPtr rs, CString field);
protected:
	//CLogDB myLog;
	bool sendfax;
	long timeout;
public:
	void ReadSettings(void);
	void WriteSettings(void);
	long GetTimeout(void);
	void SetTimeout(long time);
protected:
	bool sendEmail;
	CString pop3;
	CString smtp;
	CString username;
	CString password;
	CArray<CLocation,CLocation&> locations;
public:
	CString GetFieldString(_RecordsetPtr rs, CString field);
	bool GetSendEmail(void);
	void SetSendEmail(bool state);
	void SetPop3(CString str);
	void SetSmtp(CString str);
	CString GetPop3(void);
	CString GetSmtp(void);
	void SetUsername(CString str);
	CString GetUsername(void);
	void SetPassword(CString str);
	CString GetPassword(void);

	void ReadLocations(void);
	CLocation GetLocation(int loc);
};
	
