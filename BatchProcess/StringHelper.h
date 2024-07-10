#pragma once
#import "C:\Program Files\Common Files\System\ADO\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")
#include ".\writelog.h"
class CStringHelper
{
public:
	CStringHelper(void);
	~CStringHelper(void);
	static CString GetCString(_bstr_t bstrStart);
	static CString GetCString(long iNum);
	static CString GetCString(double fNum);
	// convert bool to cstring
	static CString GetCString(bool state);
	static CString GetCString(double fNum, int iPer);
};
