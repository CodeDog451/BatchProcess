#pragma once
#include "hlistbox.h"
#import "C:\Program Files\Common Files\System\ADO\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")
#include ".\stringhelper.h"
class DBItem
{
public:
	DBItem(void);
	~DBItem(void);
protected:
	bool GetFieldBool(_RecordsetPtr rs, CString field);
	COleDateTime GetFieldDate(_RecordsetPtr rs, CString field);
	double GetFieldDouble(_RecordsetPtr rs, CString field);
	long GetFieldLong(_RecordsetPtr rs, CString field);
	CString GetFieldString(_RecordsetPtr rs, CString field);
	CString GetCString(_bstr_t bstrStart);
};
