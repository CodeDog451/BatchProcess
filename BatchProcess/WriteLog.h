#pragma once
#import "C:\Program Files\Common Files\System\ADO\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")
#include ".\stringhelper.h"
class CWriteLog
{
public:
	CWriteLog(void);
	~CWriteLog(void);
	static void WriteFail(CString str);
	static void Write(CString str);
	static void PrintOut(CString str);
};
