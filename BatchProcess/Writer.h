#pragma once
#include "resource.h"
#include <direct.h>
#include ".\order.h"
#include ".\settings.h"
#include ".\stringhelper.h"
#include ".\writelog.h"
class CWriter
{
public:
	CWriter(CString target, HWND m_hwnd);
	~CWriter(void);
	void PrintOut(CString str);
	virtual void Write(COrder aOrder);
protected:
	CString desDir;
	HWND mainWin;
	long WriteString(CFile* pt_target, CString str);
	//CLogDB myLog;
public:
	void Progress(int iPos, int iMax);
	void WriteLog(CString str);
};
