#pragma once
#include <conio.h>
#include <iostream>
//#include "logdb.h"
#include ".\writelog.h"
//#include "fxscom.h"
#import "c:\windows\system32\fxscom.dll"
//#import "c:\windows\system32\fxscomex.dll" rename ("GetJob","GetFaxJob") rename ("GetMessage","GetFaxMessage")
using namespace std;

//a fax sending class
class CFax
{
public:
	CFax(void);
	~CFax(void);
	bool Send(void);
protected:
	HWND mainWin;
public:
	void SetParentWindow(HWND m_hwnd);
	void PrintOut(CString str);
protected:
	CString filepath;
	CString coverpage;
	CString coverpagenote;
	CString faxnumber;
public:
	void SetCoverPage(CString filename);
	void SetCoverPageNote(CString note);
	void SetFaxNumber(CString phone);
	void SetBody(CString filename);
	CString GetFaxNumber(void);
	
protected:
	int iSendCoverpage;
public:
	void SetSendCoverpage(int iNum);
protected:
	//CLogDB myLog;
};
