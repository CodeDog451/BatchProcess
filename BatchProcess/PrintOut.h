#pragma once
#include ".\writelog.h"
#include "resource.h"
class CPrintOut
{
public:
	CPrintOut(void);
	~CPrintOut(void);
protected:
	// handle to main app window
	HWND mainWin;
public:
	void SetParentWindow(HWND m_wnd);
	void Print(CString str);
};
