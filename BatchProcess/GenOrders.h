#pragma once
#include "dbitem.h"
#include "resource.h"
#include ".\writelog.h"
class CGenOrders :
	public DBItem
{
public:
	CGenOrders(void);
	~CGenOrders(void);
	CString GetSQL(void);
	bool ReadCampaigns(void);
protected:
	HWND mainWin;
public:
	void SetParentWindow(HWND m_hwnd);
	void PrintOut(CString str);
	void Progress(int iPos, int iMax);
	CString GetString(long iNum);
	CString GetBool_String(bool state);
	CString GetString(double fNum);
	void WriteLog(CString str);
};
