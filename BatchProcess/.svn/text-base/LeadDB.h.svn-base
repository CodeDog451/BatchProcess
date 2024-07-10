#pragma once
#include "lead.h"
#include "afxwin.h"
#include "afxtempl.h"
#include "logdb.h"

class CLeadDB
{
public:
	CLeadDB(void);
	~CLeadDB(void);
	//CMap theLeads;
	//CMap<long,long,CPoint,CPoint&> myLeadMap(16);
	//typedef 
	CMap<long, long, CLead, CLead&> myLeadMap;
	//CMyMap aMap;

	void Add(CLead aLead);
	HWND mainWin;
	void SetParentWindow(HWND m_hwnd);
	bool GetLead(long iID, CLead &aLead);
	void PrintOut(CString str);
	void PrintLeads(void);
	
	void RemoveAll(void);
protected:
	bool stop;
	//CLogDB myLog;
public:
	void SetStop(bool state);
	CLead GetLead(long id);
	void Update(CLead aLead);
	CLeadDB& operator=(const CLeadDB& aCLeadDB);
	CLeadDB(const CLeadDB& aCLeadDB);
};
