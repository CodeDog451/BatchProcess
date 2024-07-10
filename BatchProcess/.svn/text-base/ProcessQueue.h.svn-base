#pragma once
#include "dbitem.h"
#include ".\processjob.h"
#include "resource.h"

class CProcessQueue :
	public DBItem
{
public:
	CProcessQueue(void);
	~CProcessQueue(void);
	CProcessQueue& operator=(const CProcessQueue& aCProcessQueue);
	CProcessQueue(const CProcessQueue& aCProcessQueue);
protected:
	CArray<CProcessJob,CProcessJob> myJobs;
public:
	bool ReadJobs(void);
protected:
	HWND mainWin;
public:
	void SetParentWindow(HWND m_hwnd);
	void PrintOut(CString str);
	CProcessJob GetJobAt(int index);
	int GetCount(void);
	void DeleteAt(int index);
	CProcessJob GetNextJob(void);
};
