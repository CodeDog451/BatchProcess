#pragma once
#include ".\writelog.h"
class CIDArray
{
public:
	//CIDArray(CLeadDB* pt_LeadDB);
	CIDArray(void);
	~CIDArray(void);
	CIDArray& operator=(const CIDArray& aCIDArray);
	CIDArray(const CIDArray& aCIDArray);
	long& operator[] (unsigned long i); 
protected:
	CArray<long,long> myIDs;
	
public:
	long& GetAt(long index);
	INT_PTR Add(long id);
	INT_PTR Append(CIDArray ids);
	INT_PTR GetCount(void);
	void RemoveAll(void);
	void RemoveAt(INT_PTR index);
	void Copy(CIDArray ids);
};
