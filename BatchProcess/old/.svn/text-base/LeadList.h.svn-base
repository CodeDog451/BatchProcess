#pragma once
#include "lead.h"
#include "afxwin.h"
class LeadList
{
public:
	LeadList(void);
	~LeadList(void);
	LeadList& operator=(const LeadList& aLeadList);
	LeadList(const LeadList& aLeadList);
private:
	CArray<CLead*,CLead*&> theLeads;
public:
	CLead* add(CLead* aLead);
	CLead* getAt(long index);
	long getCount(void);
	long getUpperBound(void);
	void Append(LeadList aLeadList);
	long Contains(CLead* aLead);
	
};
