#include "StdAfx.h"
#include ".\leadidset.h"

CLeadIDSet::CLeadIDSet(CLeadDB* pt_LeadDB)
: pt_LeadDBMap(NULL)
{
	pt_LeadDBMap = pt_LeadDB;
}

CLeadIDSet::~CLeadIDSet(void)
{
}
