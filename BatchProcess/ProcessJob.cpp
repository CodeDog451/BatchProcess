#include "StdAfx.h"
#include ".\processjob.h"

CProcessJob::CProcessJob(void)
: iCategory(0)
, iGroup(0)
, iState(0)
{
}

CProcessJob::~CProcessJob(void)
{
}

void CProcessJob::SetCategory(int iNum)
{
	iCategory = iNum;
}

void CProcessJob::SetGroup(int iNum)
{
	iGroup = iNum;
}

int CProcessJob::GetiCategory(void)
{
	return iCategory;
}

int CProcessJob::GetGroup(void)
{
	return iGroup;
}

void CProcessJob::SetState(long iResID)
{
	iState = iResID;
}

long CProcessJob::GetState(void)
{
	return iState;
}
