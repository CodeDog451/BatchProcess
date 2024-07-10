#include "StdAfx.h"
#include ".\logentry.h"

CLogEntry::CLogEntry(void)
: sEntry(_T(""))
, iDiv(0)
, bFail(false)
{
}

CLogEntry::~CLogEntry(void)
{
}

void CLogEntry::SetEntry(CString sMsg)
{
	sEntry = sMsg;
}

CString CLogEntry::GetEntry(void)
{
	return sEntry;
}

void CLogEntry::SetDiv(int iLoc)
{
	iDiv = iLoc;
}

int CLogEntry::GetDiv(void)
{
	return iDiv;
}

void CLogEntry::SetFail(bool bState)
{
	bFail = bState;
}

bool CLogEntry::GetFail(void)
{
	return bFail;
}

void CLogEntry::SetEntry(CString sMsg, int iLoc, bool bState)
{
	sEntry = sMsg;
	iDiv = iLoc;
	bFail = bState;
}

void CLogEntry::Append(CString sLetter, int iLoc, bool bState)
{
	sEntry = sEntry + sLetter;
	iDiv = iLoc;
	bFail = bState;
}

void CLogEntry::Clear(void)
{
	bFail = false;
	iDiv = 0;
	sEntry = "";
}
