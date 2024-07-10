#pragma once

class CLogEntry
{
public:
	CLogEntry(void);
	~CLogEntry(void);
protected:
	CString sEntry;
	int iDiv;
	bool bFail;
public:
	void SetEntry(CString sMsg);
	CString GetEntry(void);
	void SetDiv(int iLoc);
	int GetDiv(void);
	void SetFail(bool bState);
	bool GetFail(void);
	void SetEntry(CString sMsg, int iLoc, bool bState);
	void Append(CString sLetter, int iLoc, bool bState);
	void Clear(void);
};
