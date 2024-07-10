#pragma once
#include "leadlist.h"
#include ".\dbitem.h"
class State : DBItem
{
public:
	State(void);
	~State(void);
	LeadList leads;
private:
	CString st_Name;
	CString st_Abr;
public:
	void SetSt_Name(CString name);
	void SetSt_Abr(CString abr);
	CString GetSt_Name(void);
	CString GetSt_Abr(void);
	void AddLead(CLead aLead);
	LeadList GetLeads(void);
	void ReadState(_RecordsetPtr rs);
	void PrintOut(CHListBox* m_output);
private:
	CHListBox* m_output;
public:
	void SetOutput(CHListBox* pointer);
private:
	void Output(CString str);
};
