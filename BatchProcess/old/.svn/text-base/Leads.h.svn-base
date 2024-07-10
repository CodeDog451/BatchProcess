#pragma once
#include "lead.h"
#include "afxwin.h"
#include "leadlist.h"
#include ".\statelist.h"
#import "C:\Program Files\Common Files\System\ADO\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")
class Leads
{
public:
	Leads(void);
	~Leads(void);
private:
	StateList verifiedLeads;
	StateList nonVerifiedLeads;
	//CArray<Lead,Lead&> allLeads;
	//CArray<long,long&> verifiedIndexes;
	//CArray<long,long&> nonVerifiedIndexes;
public:
	void AddLead(CLead aLead);
	//CArray GetLeads(bool verifed);
private:
	//LeadList theLeads;
	//CHListBox* m_output
public:
	LeadList GetLeads(bool verified);
	void PrintOut(CHListBox* m_output);
	void SetOutput(CHListBox* pointer);
	void Output(CString str);
private:
	CHListBox* m_output;
public:
	LeadList GetLeads(bool verified, CString state);
};
