#pragma once
#include "state.h"
#include "afxwin.h"
#import "C:\Program Files\Common Files\System\ADO\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")
class StateList
{
public:
	StateList(void);
	~StateList(void);
	StateList& operator=(const StateList& aStateList);
	StateList(const StateList& aStateList);
private:
	CArray<State,State&> theStates;
public:
	void AddState(State aState);
	int Contains(CString state);
	void PrintOut(CHListBox* m_output);
	void ReadStates(void);
	void AddLead(CLead aLead);
	LeadList GetLeads(void);
private:
	CHListBox* m_output;
public:
	void SetOutput(CHListBox* pointer);
private:
	void Output(CString str);
public:
	LeadList GetLeads(CString st_Abr);
};
