#pragma once
#include "resource.h"
#include ".\writelog.h"
//a class to contain a fax job
class CFaxJob
{
public:
	CFaxJob(void);
	~CFaxJob(void);
protected:
	CString sFaxnum;
	CString sAgentname;
	CString sAgentEmail;
	CString sCompany;
	int iDiv;
	CString sFilename;	
public:
	void SetFaxnum(CString sNum);
	CString GetFaxnum(void);
	void SetAgentname(CString sName);
	CString GetAgentname(void);
	void SetAgentEmail(CString sEmail);
	CString GetAgentEmail(void);
	void SetCompany(CString sName);
	CString GetCompany(void);
	int GetDiv(void);
	void SetDiv(int iNum);
	void SetFilename(CString sPath);
	CString GetFilename(void);
protected:
	CString sContact;
public:
	void SetContact(CString sName);
	CString GetContact(void);
protected:
	long iOrderID;
	long iCampaignID;
public:
	void SetOrderID(long iID);
	CString GetCampaignID(void);
	void SetCampaignID(long iID);
	CString GetOrderID(void);
	CString GetIDMessage(void);
	void ClearAll(void);
	void SendTo(HWND mainWin);
protected:
	void SendString(HWND mainWin, int resid, CString str);
	CString sStates;
public:
	void SetStates(CString str);
	CString GetStates(void);
protected:
	int iSendCoverpage;
public:
	void SetSendCoverpage(int iNum);
	int GetSendCoverPage(void);
	void PrintOut(CString str);
};
