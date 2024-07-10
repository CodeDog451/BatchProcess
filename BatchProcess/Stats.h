#pragma once
#include "dbitem.h"
#include "printout.h"

class CStats :
	public DBItem
{
public:
	CStats(void);
	~CStats(void);
	CString GetSQLLeadCounts(void);
protected:
	CPrintOut output;
public:
	void SetParentWindow(HWND m_wnd);
protected:
	void PrintOut(CString str);
	long iVerified;
	long iOnePass;
	long iInternet;
public:
	CString GetCountVerified(void);
	CString GetCountOnePass(void);
	CString GetCountInternet(void);
	void ReadTodaysLeadsSent(void);
protected:
	void SetCountInternet(long iNum);
public:
	void SetCountOnePass(long iNum);
protected:
	void SetCountVerified(long iNum);
public:
	CString GetCountTotal(void);
protected:
	long iInternetOrdered;
	long iOnePassOrdered;
	long iVerifiedOrdered;
public:
	CString GetSQLLeadOrderedCounts(void);
	void ReadTodaysLeadsOrdered(void);
protected:
	void SetCountVerifiedOrdered(long iNum);
	void SetCountInternetOrdered(long iNum);
	void SetCountOnePassOrdered(long iNum);
public:
	CString GetCountTotalOrdered(void);
	CString GetCountVerifiedOrdered(void);
	CString GetCountInternetOrdered(void);
	CString GetCountOnePassOrdered(void);
	CString GetYieldVerified(void);
	CString GetYieldInternet(void);
	CString GetYieldOnePass(void);
	
	CString GetYieldTotal(void);
	CString GetSQLCampaigns(void);
	CString GetSQLZerosSent(void);
	void ReadTodaysZerosSent(void);
protected:
	int iSentZero;
	int iSentAll;
	int iSentPartial;
	int iTotalCampaigns;
public:
	CString GetTotalCampaigns(void);
	CString GetSentAll(void);
	CString GetSentZero(void);
	CString GetSentPartials(void);
	CString GetPercentZero(void);
	CString GetPercentPartials(void);
	CString GetPercentAll(void);
	double fMean;
	CString GetMean(void);
};
