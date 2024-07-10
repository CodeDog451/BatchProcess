#pragma once
#include "dbitem.h"
#include "atltime.h"
#include "resource.h"
#include "logdb.h"

class CLead : DBItem
{
public:
	CLead(void);
	~CLead(void);
private:
	//CLogDB myLog;
	long iLead_Id;
	//bool lead_DoNotUse;
	CString sLead_FirstName;
	CString sLead_LastName;
	CString sLead_CoApp;
	CString sLead_City;
	CString sLead_State;
	CString sLead_Address;
	CString sLead_Zip;
	CString sLead_WorkPhone;
	CString sLead_HomePhone;
	CString sLead_HouseType;
	long iLead_CurrentValue;
	long iLead_DesiredLoan;
	long iLead_1stBalance;
	double fLead_Rate;
	CString sLead_FixedAdjust;
	CString sLead_Payment;
	CString sLead_Late;
	CString sLead_Credit;
	CString sLead_WorkInfo;
	CString sLead_TimePeriod;
	long sLead_Salary;
	CString sLead_YouShouldCall;
	CString sLead_LookingTo;
	CString sLead_Email;
	//COleDateTime sLead_CreateDate;
	COleDateTime Sys_Cre_Date;
	bool bLead_Verified;
	long iLangInt;
	CString sVerifiedBy;
	//COleDateTime ImportDate;
	CString sPersonalTouch;
	bool dupe;
	CArray<long,long&> companys;
	
public:
	void SetLead_Id(long num);
	long GetLead_Id(void);
	//void SetLead_DoNotUse(bool state);
	//bool GetLead_DoNotUse(void);
	void SetLead_FirstName(CString str);
	CString GetLead_FirstName(void);
	void SetLead_LastName(CString str);
	CString GetLead_LastName(void);
	void SetLead_CoApp(CString str);
	CString GetLead_CoApp(void);
	void SetLead_City(CString str);
	CString GetLead_City(void);
	void SetLead_State(CString str);
	CString GetLead_State(void);
	void SetLead_Address(CString str);
	CString GetLead_Address(void);
	void SetLead_Zip(CString str);
	CString GetLead_Zip(void);
	void SetLead_WorkPhone(CString str);
	CString GetLead_WorkPhone(void);
	void SetLead_HomePhone(CString str);
	CString GetLead_HomePhone(void);
	void SetLead_HouseType(CString str);
	CString GetLead_HouseType(void);
	void SetLead_CurrentValue(long num);
	long GetLead_CurrentValue(void);
	void SetLead_DesiredLoan(long num);
	long GetLead_DesiredLoan(void);
	void SetLead_1stBalance(long num);
	long GetLead_1stBalance(void);
	void SetLead_Rate(double num);
	double GetLead_Rate(void);
	void SetLead_FixedAdjust(CString str);
	CString GetLead_FixedAdjust(void);
	void SetLead_Payment(CString str);
	CString GetLead_Payment(void);
	void SetLead_Late(CString str);
	CString GetLead_Late(void);
	void SetLead_Credit(CString str);
	CString GetLead_Credit(void);
	void SetLead_WorkInfo(CString str);
	CString GetLead_WorkInfo(void);
	void SetLead_TimePeriod(CString str);
	CString GetLead_TimePeriod(void);
	void SetLead_Salary(long num);
	long GetLead_Salary(void);
	void SetLead_YouShouldCall(CString str);
	CString GetLead_YouShouldCall(void);
	void SetLead_LookingTo(CString str);
	CString GetLead_LookingTo(void);
	void SetLead_Email(CString str);
	CString GetLead_Email(void);
	//void SetLead_CreateDate(COleDateTime aDate);
	//COleDateTime GetLead_CreateDate(void);
	void SetSys_Cre_Date(COleDateTime aDate);
	COleDateTime GetSys_Cre_Date(void);
	void SetLead_Verified(bool state);
	bool GetLead_Verified(void);
	void SetLangInt(long num);
	long GetLangInt(void);
	void SetVerifiedBy(CString str);
	CString GetVerifiedBy(void);
	//void SetImportDate(COleDateTime aDate);
	//COleDateTime GetImportDate(void);
	void SetPersonalTouch(CString str);
	CString GetPersonalTouch(void);
	CLead& operator=(const CLead& aLead);
	CLead(const CLead& aLead);
	CString GetLead_Id_String(void);
	//CString GetLead_DoNotUse_String(void);
	void SetDupe(bool state);
	bool GetDupe(void);
	CString GetDupe_String(void);
	CString GetLangInt_String(void);
	CString GetLead_CurrentValue_String(void);
	CString GetLead_DesiredLoan_String(void);
	CString GetLead_1stBalance_String(void);
	CString GetLead_Rate_String(void);
	CString GetLead_Salary_String(void);
	bool HasCompany(long company_id);
	void AddCompany(long company_id);
	long CountCompanys(void);
	void PrintOut(CHListBox* m_output);
	//CString GetLead_CreateDate_String(void);
private:
	CString GetDateString(COleDateTime dt);
public:
	CString GetInt_String(int num);
	CString GetSys_Cre_Date_String(void);
	CString GetBool_String(bool state);
	CString GetLead_Verified_String(void);
	//CString GetImportDate_String(void);
	bool ReadLead(_RecordsetPtr rs);
private:
	//long GetFieldLong(_RecordsetPtr rs, CString field);
public:
	//CString GetFieldString(_RecordsetPtr rs, CString field);
private:
	//CString GetCString(_bstr_t bstrStart);
public:
	//bool GetFieldBool(_RecordsetPtr rs, CString field);
	//double GetFieldDouble(_RecordsetPtr rs, CString field);
	//COleDateTime GetFieldDate(_RecordsetPtr rs, CString field);
	bool ReadLeadFilters(_RecordsetPtr rs);
	CString ToString(void);
	void SetLeadID(long iNum);
protected:
	long iCountOfPO_Company_ID;
public:
	void SetCountOfPO_Company_ID(long iNum);
	CString GetCountOfPO_Company_ID(void);
	CString GetLTV(void);
	long GetUses(void);
protected:
	HWND mainWin;
public:
	void SetParentWindow(HWND m_hwnd);
	void PrintOut(CString str);
	bool IsPurchase(void);
};
