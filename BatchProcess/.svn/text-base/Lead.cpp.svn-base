#include "StdAfx.h"
#include ".\lead.h"

CLead::CLead(void)
: iLead_Id(0)
//, lead_DoNotUse(false)
, sLead_FirstName(_T(""))
, sLead_LastName(_T(""))
, sLead_CoApp(_T(""))
, sLead_City(_T(""))
, sLead_State(_T(""))
, sLead_Address(_T(""))
, sLead_Zip(_T(""))
, sLead_WorkPhone(_T(""))
, sLead_HomePhone(_T(""))
, sLead_HouseType(_T(""))
, iLead_CurrentValue(0)
, iLead_DesiredLoan(0)
, iLead_1stBalance(0)
, fLead_Rate(0)
, sLead_FixedAdjust(_T(""))
, sLead_Payment(_T(""))
, sLead_Late(_T(""))
, sLead_Credit(_T(""))
, sLead_WorkInfo(_T(""))
, sLead_TimePeriod(_T(""))
, sLead_Salary(0)
, sLead_YouShouldCall(_T(""))
, sLead_LookingTo(_T(""))
, sLead_Email(_T(""))
//, sLead_CreateDate()
, Sys_Cre_Date()
, bLead_Verified(false)
, iLangInt(0)
, sVerifiedBy(_T(""))
//, ImportDate()
, sPersonalTouch(_T(""))
, dupe(false)
, iCountOfPO_Company_ID(0)
{
}

CLead::~CLead(void)
{
}

void CLead::SetLead_Id(long num)
{
	
	try
	{
		iLead_Id = num;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::SetLead_Id(long num)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

long CLead::GetLead_Id(void)
{	
	try
	{
		return iLead_Id;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::GetLead_Id";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return -1;
	}
}

void CLead::SetLead_FirstName(CString str)
{
	
	try
	{
		sLead_FirstName = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::SetLead_FirstName(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CLead::GetLead_FirstName(void)
{
	
	try
	{
		return sLead_FirstName;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::GetLead_FirstName";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

void CLead::SetLead_LastName(CString str)
{
	
	try
	{
		sLead_LastName = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::SetLead_LastName(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CLead::GetLead_LastName(void)
{
	
	try
	{
		return sLead_LastName;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::GetLead_LastName";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CLead::SetLead_CoApp(CString str)
{
	
	try
	{
		sLead_CoApp = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::SetLead_CoApp(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CLead::GetLead_CoApp(void)
{
	
	try
	{
		return sLead_CoApp;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::GetLead_CoApp";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CLead::SetLead_City(CString str)
{
	
	try
	{
		sLead_City = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::SetLead_City(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CLead::GetLead_City(void)
{
	
	try
	{
		return sLead_City;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::GetLead_City";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CLead::SetLead_State(CString str)
{
	
	try
	{
		sLead_State = str.MakeUpper();
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::SetLead_State(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CLead::GetLead_State(void)
{
	
	try
	{
		return sLead_State;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::GetLead_State";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

void CLead::SetLead_Address(CString str)
{
	
	try
	{
		sLead_Address = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::SetLead_Address(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CLead::GetLead_Address(void)
{
	
	try
	{
		return sLead_Address;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::GetLead_Address";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CLead::SetLead_Zip(CString str)
{
	
	try
	{
		sLead_Zip = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::SetLead_Zip(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CLead::GetLead_Zip(void)
{
	
	try
	{
		return sLead_Zip;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::GetLead_Zip";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CLead::SetLead_WorkPhone(CString str)
{
	
	try
	{
		sLead_WorkPhone = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::SetLead_WorkPhone(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CLead::GetLead_WorkPhone(void)
{
	
	try
	{
		return sLead_WorkPhone;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::GetLead_WorkPhone";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CLead::SetLead_HomePhone(CString str)
{
	
	try
	{
		sLead_HomePhone = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CLead::SetLead_HomePhone(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CLead::GetLead_HomePhone(void)
{
	
	try
	{
		return sLead_HomePhone;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: functionname";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CLead::SetLead_HouseType(CString str)
{
	
	try
	{
		sLead_HouseType = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: functionname";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CLead::GetLead_HouseType(void)
{
	return sLead_HouseType;
}

void CLead::SetLead_CurrentValue(long num)
{
	iLead_CurrentValue = num;
}

long CLead::GetLead_CurrentValue(void)
{
	return iLead_CurrentValue;
}

void CLead::SetLead_DesiredLoan(long num)
{
	iLead_DesiredLoan = num;
}

long CLead::GetLead_DesiredLoan(void)
{
	return iLead_DesiredLoan;
}

void CLead::SetLead_1stBalance(long num)
{
	iLead_1stBalance = num;
}

long CLead::GetLead_1stBalance(void)
{
	return iLead_1stBalance;
}

void CLead::SetLead_Rate(double num)
{
	fLead_Rate = num;
}

double CLead::GetLead_Rate(void)
{
	return fLead_Rate;
}

void CLead::SetLead_FixedAdjust(CString str)
{
	sLead_FixedAdjust = str;
}

CString CLead::GetLead_FixedAdjust(void)
{
	return sLead_FixedAdjust;
}

void CLead::SetLead_Payment(CString str)
{

	CString amount = str;
	amount.Replace("$", "");
	amount.Replace(",", "");
	double fNum = atof(amount);
	long iNum = fNum;
	char buffer[200];
	itoa(iNum, buffer, 10);

	sLead_Payment = buffer;
}

CString CLead::GetLead_Payment(void)
{
	return sLead_Payment;
}

void CLead::SetLead_Late(CString str)
{
	sLead_Late = str;
}

CString CLead::GetLead_Late(void)
{
	return sLead_Late;
}

void CLead::SetLead_Credit(CString str)
{
	
	sLead_Credit = str;
}

CString CLead::GetLead_Credit(void)
{
	//AfxMessageBox(sLead_Credit);
	return sLead_Credit;
}

void CLead::SetLead_WorkInfo(CString str)
{
	sLead_WorkInfo = str;
}

CString CLead::GetLead_WorkInfo(void)
{
	return sLead_WorkInfo;
}

void CLead::SetLead_TimePeriod(CString str)
{
	sLead_TimePeriod = str;
}

CString CLead::GetLead_TimePeriod(void)
{
	return sLead_TimePeriod;
}

void CLead::SetLead_Salary(long num)
{
	sLead_Salary = num;
}

long CLead::GetLead_Salary(void)
{
	return sLead_Salary;
}

void CLead::SetLead_YouShouldCall(CString str)
{
	sLead_YouShouldCall = str;
}

CString CLead::GetLead_YouShouldCall(void)
{
	return sLead_YouShouldCall;
}

void CLead::SetLead_LookingTo(CString str)
{
	sLead_LookingTo = str;
}

CString CLead::GetLead_LookingTo(void)
{
	return sLead_LookingTo;
}

void CLead::SetLead_Email(CString str)
{
	sLead_Email = str;
}

CString CLead::GetLead_Email(void)
{
	return sLead_Email;
}

//void CLead::SetLead_CreateDate(COleDateTime aDate)
//{
	//sLead_CreateDate = aDate;
//}

//COleDateTime CLead::GetLead_CreateDate(void)
//{
	//return sLead_CreateDate;
//}

void CLead::SetSys_Cre_Date(COleDateTime aDate)
{
	Sys_Cre_Date = aDate;
}

COleDateTime CLead::GetSys_Cre_Date(void)
{
	return Sys_Cre_Date;
}

void CLead::SetLead_Verified(bool state)
{
	bLead_Verified = state;
}

bool CLead::GetLead_Verified(void)
{
	return bLead_Verified;
}

void CLead::SetLangInt(long num)
{
	iLangInt = num;
}

long CLead::GetLangInt(void)
{
	return iLangInt;
}

void CLead::SetVerifiedBy(CString str)
{
	sVerifiedBy = str;
}

CString CLead::GetVerifiedBy(void)
{
	return sVerifiedBy;
}

//void CLead::SetImportDate(COleDateTime aDate)
//{
	//ImportDate = aDate;
//}

//COleDateTime CLead::GetImportDate(void)
//{
	//return ImportDate;
//}

void CLead::SetPersonalTouch(CString str)
{
	sPersonalTouch = str;
}

CString CLead::GetPersonalTouch(void)
{
	return sPersonalTouch;
}

CLead& CLead::operator=(const CLead& aLead)
{
	//asignment
	iLead_Id = aLead.iLead_Id;
	iCountOfPO_Company_ID = aLead.iCountOfPO_Company_ID;
	//lead_DoNotUse = aLead.lead_DoNotUse;
	sLead_FirstName = aLead.sLead_FirstName;
	sLead_LastName = aLead.sLead_LastName;
	sLead_CoApp = aLead.sLead_CoApp;
	sLead_City = aLead.sLead_City;
	sLead_State = aLead.sLead_State;
	sLead_Address = aLead.sLead_Address;
	sLead_Zip = aLead.sLead_Zip;
	sLead_WorkPhone = aLead.sLead_WorkPhone;
	sLead_HomePhone = aLead.sLead_HomePhone;
	sLead_HouseType = aLead.sLead_HouseType;
	iLead_CurrentValue = aLead.iLead_CurrentValue;
	iLead_DesiredLoan = aLead.iLead_DesiredLoan;
	iLead_1stBalance = aLead.iLead_1stBalance;
	fLead_Rate = aLead.fLead_Rate;
	sLead_FixedAdjust = aLead.sLead_FixedAdjust;
	sLead_Payment = aLead.sLead_Payment;
	sLead_Late = aLead.sLead_Late;
	sLead_Credit = aLead.sLead_Credit;
	sLead_WorkInfo = aLead.sLead_WorkInfo;
	sLead_TimePeriod = aLead.sLead_TimePeriod;
	sLead_Salary = aLead.sLead_Salary;
	sLead_YouShouldCall = aLead.sLead_YouShouldCall;
	sLead_LookingTo = aLead.sLead_LookingTo;
	sLead_Email = aLead.sLead_Email;
	//sLead_CreateDate = aLead.sLead_CreateDate;
	Sys_Cre_Date = aLead.Sys_Cre_Date;
	bLead_Verified = aLead.bLead_Verified;
	dupe = aLead.dupe;
	iLangInt = aLead.iLangInt;
	sVerifiedBy = aLead.sVerifiedBy;
	//ImportDate = aLead.ImportDate;
	sPersonalTouch = aLead.sPersonalTouch;
	mainWin = aLead.mainWin;
	companys.Copy(aLead.companys);

	return *this;
}

CLead::CLead(const CLead& aLead)
{
	//copy constructor
	iLead_Id = aLead.iLead_Id;
	//lead_DoNotUse = aLead.lead_DoNotUse;
	iCountOfPO_Company_ID = aLead.iCountOfPO_Company_ID;
	sLead_FirstName = aLead.sLead_FirstName;
	sLead_LastName = aLead.sLead_LastName;
	sLead_CoApp = aLead.sLead_CoApp;
	sLead_City = aLead.sLead_City;
	sLead_State = aLead.sLead_State;
	sLead_Address = aLead.sLead_Address;
	sLead_Zip = aLead.sLead_Zip;
	sLead_WorkPhone = aLead.sLead_WorkPhone;
	sLead_HomePhone = aLead.sLead_HomePhone;
	sLead_HouseType = aLead.sLead_HouseType;
	iLead_CurrentValue = aLead.iLead_CurrentValue;
	iLead_DesiredLoan = aLead.iLead_DesiredLoan;
	iLead_1stBalance = aLead.iLead_1stBalance;
	fLead_Rate = aLead.fLead_Rate;
	sLead_FixedAdjust = aLead.sLead_FixedAdjust;
	sLead_Payment = aLead.sLead_Payment;
	sLead_Late = aLead.sLead_Late;
	sLead_Credit = aLead.sLead_Credit;
	sLead_WorkInfo = aLead.sLead_WorkInfo;
	sLead_TimePeriod = aLead.sLead_TimePeriod;
	sLead_Salary = aLead.sLead_Salary;
	sLead_YouShouldCall = aLead.sLead_YouShouldCall;
	sLead_LookingTo = aLead.sLead_LookingTo;
	sLead_Email = aLead.sLead_Email;
	//sLead_CreateDate = aLead.sLead_CreateDate;
	Sys_Cre_Date = aLead.Sys_Cre_Date;
	bLead_Verified = aLead.bLead_Verified;
	dupe = aLead.dupe;
	iLangInt = aLead.iLangInt;
	sVerifiedBy = aLead.sVerifiedBy;
	//ImportDate = aLead.ImportDate;
	sPersonalTouch = aLead.sPersonalTouch;
	mainWin = aLead.mainWin;
	companys.Copy(aLead.companys);

}

CString CLead::GetLead_Id_String(void)
{
	char buffer[200];
	_ltoa( iLead_Id, buffer, 10 );
	CString str;
	str = buffer;
	return str;
}

//CString Lead::GetLead_DoNotUse_String(void)
//{	
	//return GetBool_String(lead_DoNotUse);
//}

void CLead::SetDupe(bool state)
{
	dupe = state;
}

bool CLead::GetDupe(void)
{
	return dupe;
}

CString CLead::GetDupe_String(void)
{
	CString str;
	if(dupe)
	{
		str = "true";
	}
	else
	{
		str = "false";
	}
	return str;
}

CString CLead::GetLangInt_String(void)
{
	//0;"Don't Care";1;"English";2;"Spanish"
	CString str;
	if(iLangInt == 2)
	{
		str = "Spanish";
		
		
	}
	else 
	{
		str = "English";
	}	
	return str;
}

CString CLead::GetLead_CurrentValue_String(void)
{
	char buffer[200];
	_itoa(iLead_CurrentValue, buffer, 10 );
	CString str;
	str = buffer;
	return str;
	
}

CString CLead::GetLead_DesiredLoan_String(void)
{
	char buffer[200];
	_itoa(iLead_DesiredLoan, buffer, 10 );
	CString str;
	str = buffer;
	return str;
}

CString CLead::GetLead_1stBalance_String(void)
{
	char buffer[200];
	_itoa(iLead_1stBalance, buffer, 10 );
	CString str;
	str = buffer;
	return str;
	
}

CString CLead::GetLead_Rate_String(void)
{
	CString str;
	if(fLead_Rate>0)
	{
		char buffer[200];
		_gcvt(fLead_Rate, 7, buffer );
		
		str = buffer;
		if(str.GetAt(str.GetLength()-1)=='.')
		{
			str = str + "0";
		}
		str = str + "%";
	}
	return str;
}

CString CLead::GetLead_Salary_String(void)
{
	char buffer[200];
	_itoa(sLead_Salary, buffer, 10 );
	CString str;
	str = buffer;
	return str;
}

bool CLead::HasCompany(long company_id)
{
	bool contains = false;
	for(int x = 0; x <= companys.GetUpperBound(); x++)
	{
		if(companys[x] == company_id) contains = true;
	}
	return contains;
}

void CLead::AddCompany(long company_id)
{
	if(!HasCompany(company_id)) companys.Add(company_id);
}

long CLead::CountCompanys(void)
{
	return companys.GetCount();
}

void CLead::PrintOut(CHListBox* m_output)
{
	CString str;
	str = ToString();

	m_output->AddString(str);	
	m_output->SetCurSel(m_output->GetCount()-1);		
	m_output->RedrawWindow();
}

//CString CLead::GetLead_CreateDate_String(void)
//{	
	//return GetDateString(GetLead_CreateDate());
//}

CString CLead::GetDateString(COleDateTime dt)
{
	CString aDateStr;
	aDateStr = GetInt_String(dt.GetMonth()) + "/";
	aDateStr = aDateStr + GetInt_String(dt.GetDay()) + "/";
	aDateStr = aDateStr + GetInt_String(dt.GetYear());
	return aDateStr;
}

CString CLead::GetInt_String(int num)
{
	char buffer[200];
	_itoa( num, buffer, 10 );
	CString str;			
	str = buffer;
	return str;
}

CString CLead::GetSys_Cre_Date_String(void)
{
	return GetDateString(GetSys_Cre_Date());
}

CString CLead::GetBool_String(bool state)
{
	CString str;
	if(state)
	{
		str = "true";
	}
	else
	{
		str = "false";
	}
	return str;
}

CString CLead::GetLead_Verified_String(void)
{
	return GetBool_String(bLead_Verified);
}

//CString CLead::GetImportDate_String(void)
//{
	//return GetDateString(ImportDate);
//}
bool CLead::ReadLeadFilters(_RecordsetPtr rs)
{
	SetLead_Id(GetFieldLong(rs,"lead_id"));	
	SetLead_State(GetFieldString(rs, "lead_State"));	
	SetLead_Zip(GetFieldString(rs, "lead_Zip"));
	SetLead_WorkPhone(GetFieldString(rs, "lead_WorkPhone"));
	SetLead_HomePhone(GetFieldString(rs, "lead_HomePhone"));	
	SetLead_DesiredLoan(GetFieldDouble(rs, "lead_DesiredLoan"));	
	SetLead_Rate(GetFieldDouble(rs, "lead_Rate"));	
	SetLead_Credit(GetFieldString(rs, "lead_Credit"));	
	SetLead_LookingTo(GetFieldString(rs, "lead_LookingTo"));	
	SetSys_Cre_Date(GetFieldDate(rs, "Sys_Cre_Date"));
	SetLead_Verified(GetFieldBool(rs, "lead_Verified"));	
	SetLangInt(GetFieldLong(rs, "langInt"));

	return true;
}
bool CLead::ReadLead(_RecordsetPtr rs)
{
	//PrintOut("Reading Lead");
	ReadLeadFilters(rs);
	SetLeadID(GetFieldLong(rs, "lead_Id"));
	//SetCountOfPO_Company_ID(GetFieldLong(rs, "CountOfPO_Company_ID"));//old GetSql()
	SetCountOfPO_Company_ID(GetFieldLong(rs, "Used"));//New GetNewSQL()
	//SetLead_DoNotUse(GetFieldBool(rs, "lead_DoNotUse"));
	SetLead_FirstName(GetFieldString(rs, "lead_FirstName"));
	SetLead_LastName(GetFieldString(rs, "lead_LastName"));	

	SetLead_CoApp(GetFieldString(rs, "lead_CoApp"));
	SetLead_City(GetFieldString(rs, "lead_City"));
	SetLead_State(GetFieldString(rs, "lead_State"));
	
	SetLead_Address(GetFieldString(rs, "lead_Address"));
	
	SetLead_HouseType(GetFieldString(rs, "lead_HouseType"));
	SetLead_CurrentValue(GetFieldDouble(rs, "lead_CurrentValue"));
	
	SetLead_1stBalance(GetFieldDouble(rs, "lead_1stBalance"));
	
	SetLead_FixedAdjust(GetFieldString(rs, "lead_FixedAdjust"));
	SetLead_Payment(GetFieldString(rs, "lead_Payment"));
	SetLead_Late(GetFieldString(rs, "lead_Late"));
	SetLead_Credit(GetFieldString(rs, "lead_Credit"));
	
	SetLead_WorkInfo(GetFieldString(rs, "lead_WorkInfo"));
	SetLead_TimePeriod(GetFieldString(rs, "lead_TimePeriod"));
	SetLead_Salary(GetFieldDouble(rs, "lead_Salary"));
	SetLead_YouShouldCall(GetFieldString(rs, "lead_YouShouldCall"));
	
	SetLead_Email(GetFieldString(rs, "lead_Email"));
	//SetLead_CreateDate(GetFieldDate(rs, "lead_CreateDate"));
	
	//SetDupe(GetFieldBool(rs, "dupe"));
	SetLead_Verified(GetFieldBool(rs, "lead_Verified"));
	SetLangInt(GetFieldLong(rs, "langInt"));

	SetVerifiedBy(GetFieldString(rs, "VerifiedBy"));
	//SetImportDate(GetFieldDate(rs, "ImportDate"));
	SetPersonalTouch(GetFieldString(rs, "PersonalTouch"));
	
	return true;
}


CString CLead::ToString(void)
{
	CString str;
	str = "(LeadID:";
	str = str +GetLead_Id_String();
	str = str + ")(Used Original:";
	str = str + GetCountOfPO_Company_ID();
	str = str + ")(Used after assigns:";
	char buffer[200];
	_itoa(GetUses(), buffer, 10);
	str = str + buffer;
	str = str + ")(Verified:";
	str = str + GetLead_Verified_String();
	str = str + ")(First:";	
	str = str + GetLead_FirstName();
	str = str + ")(Last:";
	str = str + GetLead_LastName();
	str = str + ")(CoApp:";
	str = str + GetLead_CoApp();
	str = str + ")(Address:";
	str = str + GetLead_Address();
	str = str + ")(City:";
	str = str + GetLead_City();
	str = str + ")(State:";
	str = str + GetLead_State();
	str = str + ")(Zip:";
	str = str + GetLead_Zip();
	str = str + ")(Home:";
	str = str + GetLead_HomePhone();
	str = str + ")(Work:";
	str = str + GetLead_WorkPhone();	
	str = str + ")(HomeType:";
	str = str + GetLead_HouseType();
	str = str + ")(CurrentValue:";
	str = str + GetLead_CurrentValue_String();
	str = str + ")(DesiredLoan:";
	str = str + GetLead_DesiredLoan_String();
	str = str + ")(1stBal:";
	str = str + GetLead_1stBalance_String();
	str = str + ")(Rate:";
	str = str + GetLead_Rate_String();
	str = str + ")(FixedAdj:";
	str = str + GetLead_FixedAdjust();
	str = str + ")(Payment:";
	str = str + GetLead_Payment();
	str = str + ")(Late:";
	str = str + GetLead_Late();
	str = str + ")(Credit:";
	str = str + GetLead_Credit();
	str = str + ")(WorkInfo:";
	str = str + GetLead_WorkInfo();
	str = str + ")(TimePeriod:";
	str = str + GetLead_TimePeriod();	
	str = str + ")(Salary:";
	str = str + GetLead_Salary_String();
	str = str + ")(YouShouldCall:";
	str = str + GetLead_YouShouldCall();
	str = str + ")(LookingTo:";
	str = str + GetLead_LookingTo();
	str = str + ")(Email:";
	str = str + GetLead_Email();	
	str = str + ")(SysCreDate:";
	str = str + GetSys_Cre_Date_String();	
	str = str + ")(Lang:";
	str = str + GetLangInt_String();
	str = str + ")(VerifiedBy:";
	str = str + GetVerifiedBy();	
	str = str + ")(PersonalTouch:";
	str = str + GetPersonalTouch();
	
	str = str + ")(LTV:";
	str = str + GetLTV();
	str = str + ")";
	return str;
}

void CLead::SetLeadID(long iNum)
{
	iLead_Id = iNum;
}

void CLead::SetCountOfPO_Company_ID(long iNum)
{
	iCountOfPO_Company_ID = iNum;
}

CString CLead::GetCountOfPO_Company_ID(void)
{
	char buffer[200];
	_ltoa( iCountOfPO_Company_ID, buffer, 10 );
	CString str;
	str = buffer;
	return str;
}

CString CLead::GetLTV(void)
{
	CString str;
	if(iLead_DesiredLoan > 0)
	{
		double ltv = (float)iLead_1stBalance/(float)iLead_CurrentValue;
		char buffer[200];
		_gcvt(ltv, 7, buffer );	
		str = buffer;		
		
	}
	else
	{
		str = "NA";
	}
	return str;
}

long CLead::GetUses(void)
{	
	//if(companys.GetCount() > 0) PrintOut("Lead has companys");
	return iCountOfPO_Company_ID + companys.GetCount();
}

void CLead::SetParentWindow(HWND m_hwnd)
{
	mainWin = m_hwnd;
}

void CLead::PrintOut(CString str)
{
	//myLog.WriteLog(str, 0);
	if(mainWin)
	{
		SendMessage(mainWin, S_START_MESSAGE, 0, 0);
		int iLen = str.GetLength();
		for(int x = 0; x < iLen; x++)
		{
			int iLetter = (int) str.GetAt(x);
			SendMessage(mainWin, S_CHAR_MESSAGE, iLetter, 0);
		}

		SendMessage(mainWin, S_END_MESSAGE, 0, 0);
	}
}

bool CLead::IsPurchase(void)
{	
	
	if((sLead_LookingTo.Find("Pur") > -1) || (sLead_LookingTo.Find("pur") > -1))
	{
		return true;
	}
	else
	{
		return false;
	}

}








 
