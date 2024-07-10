#include "StdAfx.h"
#include ".\location.h"

CLocation::CLocation(void)
: returnEmail(_T(""))
, div(0)
, subject(_T(""))
, body(_T(""))
, zeroBody(_T(""))
, desDir(_T(""))
, templateDir(_T(""))
, filterStatement(_T(""))
, sEmailNoteVerifiedPage(_T(""))
, sEmailNoteNonVerifiedPage(_T(""))
, iAgeLimitVerified(0)
, iUseLimitVerified(0)
, iAgeLimitInternet(0)
, iUseLimitInternet(0)
, iUseLimitOnePass(0)
, iAgeLimitOnePass(0)
{
}

CLocation::~CLocation(void)
{
}

CString CLocation::GetReturnEmail(void)
{
	return returnEmail;
}

int CLocation::GetDiv(void)
{
	return div;
}

CString CLocation::GetSubject(void)
{
	return subject;
}

CString CLocation::GetBody(void)
{
	return body;
}



void CLocation::SetReturnMail(CString email)
{
	returnEmail = email;
}

void CLocation::SetDiv(int iNum)
{
	div = iNum;
}

void CLocation::SetSubject(CString str)
{
	subject = str;
}



void CLocation::SetBody(CString str)
{
	body = str;
}

void CLocation::SetZeroBody(CString str)
{
	zeroBody = str;
}

CString CLocation::GetZeroBody(void)
{
	return zeroBody;
}

void CLocation::SetDesDir(CString sPath)
{
	desDir = sPath;
}

void CLocation::SetTemplateDir(CString sPath)
{
	templateDir = sPath;
}

CString CLocation::GetDesDir(void)
{
	return desDir;
}

CString CLocation::GetTemplateDir(void)
{
	return templateDir;
}

void CLocation::SetFilterStatement(CString str)
{
	
	filterStatement = str;
}

CString CLocation::GetFilterStatement(void)
{
	return filterStatement;
}

CString CLocation::GetBody(CString brokerEmail, CString agentname, CString agentEmail, CString company, CString contact, CString orderid, CString states)
{
	CString body;
	body = GetBody();
	body = Replace(body, brokerEmail, agentname, agentEmail, company, contact, orderid, states);
	return body;
}

CString CLocation::Replace(CString sSource, CString brokerEmail, CString agentname, CString agentEmail, CString company, CString contact, CString orderid, CString states)
{
	CString sResult = sSource;
	sResult.Replace("<brokeremail>", brokerEmail);	
	sResult.Replace("<agentname>", agentname);	
	sResult.Replace("<agentemail>", agentEmail);
	sResult.Replace("<orderid>", orderid);
	sResult.Replace("<company>", company);
	sResult.Replace("<contact>", contact);
	sResult.Replace("<state>", states);

	return sResult;
}

CString CLocation::GetSubject(CString brokerEmail, CString agentname, CString agentEmail, CString company, CString contact, CString orderid, CString states)
{
	CString sSubject;
	sSubject = GetSubject();
	sSubject = Replace(sSubject, brokerEmail, agentname, agentEmail, company, contact, orderid, states);
	return sSubject;
}

void CLocation::SetEmailNoteVerifiedPage(CString sPage)
{
	sEmailNoteVerifiedPage = sPage;
}

void CLocation::SetEmailNoteNonVerifiedPage(CString sPage)
{
	sEmailNoteNonVerifiedPage = sPage;
}

CString CLocation::GetEmailNoteVerifiedPage(void)
{
	return sEmailNoteVerifiedPage;
}

CString CLocation::GetEmailNoteNonVerifiedPage(void)
{
	return sEmailNoteNonVerifiedPage;
}


void CLocation::SetAgeLimitVerified(long iNum)
{
	iAgeLimitVerified = iNum;
}

void CLocation::SetUseLimitVerified(long iNum)
{
	iUseLimitVerified = iNum;
}

long CLocation::GetUseLimitVerified(void)
{
	return iUseLimitVerified;
}

long CLocation::GetAgeLimitVerified(void)
{
	return iAgeLimitVerified;
}

CString CLocation::GetAgeLimitVerifiedString(void)
{
	char buffer[200];
	sprintf(buffer, "%d", iAgeLimitVerified);
	return CString(buffer);
}

void CLocation::SetAgeLimitInternet(long iNum)
{
	iAgeLimitInternet = iNum;
}

long CLocation::GetAgeLimitInternet(void)
{
	return iAgeLimitInternet;
}

CString CLocation::GetAgeLimitInternetString(void)
{
	char buffer[200];
	sprintf(buffer, "%d", iAgeLimitInternet);
	return CString(buffer);
}

void CLocation::SetUseLimitInternet(long iNum)
{
	iUseLimitInternet = iNum;
}

long CLocation::GetUseLimitInternet(void)
{
	return iUseLimitInternet;
}

void CLocation::SetUseLimitOnePass(long iNum)
{
	iUseLimitOnePass = iNum;
}

long CLocation::GetUseLimitOnePass(void)
{
	return iUseLimitOnePass;
}

void CLocation::SetAgeLimitOnePass(long iNum)
{
	iAgeLimitOnePass = iNum;
}

long CLocation::GetAgeLimitOnePass(void)
{
	return iAgeLimitOnePass;
}

CString CLocation::GetAgeLimitOnePassString(void)
{
	char buffer[200];
	sprintf(buffer, "%d", iAgeLimitOnePass);
	return CString(buffer);
}

CString CLocation::GetLimitsString(void)
{
	CString str;
	str = "[Verified Age Limit:" + this->GetAgeLimitVerifiedString();
	str = str + "][Verified Use Limit:" + GetString(this->iUseLimitVerified);
	str = str + "][Internet Age Limit:" + this->GetAgeLimitInternetString();
	str = str + "][Internet Use Limit:" + GetString(this->iUseLimitInternet);
	str = str + "][One Pass Age Limit:" + this->GetAgeLimitOnePassString();
	str = str + "][One Pass Use Limit:" + GetString(this->iUseLimitOnePass);
	str = str + "]";
	return str;
}

CString CLocation::GetString(long iNum)
{
	char buffer[200];
	sprintf(buffer, "%d", iNum);
	return CString(buffer);
}
