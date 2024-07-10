#pragma once


class CLocation
{
public:
	CLocation(void);
	~CLocation(void);
protected:
	CString returnEmail;
	int div;
	CString subject;
	CString body;
	
public:
	CString GetReturnEmail(void);
	int GetDiv(void);
	CString GetSubject(void);
	CString GetBody(void);
	
	void SetReturnMail(CString email);
	void SetDiv(int iNum);
	void SetSubject(CString str);
	
	void SetBody(CString str);
protected:
	CString zeroBody;
public:
	void SetZeroBody(CString str);
	CString GetZeroBody(void);
protected:
	CString desDir;
	CString templateDir;
public:
	void SetDesDir(CString sPath);
	void SetTemplateDir(CString sPath);
	CString GetDesDir(void);
	CString GetTemplateDir(void);
protected:
	CString filterStatement;
	CString Replace(CString sSource, CString brokerEmail, CString agentname, CString agentEmail, CString company, CString contact, CString orderid, CString states);
public:
	void SetFilterStatement(CString str);
	CString GetFilterStatement(void);
	CString GetBody(CString brokerEmail, CString agentname, CString agentEmail, CString company, CString contact, CString orderid, CString states);
	
	CString GetSubject(CString brokerEmail, CString agentname, CString agentEmail, CString company, CString contact, CString orderid, CString states);
protected:
	CString sEmailNoteVerifiedPage;
	CString sEmailNoteNonVerifiedPage;
public:
	void SetEmailNoteVerifiedPage(CString sPage);
	void SetEmailNoteNonVerifiedPage(CString sPage);
	CString GetEmailNoteVerifiedPage(void);
	CString GetEmailNoteNonVerifiedPage(void);
	
	
	
protected:
	long iAgeLimitVerified;
	long iUseLimitVerified;
	long iAgeLimitInternet;
	long iUseLimitInternet;
	long iUseLimitOnePass;
	long iAgeLimitOnePass;
public:
	void SetAgeLimitVerified(long iNum);
	void SetUseLimitVerified(long iNum);
	long GetUseLimitVerified(void);
	long GetAgeLimitVerified(void);
	CString GetAgeLimitVerifiedString(void);
	void SetAgeLimitInternet(long iNum);
	long GetAgeLimitInternet(void);
	CString GetAgeLimitInternetString(void);
	void SetUseLimitInternet(long iNum);
	long GetUseLimitInternet(void);
	void SetUseLimitOnePass(long iNum);
	long GetUseLimitOnePass(void);
	void SetAgeLimitOnePass(long iNum);
	long GetAgeLimitOnePass(void);
	CString GetAgeLimitOnePassString(void);
	CString GetLimitsString(void);
	CString GetString(long iNum);
};
