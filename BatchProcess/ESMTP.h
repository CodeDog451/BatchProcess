// ESMTP.h: interface for the ESMTP class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_ESMTP_H__A614358F_DAC7_4ECD_B11B_43686221109D__INCLUDED_)
#define AFX_ESMTP_H__A614358F_DAC7_4ECD_B11B_43686221109D__INCLUDED_
#include "SMTP.h"
//#include "logdb.h"
#include ".\stringhelper.h"
#include ".\writelog.h"
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

const int NO_LOGIN = 0;
const int CRAM_MD5 = 1;
const int AUTH_LOGIN = 2;
const int LOGIN_PLAIN = 3;
class ESMTP  
{
public:
	void SetBody(CString body);
	CSMTPMessage* CreateMessage();
	bool Send();
	void SetPassword(CString password);
	void SetUserName(CString username);
	void SetAuthenticate(int authlevel);
	void SetName(CString name);
	void SetAddress(CString address);

protected:
	CSMTPMessage* pMessage;	
	CPJNSMTPConnection connection;
	CString	m_sBCC;
	CString	m_sBody;
	CString	m_sCC;
	CString	m_sFile;
	CString	m_sSubject;
	CString	m_sTo;
	BOOL	m_bDirectly;
	CString	m_sAddress;
	CString	m_sHost;
	CString	m_sName;
	int m_nPort;
	CPJNSMTPConnection::LoginMethod m_Authenticate;
	CString m_sUsername;
	CString m_sPassword;
	BOOL    m_bAutoDial;
	CString m_sBoundIP;
	CString m_sEncodingFriendly;
	CString m_sEncodingCharset;
	BOOL    m_bMime;
	BOOL    m_bHTML;
	CSMTPMessage::PRIORITY m_Priority;

public:
	ESMTP();
	virtual ~ESMTP();

protected:
	HWND mainWin;
	CArray<CString,CString&> myAttach;
	//CLogDB myLog;
public:
	void SetParentWindow(HWND m_hwnd);
	void PrintOut(CString str);
	void AddFilename(CString sFilename);
	void ClearAttach(void);
	void PrintFail(CString str);
protected:
	int iDiv;
public:
	void SetDiv(int iLoc);
	void WriteFailLog(CString str);
	CString GetCString(_bstr_t bstrStart);
	// set who the email is going to
	void SetTo(CString str);
	// set the blind carbon copies
	void SetBCC(CString str);
	// set the subject line of the email
	void SetSubject(CString str);
	void SetHost(CString str);
	// get the email addres the email is being sent too
	CString GetTo(void);
	// get the BCC
	CString GetBCC(void);
};

#endif // !defined(AFX_ESMTP_H__A614358F_DAC7_4ECD_B11B_43686221109D__INCLUDED_)
