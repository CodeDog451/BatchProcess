#pragma once
#include "stdio.h"
#include "winsock2.h"
#include "socketexception.h"
class Pop3Socket
{
public:
	Pop3Socket(void);
	~Pop3Socket(void);	
	
	// send a message to the server	
	bool sendMessage(CString msg);
	// set the output list box
	void setPrintOut(CListBox* pt_pOut);
	CListBox* pOut;	
private:
	
	// prints a message to the screen list box
	void printOut(CString str);
//////////////////////////////////////////////////////////////////
	public:
	Pop3Socket( char *szRemoteAddr, int iPort);
	Pop3Socket( int iPort);
	Pop3Socket( SOCKET Socket);

	

	Pop3Socket * Accept( void);
	void Close( void);
	int Read( void * pData, unsigned int iLen);
	int Write( void * pData, unsigned int iLen);

private:
	SOCKET m_Socket;
	WSADATA m_WSAData;
	SOCKADDR_IN m_sockaddr;

	void Reset( unsigned int iPort);

//////////////////////////////////////////////////////////////////

public:
	CString GetResponse(void);
};
