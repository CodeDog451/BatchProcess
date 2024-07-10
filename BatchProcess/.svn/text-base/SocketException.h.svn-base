#pragma once

#include "stdio.h"
#include "winsock2.h"

class CSocketException
{
public:
	CSocketException( char * szText)
	{
		strcpy( m_szText, szText);
	}

	~CSocketException(){};

	char * getText(){ return( m_szText);}

private:
	char m_szText[ 128];
};
