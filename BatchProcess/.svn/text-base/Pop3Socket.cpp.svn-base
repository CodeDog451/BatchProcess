#include "StdAfx.h"
#include ".\pop3socket.h"


// set the output list box
void Pop3Socket::setPrintOut(CListBox* pt_pOut)
{
	pOut = pt_pOut;
}

// prints a message to the screen list box
void Pop3Socket::printOut(CString str)
{
	if(pOut)
	{
		pOut->AddString(str);
	}
}

// send a message to the server
bool Pop3Socket::sendMessage(CString msg)
{
	CString str;	
	bool success = false;

	//AfxMessageBox("Put your send socket code here", MB_OK);
	success = true;// if the send went ok set this to true if it didn't go through set it to false
	if(success)
	{
		str = msg;
	}
	else
	{
		str = "There was an error sending your message to the server";
	}
	printOut(str);


	return success;
}
////////////////////////////////////////////////////////////////////////
/**
 * @param iPort local listenning port
 * @throws CSocketException is server socket could not be created
 */
Pop3Socket::Pop3Socket( int iPort)
: pOut(NULL)
{
	Reset( iPort);

	if( bind( m_Socket, ( SOCKADDR *)&m_sockaddr, sizeof( sockaddr)) != 0)
		throw CSocketException( "bind() failed");

	if( listen( m_Socket, 0) != 0)
		throw CSocketException( "accept() failed");
}

/**
 * @param szRemoteAddr Remote Machine Address
 * @param iRemotePort Server Listenning Port
 * @throws CSocketException if client socket could not be created
 */
Pop3Socket::Pop3Socket( char *szRemoteAddr, int iPort)
: pOut(NULL)
{
	if( !szRemoteAddr)
		throw CSocketException( "Invalid parameters");

	Reset( iPort);

	// first guess => try to resolve it as IP@
	m_sockaddr.sin_addr.s_addr = inet_addr( szRemoteAddr);
	if( m_sockaddr.sin_addr.s_addr == INADDR_NONE)
	{	// screwed => try to resolve it as name
	LPHOSTENT lpHost = gethostbyname( szRemoteAddr);
		if( !lpHost)
			throw CSocketException( "Unable to solve this address");
		m_sockaddr.sin_addr.s_addr = **(int**)(lpHost->h_addr_list);
	}

	// actually performs connection
	if( connect( m_Socket, ( SOCKADDR*)&m_sockaddr, sizeof( sockaddr)) != 0)
		throw CSocketException( "connect() failed");
	

}

/**
 * Create a socket for data transfer (typically after Accept)
 * @param Socket the socket descriptor for this new object
 */
Pop3Socket::Pop3Socket( SOCKET Socket)
: pOut(NULL)
{
	m_Socket = Socket;
}

/**
 * Destructor
 */
Pop3Socket::~Pop3Socket()
{
	Close();
}

/**
 * Wait for incomming connections on server socket
 * @return CSocket new data socket for this incomming client. Can be NULL is anything went wrong
 */
Pop3Socket * Pop3Socket::Accept()
{
int nlen = sizeof( sockaddr);
SOCKET Socket = accept( m_Socket, ( SOCKADDR *)&m_sockaddr, &nlen);

	if( Socket == -1)
		return( NULL);

	return( new Pop3Socket( Socket));
}


/**
 * Close current socket
 */
void Pop3Socket::Close()
{
	if( m_Socket != INVALID_SOCKET)
		closesocket( m_Socket);
}

/**
 * Read data available in socket or waits for incomming informations
 * @param pData Buffer where informations will be stored
 * @param iLen Max length of incomming data
 * @return Number of bytes read or -1 if anything went wrong
 */
int Pop3Socket::Read( void * pData, unsigned int iLen)
{
	if( !pData || !iLen)
		return( -1);

	return( recv( m_Socket, ( char *)pData, iLen, 0));
}

/**
 * Initialisation common to all constructors
 */
void Pop3Socket::Reset( unsigned int iPort)
{
	// Initialize winsock
	if( WSAStartup( MAKEWORD(2,0), &m_WSAData) != 0)
		throw CSocketException( "WSAStartup() failed");

	// Actually create the socket
	m_Socket = socket( PF_INET, SOCK_STREAM, IPPROTO_TCP);
	if( m_Socket == INVALID_SOCKET)
		throw CSocketException( "socket() failed");

	//ULONG icmd = 0;   
	//int status = ioctlsocket(m_Socket,FIONBIO,&icmd);
	// sockaddr initialisation
	memset( &m_sockaddr, 0, sizeof( sockaddr));

	m_sockaddr.sin_family = AF_INET;
	m_sockaddr.sin_port = htons( iPort);
	m_sockaddr.sin_addr.s_addr = INADDR_ANY;
}

/**
 * @param pData Buffer to be sent
 * @param iSize Number of bytes to be sent from buffer
 * @return the number of sent bytes or -1 if anything went wrong
 */
int Pop3Socket::Write( void * pData, unsigned int iSize)
{
	if( !pData || !iSize)
		return( -1);

	return( ( int)send( m_Socket, ( LPCSTR)pData, iSize, 0));
}
///////////////////////////////////////////////////////////////////////
CString Pop3Socket::GetResponse(void)
{
	CTime startTime = CTime::GetCurrentTime();

	// ... perform time-consuming task ...

	

	CString str = "";
	bool done = false;
	
	while(!done)
	{
		CTime endTime = CTime::GetCurrentTime();

		CTimeSpan elapsedTime = endTime - startTime;
		if(elapsedTime.GetTotalSeconds()>1)
			break;
		
		// read whatever is available in the socket
		char szText[1024];
		memset( szText, 0, sizeof( szText));
		int iLen = Read( szText, sizeof( szText));
		if (iLen != -1)
		{
			
			str = str + szText;
			//Lock();		
			
			
			if(str.MakeLower().Find("+ok bye-bye.\r\n") >=0)
			{
				//pSocket->pOut->AddString("Got a Bye Bye");
				//PrintOut("Got a Bye Bye");
				Close();
			}
			if(str.GetLength()>0)
			{
				if(str.MakeLower().Find("\n") >=0)
				{
					//PrintOut(str);
					done = true;
				}
				//pSocket->pOut->AddString(str);
				
			}
		}
	}
	return str;
}
