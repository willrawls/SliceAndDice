Attribute VB_Name = "CrescentInternetToolPak"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~ SUBJECT:     Crescent Internet ToolPak method/event state and error
'~~~                                                    constants
'~~~ AUTHOR:      Stephen R. Casella
'~~~ DATE:        January 10, 1996
'~~~ MODIFIED:    December 9, 1996
'~~~
'~~~ DESCRIPTION: Crescent Internet ToolPak method/event state constants
'~~~              fired by the EventStateChanged/MethodStateChanged event
'~~~              and found in the EventState/MethodState property.
'~~~              The numbering scheme allows the programmer to monitor
'~~~              multiple protocols within a project.
'~~~ ADDITIONS:   Two additional function were added, WSAErrDescription() and
'~~~              InternetErrDescription().  Both of these funtions will return
'~~~              a long string description of the error code that is passed,
'~~~              and a short name (the constant).
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~ Event State Constants in CIFTP found CIFTP.OCX
Global Const CIFTP_FCLOSED = 100
Global Const CIFTP_SCLOSED = 101
Global Const CIFTP_ACCLOSED = 102
Global Const CIFTP_DCCLOSED = 103
Global Const CIFTP_ACCONN = 104
Global Const CIFTP_DCCONN = 105
Global Const CIFTP_DPORTSET = 106
Global Const CIFTP_LBPOP = 107

'~~~ Method State Constants in CIFTP found CIFTP.OCX
Global Const CIFTP_PASS = 1
Global Const CIFTP_USER = 2
Global Const CIFTP_SYST = 3
Global Const CIFTP_TYPE = 4
Global Const CIFTP_LIST = 5
Global Const CIFTP_QUIT = 6
Global Const CIFTP_PASV = 7
Global Const CIFTP_RETR = 8
Global Const CIFTP_STOR = 9
Global Const CIFTP_CWD = 10
Global Const CIFTP_PWD = 11
Global Const CIFTP_NLST = 12
Global Const CIFTP_CDUP = 13
Global Const CIFTP_APPE = 14

'~~~ Event State Constants in CIHTTP found CIHTTP.OCX
Global Const CIHTTP_LBPOP = 120
Global Const CIHTTP_CONN = 121
Global Const CIHTTP_SCLOSED = 122
Global Const CIHTTP_CONNCLOSED = 123
Global Const CIHTTP_FCLOSED = 124

'~~~ Method State Constants in CIHTTP found CIHTTP.OCX
Global Const CIHTTP_GET = 14
Global Const CIHTTP_HEAD = 15
Global Const CIHTTP_POST = 16

'~~~ Event State Constants in CINEWS found CINEWS.OCX
Global Const CINEWS_LBPOP = 130
Global Const CINEWS_CONN = 131
Global Const CINEWS_SCLOSED = 132
Global Const CINEWS_CONNCLOSED = 133
Global Const CINEWS_FCLOSED = 134
Global Const CINEWS_ARECEIVED = 135

'~~~ Method State Constants in CINEWS found CINEWS.OCX
Global Const CINEWS_ARTICLE = 35
Global Const CINEWS_GROUP = 36
Global Const CINEWS_STAT = 37
Global Const CINEWS_LAST = 38
Global Const CINEWS_LIST = 39
Global Const CINEWS_NEWGROUPS = 40
Global Const CINEWS_NEWNEWS = 41
Global Const CINEWS_NEXT = 42
Global Const CINEWS_POST = 43
Global Const CINEWS_QUIT = 44
Global Const CINEWS_SLAVE = 45
Global Const CINEWS_HEAD = 46
Global Const CINEWS_BODY = 47

'~~~ Winsock error constants
'~~~ Winsock Base Error Value
Global Const WSABASEERR = 10000

'~~~ Base + ...
Global Const WSAEINTR = (WSABASEERR + 4)
Global Const WSAEBADF = (WSABASEERR + 9)
Global Const WSAEACCES = (WSABASEERR + 13)
Global Const WSAEFAULT = (WSABASEERR + 14)
Global Const WSAEINVAL = (WSABASEERR + 22)
Global Const WSAEMFILE = (WSABASEERR + 24)

'~~~ Intentional number sequence skip to
'~~~ keep consistent with Winsock numbering
Global Const WSAEWOULDBLOCK = (WSABASEERR + 35)
Global Const WSAEINPROGRESS = (WSABASEERR + 36)
Global Const WSAEALREADY = (WSABASEERR + 37)
Global Const WSAENOTSOCK = (WSABASEERR + 38)
Global Const WSAEDESTADDRREQ = (WSABASEERR + 39)
Global Const WSAEMSGSIZE = (WSABASEERR + 40)
Global Const WSAEPROTOTYPE = (WSABASEERR + 41)
Global Const WSAENOPROTOOPT = (WSABASEERR + 42)
Global Const WSAEPROTONOSUPPORT = (WSABASEERR + 43)
Global Const WSAESOCKTNOSUPPORT = (WSABASEERR + 44)
Global Const WSAEOPNOTSUPP = (WSABASEERR + 45)
Global Const WSAEPFNOSUPPORT = (WSABASEERR + 46)
Global Const WSAEAFNOSUPPORT = (WSABASEERR + 47)
Global Const WSAEADDRINUSE = (WSABASEERR + 48)
Global Const WSAEADDRNOTAVAIL = (WSABASEERR + 49)
Global Const WSAENETDOWN = (WSABASEERR + 50)
Global Const WSAENETUNREACH = (WSABASEERR + 51)
Global Const WSAENETRESET = (WSABASEERR + 52)
Global Const WSAECONNABORTED = (WSABASEERR + 53)
Global Const WSAECONNRESET = (WSABASEERR + 54)
Global Const WSAENOBUFS = (WSABASEERR + 55)
Global Const WSAEISCONN = (WSABASEERR + 56)
Global Const WSAENOTCONN = (WSABASEERR + 57)
Global Const WSAESHUTDOWN = (WSABASEERR + 58)
Global Const WSAETOOMANYREFS = (WSABASEERR + 59)
Global Const WSAETIMEDOUT = (WSABASEERR + 60)
Global Const WSAECONNREFUSED = (WSABASEERR + 61)
Global Const WSAELOOP = (WSABASEERR + 62)
Global Const WSAENAMETOOLONG = (WSABASEERR + 63)
Global Const WSAEHOSTDOWN = (WSABASEERR + 64)
Global Const WSAEHOSTUNREACH = (WSABASEERR + 65)
Global Const WSAENOTEMPTY = (WSABASEERR + 66)
Global Const WSAEPROCLIM = (WSABASEERR + 67)
Global Const WSAEUSERS = (WSABASEERR + 68)
Global Const WSAEDQUOT = (WSABASEERR + 69)
Global Const WSAESTALE = (WSABASEERR + 70)
Global Const WSAEREMOTE = (WSABASEERR + 71)

'~~~ Intentional number sequence skip to
'~~~ keep consistent with Winsock numbering
Global Const WSASYSNOTREADY = (WSABASEERR + 91)
Global Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)
Global Const WSANOTINITIALISED = (WSABASEERR + 93)

'~~~ Intentional number sequence skip to
'~~~ keep consistent with Winsock numbering
Global Const WSAEDISCON = (WSABASEERR + 101)

'~~~ Intentional number sequence skip to
'~~~ keep consistent with Winsock numbering
Global Const WSAHOST_NOT_FOUND = (WSABASEERR + 1001)
Global Const HOST_NOT_FOUND = WSAHOST_NOT_FOUND
Global Const WSATRY_AGAIN = (WSABASEERR + 1002)
Global Const TRY_AGAIN = WSATRY_AGAIN
Global Const WSANO_RECOVERY = (WSABASEERR + 1003)
Global Const NO_RECOVERY = WSANO_RECOVERY
Global Const WSANO_DATA = (WSABASEERR + 1004)
Global Const NO_DATA = WSANO_DATA
Global Const WSANO_ADDRESS = WSANO_DATA
Global Const NO_ADDRESS = WSANO_ADDRESS

'Default ports for various protocols
Global Const DEFAULT_PORT_SMTP = 25
Global Const DEFAULT_PORT_POP = 110
Global Const DEFAULT_PORT_FTP = 21
Global Const DEFAULT_PORT_HTTP = 80
Global Const DEFAULT_PORT_NNTP = 119

Public Function WSAErrDescription(Err As Integer, Optional ByRef ConstName As Variant) As String

Dim TmpConstName As String, TmpErrDescription As String

Select Case Err
    Case -10000: TmpConstName = "WSABASEERR": TmpErrDescription = "No Error"
    Case 4: TmpConstName = "WSAEINTR": TmpErrDescription = "Interrupted system call"
    Case 9: TmpConstName = "WSAEBADF": TmpErrDescription = "Bad file number"
    Case 13: TmpConstName = "WSAEACCES": TmpErrDescription = "Permission denied"
    Case 14: TmpConstName = "WSAEFAULT": TmpErrDescription = "Bad address"
    Case 22: TmpConstName = "WSAEINVAL": TmpErrDescription = "Invalid argument"
    Case 24: TmpConstName = "WSAEMFILE": TmpErrDescription = "Too many open files"
    Case 35: TmpConstName = "WSAEWOULDBLOCK": TmpErrDescription = "Operation would block"
    Case 36: TmpConstName = "WSAEINPROGRESS": TmpErrDescription = "Operation now in progress"
    Case 37: TmpConstName = "WSAEALREADY": TmpErrDescription = "Operation already in progress"
    Case 38: TmpConstName = "WSAENOTSOCK": TmpErrDescription = "Socket operation on non-socket"
    Case 39: TmpConstName = "WSAEDESTADDRREQ": TmpErrDescription = "Destination address required"
    Case 40: TmpConstName = "WSAEMSGSIZE": TmpErrDescription = "Message too long"
    Case 41: TmpConstName = "WSAEPROTOTYPE": TmpErrDescription = "Protocol wrong type for socket"
    Case 42: TmpConstName = "WSAENOPROTOOPT": TmpErrDescription = "Bad protocol option"
    Case 43: TmpConstName = "WSAEPROTONOSUPPORT": TmpErrDescription = "Protocol not supported"
    Case 44: TmpConstName = "WSAESOCKTNOSUPPORT": TmpErrDescription = "Socket type not supported"
    Case 45: TmpConstName = "WSAEOPNOTSUPP": TmpErrDescription = "Operation not supported on socket"
    Case 46: TmpConstName = "WSAEPFNOSUPPORT": TmpErrDescription = "Protocol family not supported"
    Case 47: TmpConstName = "WSAEAFNOSUPPORT": TmpErrDescription = "Address family not supported by protocol family"
    Case 48: TmpConstName = "WSAEADDRINUSE": TmpErrDescription = "Address already in use"
    Case 49: TmpConstName = "WSAEADDRNOTAVAIL": TmpErrDescription = "Can't assign requested address"
    Case 50: TmpConstName = "WSAENETDOWN": TmpErrDescription = "Network is down"
    Case 51: TmpConstName = "WSAENETUNREACH": TmpErrDescription = "Network is unreachable"
    Case 52: TmpConstName = "WSAENETRESET": TmpErrDescription = "Net dropped connection or reset"
    Case 53: TmpConstName = "WSAECONNABORTED": TmpErrDescription = "Software caused connection abort"
    Case 54: TmpConstName = "WSAECONNRESET": TmpErrDescription = "Connection reset by peer"
    Case 55: TmpConstName = "WSAENOBUFS": TmpErrDescription = "No buffer space available"
    Case 56: TmpConstName = "WSAEISCONN": TmpErrDescription = "Socket is already connected"
    Case 57: TmpConstName = "WSAENOTCONN": TmpErrDescription = "Socket is not connected"
    Case 58: TmpConstName = "WSAESHUTDOWN": TmpErrDescription = "Can't send after socket shutdown"
    Case 59: TmpConstName = "WSAETOOMANYREFS": TmpErrDescription = "Too many references, can't splice"
    Case 60: TmpConstName = "WSAETIMEDOUT": TmpErrDescription = "Connection timed out"
    Case 61: TmpConstName = "WSAECONNREFUSED": TmpErrDescription = "Connection refused"
    Case 62: TmpConstName = "WSAELOOP": TmpErrDescription = "Too many levels of symbolic links"
    Case 63: TmpConstName = "WSAENAMETOOLONG": TmpErrDescription = "File name too long"
    Case 64: TmpConstName = "WSAEHOSTDOWN": TmpErrDescription = "Host is down"
    Case 65: TmpConstName = "WSAEHOSTUNREACH": TmpErrDescription = "No Route to Host"
    Case 66: TmpConstName = "WSAENOTEMPTY": TmpErrDescription = "Directory not empty"
    Case 67: TmpConstName = "WSAEPROCLIM": TmpErrDescription = "Too many processes"
    Case 68: TmpConstName = "WSAEUSERS": TmpErrDescription = "Too many users"
    Case 69: TmpConstName = "WSAEDQUOT": TmpErrDescription = "Disc Quota Exceeded"
    Case 70: TmpConstName = "WSAESTALE": TmpErrDescription = "Stale NFS file handle"
    Case 91: TmpConstName = "WSASYSNOTREADY": TmpErrDescription = "Network SubSystem is unavailable"
    Case 92: TmpConstName = "WSAVERNOTSUPPORTED": TmpErrDescription = "WINSOCK DLL Version out of range"
    Case 93: TmpConstName = "WSANOTINITIALISED": TmpErrDescription = "Successful WSASTARTUP not yet performed"
    Case 71: TmpConstName = "WSAEREMOTE": TmpErrDescription = "Too many levels of remote in path"
    Case 1001: TmpConstName = "WSAHOST_NOT_FOUND": TmpErrDescription = "Host not found"
    Case 1002: TmpConstName = "WSATRY_AGAIN": TmpErrDescription = "Non-Authoritative Host not found"
    Case 1003: TmpConstName = "WSANO_RECOVERY": TmpErrDescription = "Non-Recoverable errors: FORMERR, REFUSED, NOTIMP"
    Case 1004: TmpConstName = "WSANO_DATA or WSANO_ADDRESS": TmpErrDescription = "Valid name, no data record of requested type *OR* No address, look for MX record"
    Case Else: TmpConstName = "": TmpErrDescription = "Unknown Error"
End Select

If Not IsMissing(ConstName) Then
    ConstName = TmpConstName
End If

WSAErrDescription = TmpErrDescription

End Function
Public Function InternetErrDescription(Err As Long, Optional ByRef ConstName As Variant) As String

'~~~ A long description was not found for some of the error numbers,
'~~~ the constant name has been used in it place.

Dim TmpConstName As String, TmpErrDescription As String

Select Case Err
    Case 80: TmpConstName = "": TmpErrDescription = "File Exists"
    Case 87: TmpConstName = "": TmpErrDescription = "Bad or Missing Parameter"
    Case 12001: TmpConstName = "ERROR_INTERNET_OUT_OF_HANDLES": TmpErrDescription = "No more handles could be generated at this time."
    Case 12002: TmpConstName = "ERROR_INTERNET_TIMEOUT": TmpErrDescription = "The request has timed out."
    Case 12003: TmpConstName = "ERROR_INTERNET_EXTENDED_ERROR": TmpErrDescription = "An extended error was returned from the server. This is typically a string or buffer containing a verbose error message. Call InternetGetLastResponseInfo to retrieve the error text."
    Case 12004: TmpConstName = "ERROR_INTERNET_INTERNAL_ERROR ": TmpErrDescription = "An internal error has occurred. "
    Case 12005: TmpConstName = "ERROR_INTERNET_INVALID_URL": TmpErrDescription = "The URL is invalid."
    Case 12006: TmpConstName = "ERROR_INTERNET_UNRECOGNIZED_SCHEME": TmpErrDescription = "The URL scheme could not be recognized, or is not supported. "
    Case 12007: TmpConstName = "ERROR_INTERNET_NAME_NOT_RESOLVED": TmpErrDescription = "The server name could not be resolved."
    Case 12008: TmpConstName = "ERROR_INTERNET_PROTOCOL_NOT_FOUND": TmpErrDescription = "The requested protocol could not be located. "
    Case 12009: TmpConstName = "ERROR_INTERNET_INVALID_OPTION": TmpErrDescription = "A request to InternetQueryOption or InternetSetOption specified an invalid option value."
    Case 12010: TmpConstName = "ERROR_INTERNET_BAD_OPTION_LENGTH": TmpErrDescription = "The length of an option supplied to InternetQueryOption or InternetSetOption is incorrect for the type of option specified."
    Case 12011: TmpConstName = "ERROR_INTERNET_OPTION_NOT_SETTABLE": TmpErrDescription = "The request option can not be set, only queried. "
    Case 12012: TmpConstName = "ERROR_INTERNET_SHUTDOWN": TmpErrDescription = "The Win32 Internet function support is being shut down or unloaded."
    Case 12013: TmpConstName = "ERROR_INTERNET_INCORRECT_USER_NAME ": TmpErrDescription = "The request to connect and log on to an FTP server could not be completed because the supplied user name is incorrect."
    Case 12014: TmpConstName = "ERROR_INTERNET_INCORRECT_PASSWORD ": TmpErrDescription = "The request to connect and log on to an FTP server could not be completed because the supplied password is incorrect."
    Case 12015: TmpConstName = "ERROR_INTERNET_LOGIN_FAILURE ": TmpErrDescription = "The request to connect to and log on to an FTP server failed. "
    Case 12016: TmpConstName = "ERROR_INTERNET_INVALID_OPERATION": TmpErrDescription = "The requested operation is invalid."
    Case 12017: TmpConstName = "ERROR_INTERNET_OPERATION_CANCELLED": TmpErrDescription = "The operation was canceled, usually because the handle on which the request was operating was closed before the operation completed."
    Case 12018: TmpConstName = "ERROR_INTERNET_INCORRECT_HANDLE_TYPE ": TmpErrDescription = "The type of handle supplied is incorrect for this operation."
    Case 12019: TmpConstName = "ERROR_INTERNET_INCORRECT_HANDLE_STATE": TmpErrDescription = "The requested operation cannot be carried out because the handle supplied is not in the correct state."
    Case 12020: TmpConstName = "ERROR_INTERNET_NOT_PROXY_REQUEST": TmpErrDescription = "The request cannot be made via a proxy."
    Case 12021: TmpConstName = "ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND": TmpErrDescription = "A required registry value could not be located."
    Case 12022: TmpConstName = "ERROR_INTERNET_BAD_REGISTRY_PARAMETER": TmpErrDescription = "A required registry value was located but is an incorrect type or has an invalid value."
    Case 12023: TmpConstName = "ERROR_INTERNET_NO_DIRECT_ACCESS": TmpErrDescription = "Direct network access cannot be made at this time."
    Case 12024: TmpConstName = "ERROR_INTERNET_NO_CONTEXT": TmpErrDescription = "An asynchronous request could not be made because a zero context value was supplied."
    Case 12025: TmpConstName = "ERROR_INTERNET_NO_CALLBACK": TmpErrDescription = "An asynchronous request could not be made because a callback function has not been set."
    Case 12026: TmpConstName = "ERROR_INTERNET_REQUEST_PENDING": TmpErrDescription = "The required operation could not be completed because one or more requests are pending."
    Case 12027: TmpConstName = "ERROR_INTERNET_INCORRECT_FORMAT": TmpErrDescription = "The format of the request is invalid."
    Case 12028: TmpConstName = "ERROR_INTERNET_ITEM_NOT_FOUND": TmpErrDescription = "The requested item could not be located."
    Case 12029: TmpConstName = "ERROR_INTERNET_CANNOT_CONNECT": TmpErrDescription = "The attempt to connect to the server failed."
    Case 12030: TmpConstName = "ERROR_INTERNET_CONNECTION_ABORTED": TmpErrDescription = "The connection with the server has been terminated."
    Case 12031: TmpConstName = "ERROR_INTERNET_CONNECTION_RESET": TmpErrDescription = "The connection with the server has been reset."
    Case 12032: TmpConstName = "ERROR_INTERNET_FORCE_RETRY": TmpErrDescription = "Calls for the Win32 Internet function to redo the request."
    Case 12033: TmpConstName = "ERROR_INTERNET_PROXY_REQUEST": TmpErrDescription = "ERROR_INTERNET_PROXY_REQUEST"
    Case 12034: TmpConstName = "ERROR_INTERNET_NEED_UI": TmpErrDescription = "ERROR_INTERNET_NEED_UI"
    
    Case 12036: TmpConstName = "ERROR_INTERNET_HANDLE_EXISTS": TmpErrDescription = "The request failed because the handle already exists."
    Case 12037: TmpConstName = "ERROR_INTERNET_SEC_CERT_DATE_INVALID": TmpErrDescription = "ERROR_INTERNET_SEC_CERT_DATE_INVALID"
    Case 12038: TmpConstName = "ERROR_INTERNET_SEC_CERT_CN_INVALID": TmpErrDescription = "ERROR_INTERNET_SEC_CERT_CN_INVALID"
    Case 12039: TmpConstName = "ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR": TmpErrDescription = "ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR"
    Case 12040: TmpConstName = "ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR": TmpErrDescription = "ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR"
    Case 12041: TmpConstName = "ERROR_INTERNET_MIXED_SECURITY": TmpErrDescription = "Indicates that the content is not entirely secure. Some of the content being viewed may have come from unsecured servers."
    Case 12042: TmpConstName = "ERROR_INTERNET_CHG_POST_IS_NON_SECURE": TmpErrDescription = "ERROR_INTERNET_CHG_POST_IS_NON_SECURE"
    Case 12043: TmpConstName = "ERROR_INTERNET_POST_IS_NON_SECURE": TmpErrDescription = "ERROR_INTERNET_POST_IS_NON_SECURE"
    Case 12044: TmpConstName = "ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED": TmpErrDescription = "ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED"
    Case 12045: TmpConstName = "ERROR_INTERNET_INVALID_CA": TmpErrDescription = "ERROR_INTERNET_INVALID_CA"
    Case 12046: TmpConstName = "ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP": TmpErrDescription = "Client authorization is not set up on this computer"
    Case 12047: TmpConstName = "ERROR_INTERNET_ASYNC_THREAD_FAILED": TmpErrDescription = "ERROR_INTERNET_ASYNC_THREAD_FAILED"
    Case 12048: TmpConstName = "ERROR_INTERNET_REDIRECT_SCHEME_CHANGE": TmpErrDescription = "ERROR_INTERNET_REDIRECT_SCHEME_CHANGE"
    Case 12110: TmpConstName = "ERROR_FTP_TRANSFER_IN_PROGRESS": TmpErrDescription = "The FTP operation was not completed because the session was aborted."
    Case 12111: TmpConstName = "ERROR_FTP_DROPPED": TmpErrDescription = "Non-Recoverable errors: FORMERR, REFUSED, NOTIMP"
    Case Else:  TmpConstName = "": TmpErrDescription = "Unknown Error"
End Select

If Not IsMissing(ConstName) Then
    ConstName = TmpConstName
End If

InternetErrDescription = TmpErrDescription

End Function

