' WEBSITES I FOUND USEFUL (splitting to fit better):
' http://stackoverflow.com/questions/4971982/vbscript-for-creating-local-account-
'	and-adding-to-admin-group-used-to-work-prior	
' https://msdn.microsoft.com/en-us/library/ms676723(v=vs.85).aspx?cs-save-lang=1&cs-
'	lang=vb#code-snippet-1
' https://msdn.microsoft.com/en-us/library/aa772300%28v=vs.85%29.aspx?f=255&MSPPError=-
'	2147217396
' http://etutorials.org/Server+Administration/Active+directory/Part+III+Scripting+Active+
'	Directory+with+ADSI+ADO+and+WMI/Chapter+21.+Users+and+Groups/21.2+Creating+a+Full-
'	Featured+User+Account/

' set variables/objects ready for createUser procedure
' keeping all these constants here handy, just in case.
Const ADS_UF_SCRIPT = &H1
Const ADS_UF_ACCOUNTDISABLE = &H2
Const ADS_UF_HOMEDIR_REQUIRED = &H8
Const ADS_UF_LOCKOUT = &H10
Const ADS_UF_PASSWD_NOTREQD = &H20
Const ADS_UF_PASSWD_CANT_CHANGE = &H40
Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &H80
Const ADS_UF_TEMP_DUPLICATE_ACCOUNT = &H100
Const ADS_UF_NORMAL_ACCOUNT = &H200
Const ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = &H800
Const ADS_UF_WORKSTATION_TRUST_ACCOUNT = &H1000
Const ADS_UF_SERVER_TRUST_ACCOUNT = &H2000
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
Const ADS_UF_MNS_LOGON_ACCOUNT = &H20000
Const ADS_UF_SMARTCARD_REQUIRED = &H40000
Const ADS_UF_TRUSTED_FOR_DELEGATION = &H80000
Const ADS_UF_NOT_DELEGATED = &H100000
Const ADS_UF_USE_DES_KEY_ONLY = &H200000
Const ADS_UF_DONT_REQUIRE_PREAUTH = &H400000
Const ADS_UF_PASSWORD_EXPIRED = &H800000
Const ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = &H1000000