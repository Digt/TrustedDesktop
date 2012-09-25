
//-----------------------------------------------------------------------------

//Obtain version of DigtCrypto
//��������� ������ DigtCrypto

function GetDigtCryptoVersion( bDisplayUI )
{
    var oUtil = new ActiveXObject("DigtCrypto.Util");
    //��������� ������ ���������� DigtCrypto
    var sVersion = oUtil.Version;

    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( sVersion, 0, "Version of DigtCrypto" );
    }

    return sVersion;
}

//-----------------------------------------------------------------------------

//Installing license
//��������� ��������

function SetUpLicense( sNewLicense, bDisplayUI )
{
    //��������� ��������
    var oUtil = new ActiveXObject("DigtCrypto.Util");
    var sOldLicense = oUtil.License;

    if( typeof(sNewLicense) == "string" && sNewLicense != "" )
    {
        oUtil.License = sNewLicense;
    }

    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "New license: " + sNewLicense + "\nOld license: " + sOldLicense, 0, "Licenses" );
    }

    return sOldLicense;
}

//-----------------------------------------------------------------------------

//Turning on logging
//����������� ��������������

function StartUpAndShutDownLogger( bDisplayUI )
{
    var LOG_LEVEL_DISABLED = 0; //����������� ��������� 
    var LOG_LEVEL_ERROR = 1;    //������ 
    var LOG_LEVEL_WARN = 2;     //�������������� 
    var LOG_LEVEL_INFO = 3;     //���������� 
    var LOG_LEVEL_DEBUG = 4;    //������� 
    var LOG_LEVEL_TRACE = 5;    //����������� 

    var strPattern = "log4cplus.appender.toDSL.layout.ConversionPattern=%d [%5t] [%p] [%c] [%x] - %m%n";

    var oUtil = new ActiveXObject("DigtCrypto.Util");


    var sLoggerID = oUtil.StartupLogger(strPattern, LOG_LEVEL_TRACE, 500*1024, "DSS" , 0); // ������ ��������������

    strLoggerStatement = oUtil.GetLoggerStatements(sLoggerID ); // ��������� ������ ����������������� �������

    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( strLoggerStatement );
    }

    oUtil.ShutdownLogger( sLoggerID );

    return strLoggerStatement;
}

//-----------------------------------------------------------------------------

