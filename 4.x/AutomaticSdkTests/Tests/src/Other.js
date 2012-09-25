
//-----------------------------------------------------------------------------

//Obtain version of DigtCrypto
//Получение версии DigtCrypto

function GetDigtCryptoVersion( bDisplayUI )
{
    var oUtil = new ActiveXObject("DigtCrypto.Util");
    //Получение версии библиотеки DigtCrypto
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
//Установка лицензии

function SetUpLicense( sNewLicense, bDisplayUI )
{
    //Получение лицензии
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
//Подключение журналирования

function StartUpAndShutDownLogger( bDisplayUI )
{
    var LOG_LEVEL_DISABLED = 0; //Логирование отключено 
    var LOG_LEVEL_ERROR = 1;    //Ошибки 
    var LOG_LEVEL_WARN = 2;     //Предупреждения 
    var LOG_LEVEL_INFO = 3;     //Информация 
    var LOG_LEVEL_DEBUG = 4;    //Отладка 
    var LOG_LEVEL_TRACE = 5;    //Трассировка 

    var strPattern = "log4cplus.appender.toDSL.layout.ConversionPattern=%d [%5t] [%p] [%c] [%x] - %m%n";

    var oUtil = new ActiveXObject("DigtCrypto.Util");


    var sLoggerID = oUtil.StartupLogger(strPattern, LOG_LEVEL_TRACE, 500*1024, "DSS" , 0); // Запуск журналирования

    strLoggerStatement = oUtil.GetLoggerStatements(sLoggerID ); // Получение списка зажурналированных событий

    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( strLoggerStatement );
    }

    oUtil.ShutdownLogger( sLoggerID );

    return strLoggerStatement;
}

//-----------------------------------------------------------------------------

