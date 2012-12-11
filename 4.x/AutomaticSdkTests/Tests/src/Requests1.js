
//-----------------------------------------------------------------------------

//TODO: Remove after implementation of an automaic request and keys removing
function GenerateContainerName()
{
    var strNumber = "" + Math.ceil(100000 * Math.random());

    while( strNumber.length < 5 )
    {
        strNumber = "0" + strNumber;
    }
    
    return ("deleteme_" + strNumber);
}

//-----------------------------------------------------------------------------

//Создание запроса на сертификат в интерактивном режиме
//Creating certificate request in interactive mode

function CreateCertificateRequestInteractive()
{
    var oRequestCertificate = new ActiveXObject( "DigtCrypto.Request" );

    //Вызов мастера создания запроса на сертификат
    oRequestCertificate.Display();

    oRequestCertificate = null;
}

//-----------------------------------------------------------------------------

//Отправка запроса в Удостоверяющий центр
//Sending request into certification authority (CA)

function SendRequestIntoCA()
{
    //const CA_ADDRESS = "172.17.2.72"
    var CA_ADDRESS = "localhost";
    //var CA_ADDRESS = "http://www.cryptopro.ru/certsrv/";

    //Заполним шаблон запроса на сертификат
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.RequestTemplate" );
    oRequestTemplate.CN = "John Doe";
    oRequestTemplate.C = "RU";
    oRequestTemplate.E = "johndoe@example.com";
    oRequestTemplate.L = "Город";
    oRequestTemplate.O = "Компания";
    oRequestTemplate.OU = "Отдел";
    oRequestTemplate.S = "Регион";
    oRequestTemplate.ExtendedKeyUsage = "<keyPurposeId>1.3.6.1.5.5.7.3.1</keyPurposeId>";
    oRequestTemplate.CryptoProvider = "Microsoft Base Cryptographic Provider v1.0";
    oRequestTemplate.CryptoProviderType = 1;
    oRequestTemplate.KeyUsage = 1;
    oRequestTemplate.CreateNewKeySet = true;
    oRequestTemplate.Keyset = GenerateContainerName();
    oRequestTemplate.ShowSendRequestWindow = true;
    oRequestTemplate.MarkKeysExportable = true;
    oRequestTemplate.KeyLength = 512;
    //WTF: Unsupported property
//    oRequestTemplate.HashAlgorithm = "1.3.14.3.2.26";

    //Установим шаблон запроса на сертификат
    var oRequest = new ActiveXObject("DigtCrypto.Request");
    oRequest.Template = oRequestTemplate;

    //Сгенерируем запрос
    oRequest.Generate();

    oRequestTemplate = null;

    //Получим параметры шаблона
    //oRequestTemplate = new ActiveXObject( "DigtCrypto.RequestTemplate" );
    oRequestTemplate = oRequest.Template;

    var sResult = "";
    sResult += "CN = " + oRequestTemplate.CN + "\n";
    sResult += "C = " + oRequestTemplate.C + "\n";
    sResult += "E = " + oRequestTemplate.E + "\n";
    sResult += "L = " + oRequestTemplate.L + "\n";
    sResult += "O = " + oRequestTemplate.O + "\n";
    sResult += "OU = " + oRequestTemplate.OU + "\n";
    sResult += "S = " + oRequestTemplate.S + "\n";
    sResult += "ExtendedKeyUsage = " + oRequestTemplate.ExtendedKeyUsage + "\n";
    sResult += "CryptoProvider = " + oRequestTemplate.CryptoProvider + "\n";
    sResult += "KeyUsage = " + oRequestTemplate.KeyUsage + "\n";
    sResult += "MarkKeysExportable = " +  oRequestTemplate.MarkKeysExportable + "\n";
    sResult += "KeyLength = " + oRequestTemplate.KeyLength + "\n";
    sResult += "HashAlgorithm = " + oRequestTemplate.HashAlgorithm;

    var msgBox = new ActiveXObject("wscript.shell");
    msgBox.Popup( sResult, 0, "Параметры шаблона запроса на сертификат" );

    //Загрузим сертификат корневого ЦС, где будет обрабатываться запрос 
    var oIssuerCertificate = new ActiveXObject("DigtCrypto.Certificate");
    oIssuerCertificate.Load( "root.cer" );

    //Отправим запрос на обработку в ЦС
    sNewCert = oRequest.Send(CA_ADDRESS, oIssuerCertificate, 1);

    //Проверим код статуса обработки запроса. Если статус не равен 5, то обработка прошла успешно 
    if( oRequest.CADisposition != 5)
    {
        var oCert = new ActiveXObject("DigtCrypto.Certificate");
        oCert.Import( String(sNewCert) );
        oCert.Display();
    }
    else
    {
        msgBox.Popup( sResult, 0, "УЦ настроен на отложенную обработку запроса." );
    }
}

//-----------------------------------------------------------------------------

//Создание запроса на сертификат с использованием шаблона
//Шаблон в файле template.xml
//Request template creating with template
//Template stored in file template.xml

function CreateRequestWithTemplate( strRequestFileName, strResultReqFile, bDisplayUI )
{
    var oRequestCertificate = new ActiveXObject( "DigtCrypto.Request" );
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.RequestTemplate" );

    //Загрузка файла шаблона запроса на сертификат из XML файла
    oRequestTemplate.LoadXMLTemplate(strRequestFileName);
    
    //Perceptible container name for automatic tests (and manual container removing :] )
    oRequestTemplate.Keyset = GenerateContainerName();

    //Вызов мастера создания запроса на сертификат
    oRequestCertificate.Template = oRequestTemplate;

    //Вызов мастера генерации запроса на сертификат.
    //Параметры идентификационной информации и параметры ключа заполняются данными из шаблона.
    if( bDisplayUI )
    {
        oRequestCertificate.Display();
    }

    oRequestCertificate.Generate();
    oRequestCertificate.Save( strResultReqFile, 0 );

    //oRequestCertificate = null;
    oRequestTemplate = null;

    return oRequestCertificate;
}

//-----------------------------------------------------------------------------

//Проверка статуса запроса
//Метод Retry (получение статуса обработки по идентификатору запроса)
//Request status checkup
//Method Retry (checking status of processing by request identifier)

function CheckRequestStatusByID()
{
    var ROOT_ISSUER = "CN=УЦ 2001,OU=Unit,O=Digt,L=Yoshkar-Ola,S=Марий эл,C=RU,E=ca@mail.ru,";
    var CA_ADDRESS = "172.17.2.72";

    //Создаем шаблон запроса на сертификат
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.RequestTemplate" );
    oRequestTemplate.CN = "John Doe";
    oRequestTemplate.C = "RU";
    oRequestTemplate.E = "johndoe@example.com";
    oRequestTemplate.ExtendedKeyUsage = "<KEYPURPOSEID>1.3.6.1.5.5.7.3.1</KEYPURPOSEID>";
    oRequestTemplate.CryptoProvider = "Microsoft Base Cryptographic Provider v1.0";
    oRequestTemplate.CryptoProviderType = 1;
    oRequestTemplate.KeyUsage = 1;
    oRequestTemplate.CreateNewKeySet = true;
    oRequestTemplate.Keyset = GenerateContainerName();

    //Установим шаблон запроса на сертификат
    var oRequest = new ActiveXObject("DigtCrypto.Request");
    oRequest.Template = oRequestTemplate;
    //Сгенерируем запрос
    oRequest.Generate();

    //Загрузим сертификат корневого ЦС     
    var oIssuerCertificate = new ActiveXObject("DigtCrypto.Certificate");
    oIssuerCertificate.Load( "root.cer" );

    //Отправим запрос на обработку в ЦС (ЦС настроен на отложенную обработку запроса)
    //В качестве ответа получим идентификатор запроса
    sNewCert = oRequest.Send(CA_ADDRESS, oIssuerCertificate, 1);

    //Здесь должна быть произведена ручная обработка запроса в ЦС
    //MsgBox "Идентификатор запроса: "+String(sNewCert)+ "Выпустите сертификат по данному запросу в ЦС"
    var msgBox = new ActiveXObject("wscript.shell");
    msgBox.Popup( "Идентификатор запроса: " + String(sNewCert) + "Выпустите сертификат по данному запросу в ЦС" );

    //Проверим статус запроса и получен выпущенный сертификат
    var sNewCert1 = oRequest.Retry(sNewCert, CA_ADDRESS, oIssuerCertificate, 1);

    var oCert = new ActiveXObject("DigtCrypto.Certificate") ;
    oCert.Import( String(sNewCert1) );
    oCert.Display();
}

//-----------------------------------------------------------------------------


//Проверка статуса запроса
//Метод RetrievePending (получение статуса обработки по запросу)
//Request status checkup
//Method RetrievePending (checking status of processing by request)

function CheckRequestStatusByRequest()
{
    //Описание переменных и констант

    var CURRENT_USER_STORE = 1;

    //Откроем хранилище запросов и выберем оттуда запрос, статус которого хотим проверить
    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");
    oCertStore.Open( CURRENT_USER_STORE, "request" );
    // Вызов формы просмотра хранилища
    var oCerts = oCertStore.Display(32, 0, 0);
    var oCertTemplate = oCerts.Item(0);

    //Запросим статус обработки запроса в УЦ
    var oRequest = new ActiveXObject("DigtCrypto.Request");
    var sNewCert1 = oRequest.RetrievePending(oCertTemplate, 1);

    //Если запрос обработан, то просмотрим полученный сертификат
    if( oRequest.CADisposition == 3 )
    {
        var oCert = new ActiveXObject("DigtCrypto.Certificate");
        oCert.Import( String(sNewCert1) );
        oCert.Display();
    }
    else
    {
        //MsgBox "Запрос не обработан"
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "Запрос не обработан" );
    }
}

//-----------------------------------------------------------------------------

//Экспорт, импорт и сохранение запроса на сертификат 
//Export, import and saving certificate request

function RequestExportAndImport(reqFileName)
{
    var DER_TYPE = 1;

    //Создаем шаблон запроса на сертификат
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.RequestTemplate" );

    //Заполним шаблон запроса на сертификат
    oRequestTemplate.CN = "John Doe";
    oRequestTemplate.C = "RU";
    oRequestTemplate.E = "johndoe@example.com";
    oRequestTemplate.ExtendedKeyUsage = "<keyPurposeId>1.3.6.1.5.5.7.3.1</keyPurposeId>";
    oRequestTemplate.CryptoProvider = "Microsoft Base Cryptographic Provider v1.0";
    oRequestTemplate.CryptoProviderType = 1;
    oRequestTemplate.KeyUsage = 1;
    oRequestTemplate.CreateNewKeySet = true;
    oRequestTemplate.Keyset = GenerateContainerName();

    //Установим шаблон запроса на сертификат
    var oRequest = new ActiveXObject("DigtCrypto.Request");
    oRequest.Template = oRequestTemplate;

    //Переходим к генерации запроса"
    oRequest.Generate();

    //Экспортируем запрос на сертификат в строку
    szData = oRequest.Export(DER_TYPE);

    oRequest = null;

    //Импортируем запрос на сертификат из строки
    var oRequestImport = new ActiveXObject( "DigtCrypto.Request" );
    oRequestImport.Import( szData );

    //Сохраним запрос на сертификат в формате DER
    //oRequestImport.Save( "CertificateReq_der.pem", DER_TYPE );
    var resultFile = "";
    if( typeof(reqFileName) == "string" )
    {
        resultFile = reqFileName;
    }

    if( String(resultFile).length == 0 )
    {
        resultFile = "CertificateReq_der.pem";
    }

    oRequestImport.Save( resultFile, DER_TYPE );
    //oRequestImport = null;
    return oRequestImport;
}

//-----------------------------------------------------------------------------

//Создание самоподписанного сертификата
//Creating self signed certificate

function CreateSelfSignedCert()
{
    //var oCert = new ActiveXObject("DigtCrypto.Certificate");

    //Создаем шаблон запроса на сертификат
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.RequestTemplate" );
    oRequestTemplate.CN = "John Doe";
    oRequestTemplate.C = "RU";
    oRequestTemplate.E = "johndoe@example.com";
    oRequestTemplate.ExtendedKeyUsage = "<keyPurposeId>1.3.6.1.5.5.7.3.1</keyPurposeId>";
    oRequestTemplate.CryptoProvider = "Microsoft Base Cryptographic Provider v1.0";
    oRequestTemplate.CryptoProviderType = 1;
    oRequestTemplate.KeyUsage = 1;
    oRequestTemplate.CreateNewKeySet = true;
    oRequestTemplate.Keyset = GenerateContainerName();

    //Установим шаблон запроса на сертификат
    var oRequest = new ActiveXObject("DigtCrypto.Request");
    oRequest.Template = oRequestTemplate;

    //Сгенерируем самоподписанный сертификат 
    var oCert = oRequest.CreateSelfSignedCertificate();
    //var oCert = oRequest.GenerateSelfSigned();

    oCert.Display();
}

//-----------------------------------------------------------------------------

