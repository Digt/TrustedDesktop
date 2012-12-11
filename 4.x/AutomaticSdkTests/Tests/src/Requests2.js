
//CA_TYPE
var CA_TYPE_CMS = 1;
var CA_TYPE_CRYPTO_PRO = 2;

//REQUEST_TYPE
var REQUEST_TYPE_REVOCATION = 1;
var REQUEST_TYPE_SUSPENDING = 2;
var REQUEST_TYPE_RESUMING = 4;
var REQUEST_TYPE_UPDATE_CERT = 8;
var REQUEST_TYPE_UPDATE_KEY = 16;
var REQUEST_TYPE_CERT_PKCS10 = 32;
var REQUEST_TYPE_SELF_SIGNED_CERT = 64;
//var REQUEST_TYPE_INFO = &H40000000;

// REQUEST_SENDING_RESULT
var REQUEST_SENDING_SUCCESS = 1;
var REQUEST_SENDING_REQUEST_MISSING = 2;
var REQUEST_SENDING_REQUEST_SOAP_MISSING = 3;
var REQUEST_SENDING_REQUEST_SOAP_INIT_ERROR = 4;
var REQUEST_SENDING_REQUEST_SUBMIT_ERROR = 5;
var REQUEST_SENDING_REQUEST_MISSING_PROXY_PARAMS = 6;

//REQUEST_GENERATION_RESULT
var REQUEST_GENERATION_SUCCESS = 1;
var REQUEST_GENERATION_OPER_CERT_NOT_FOUND = 2;
var REQUEST_GENERATION_OPER_CERT_SN_NOT_FOUND = 3;
var REQUEST_GENERATION_SIGN_REQUEST_FAILED = 4

//REVOCATION_REASON
var REVOCATION_REASON_UNSPECIFIED = 0;
var REVOCATION_REASON_KEY_COMPROMISE = 1;
var REVOCATION_REASON_CA_COMPROMISE = 2;
var REVOCATION_REASON_CHANGE_OF_AFFILIATION = 3;
var REVOCATION_REASON_SUPERSEDED = 4;
var REVOCATION_REASON_CEASE_OF_OPERATION = 5;
var REVOCATION_REASON_CERTIFICATE_HOLD = 6;

//PROFILEEXITFORMAT
var BASE64 = 0;
var DER = 1;
var XML = 2;

// Тип хранилищ
var CURRENT_USER_STORE = 1;
var LOCAL_MACHINE_STORE = 0;

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

//Вызов мастера сбора данных для создания PKI запроса 
//Создание, отправка, сохранение запроса в интерактивном (GUI) режиме.
//Using PKI request creation wizard
//Creating, sending and saving request in interactive mode (GUI).

function CreatePKIRequestWithWizard()
{
    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");
    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");

    //Вызов мастера сбора данных для создания PKI запроса
    oPKIProfile.CollectData();

    //Устанавливаем заполненный профиль для создания PKI запроса
    oPKIRequest.PKIProfile = oPKIProfile;

    //Генерация запроса на приостановление сертификата
    var GenResult = oPKIRequest.Generate();

    var msgBox = new ActiveXObject("wscript.shell");
    msgBox.Popup( GenResult );

    //Если в профиле установлено сохранение запроса в файл - сохранение запроса в файл с установленным именем
    if( oPKIProfile.SaveRequest )
    {
        oPKIRequest.Save( oPKIProfile.RequestFilename );
    }

    //Если в профиле установлена отправка запроса по эл. почте -  выполнение отправки
    if( oPKIProfile.SendRequestByEmail )
    {
        var SendEmailResult = oPKIRequest.SendEmail();       
        msgBox.Popup( SendEmailResult, 0, "Реультат отправки запроса по e-mail" );
    }

    //Если в профиле установлена отправка запроса в УЦ -  выполнение отправки
    if( oPKIProfile.SendRequest )
    {
        var SendResult = oPKIRequest.Send();
        msgBox.Popup( SendResult, 0, "Реультат отправки запроса в УЦ" );
    }
}

//-----------------------------------------------------------------------------

//Создание, сохранение и отправка PKI запроса в Silent режиме 
//Creating, saving and sending PKI request in "silent" mode

function CreateRequestAndSend()
{
    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");
    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");
    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");

    //Для выполнения приостановления сертификата необходимо использовать 3 сертификата:
    //приостанавливаемый сертификат, сертификат подписи запроса, сертфикат для установки ssl соединения.
    //В данном примере в роли этих сертификатов будет использоваться один и тот же сертификат.

    //Примечание: данный сертификат должен иметь назначение "Аутентификация клиента".
    //Для выбора сертификата открываем форму выбора сертификата из личного хранилища

    //Открытие хранилища
    oCertStore.Open( CURRENT_USER_STORE, "my" );

    // Вызов формы просмотра хранилища, получение коллекции с единственным выбранным сертификатом
    var oCerts = oCertStore.Display(1);

    //Получение сертификата из коллекции и установка его в профиль
    //в качестве приостанавливаемого сертификата, сертификата подписи,
    //сертификата установки ssl соединения

    var oSSLCertificate = oCerts.Item(0);
    var oOperationCertificate = oCerts.Item(0);
    var oSignCertificate = oCerts.Item(0);

    oPKIProfile.OperationCertificate = oOperationCertificate;
    oPKIProfile.SignCertificate = oSignCertificate;
    oPKIProfile.SignatureCertificatePin = "1";
    //Установка типа запроса - приостановление сертификата, Типа УЦ - КриптоПРО УЦ
    oPKIProfile.RequestType = REQUEST_TYPE_SUSPENDING;
    oPKIProfile.CAType = CA_TYPE_CRYPTO_PRO;
    oPKIProfile.Comment = "Приостановление";
    oPKIProfile.RevocationReason = REVOCATION_REASON_CERTIFICATE_HOLD;
    oPKIProfile.CertificateHoldDuration = "1-0-0-0-0-0";
    //Установка параметров сохранения запроса
    oPKIProfile.SaveRequest = true;
    oPKIProfile.RequestFilenameFormat = BASE64;
    oPKIProfile.RequestFilename = "Hold.req";
    //Установка параметров отправки запроса в КриптоПРО УЦ
    oPKIProfile.SendRequest = true;
    oPKIProfile.ServiceURL = "https://ca-server/ra/ra.wsdl";
    oPKIProfile.SSLCertificate = oSSLCertificate
    //Установка параметров отправки запроса по E-Mail
    oPKIProfile.SendRequestByEmail = true;
    oPKIProfile.EmailMessage = "Примите запрос на приостановление";
    oPKIProfile.EmailRecipient = "johndoe@ca.example.com";
    oPKIProfile.EmailSubject = "Запрос на приостановление";

    //Установка заполненного профиля для создания запроса
    oPKIRequest.PKIProfile = oPKIProfile
    var GenResult = oPKIRequest.Generate();

    var msgBox = new ActiveXObject("wscript.shell");
    msgBox.Popup( GenResult );

    //Если в профиле установлено сохранение запроса в файл - сохранение запроса в файл с установленным именем
    if( oPKIProfile.SaveRequest )
    {
        oPKIRequest.Save( oPKIProfile.RequestFilename );
    }

    //Если в профиле установлена отправка запроса по эл. почте -  выполнение отправки
    if( oPKIProfile.SendRequestByEmail )
    {
        var SendEmailResult = oPKIRequest.SendEmail();       
        msgBox.Popup( SendEmailResult, 0, "Реультат отправки запроса по e-mail" );
    }

    //Если в профиле установлена отправка запроса в УЦ -  выполнение отправки
    if( oPKIProfile.SendRequest )
    {
        var SendResult = oPKIRequest.Send();
        msgBox.Popup( SendResult, 0, "Реультат отправки запроса в УЦ" );
    }
}

//-----------------------------------------------------------------------------

//Создание запроса на отзыв сертификата в silent режиме 
//Creating request for certificate revocation in "silent" mode

function CreateRevocationRequest( strRevokeReqFileName, bDisplayUI )
{
    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");

    //Для создания запроса на отзыв сертификата необходимо использовать 2 сертификата: 
    //отзываемый сертификат, сертификат подписи запроса.
    //В данном примере в роли этих сертификатов будет использоваться один и тот же сертификат. 

    var oOperationCertificate;
    var oSignCertificate;

    //Для выбора сертификата открываем форму выбора сертификата из личного хранилища

    //Открытие хранилища
    oCertStore.Open( CURRENT_USER_STORE, "my" );

    var oCerts;
    if( bDisplayUI )
    {
        // Вызов формы просмотра хранилища, получение коллекции с единственным выбранным сертификатом
        oCerts = oCertStore.Display(1);
    }
    else
    {
        oCerts = oCertStore.Store;
    }

    //Получение сертификата из коллекции и установка его в профиль в качестве приостанавливаемого сертификата, сертификата подписи.
    oOperationCertificate = oCerts.Item(0);
    oSignCertificate = oCerts.Item(0);

    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");

    oPKIProfile.OperationCertificate = oOperationCertificate;
    oPKIProfile.SignCertificate = oSignCertificate;
    oPKIProfile.SignatureCertificatePin = "1";

    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");

    //Установка типа запроса - приостановление сертификата, Типа УЦ - КриптоПРО УЦ
    oPKIProfile.RequestType = REQUEST_TYPE_SUSPENDING;
    oPKIProfile.CAType = CA_TYPE_CRYPTO_PRO;
    oPKIProfile.Comment = "Отзыв";
    oPKIProfile.RevocationReason = REVOCATION_REASON_KEY_COMPROMISE;
    //Установка параметров сохранения запроса
    //oPKIProfile.SaveRequest = true;
    oPKIProfile.RequestFilenameFormat = DER;
    oPKIProfile.RequestFilename = strRevokeReqFileName;

    oPKIRequest.PKIProfile = oPKIProfile;

    //Генерация запроса
    var GenResult = oPKIRequest.Generate();
    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( GenResult );
    }

    //Сохранение запроса в файл с установленным именем
    oPKIRequest.Save( oPKIProfile.RequestFilename, oPKIProfile.RequestFilenameFormat );

    return GenResult;
}

//-----------------------------------------------------------------------------

//Создание запроса на возобновление сертификата в silent режиме 
//Creating request for resuming certificate

function CreateResumingRequest( strResumeReqFileName, bDisplayUI )
{
    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");
    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");
    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");

    //Для создания запроса на отзыв сертификата необходимо использовать 2 сертификата: 
    //отзываемый сертификат, сертификат подписи запроса.
    //В данном примере в роли этих сертификатов будет использоваться один и тот же сертификат. 

    //Для выбора сертификата открываем форму выбора сертификата из личного хранилища
    //Открытие хранилища
    oCertStore.Open( CURRENT_USER_STORE, "my" );

    var oCerts;
    if( bDisplayUI )
    {
        // Вызов формы просмотра хранилища, получение коллекции с единственным выбранным сертификатом
        oCerts = oCertStore.Display(1);
    }
    else
    {
        oCerts = oCertStore.Store;
    }

    //Получение сертификата из коллекции и установка его в профиль в качестве приостанавливаемого сертификата, сертификата подписи.
    var oOperationCertificate = oCerts.Item(0);
    var oSignCertificate = oCerts.Item(0);
    oPKIProfile.OperationCertificate = oOperationCertificate;
    oPKIProfile.SignCertificate = oSignCertificate;
    oPKIProfile.SignatureCertificatePin = "1";

    //Установка типа запроса - приостановление сертификата, Типа УЦ - КриптоПРО УЦ
    oPKIProfile.RequestType = REQUEST_TYPE_RESUMING;
    oPKIProfile.CAType = CA_TYPE_CRYPTO_PRO;
    oPKIProfile.Comment = "Возобновление";
    //Установка параметров сохранения запроса
    //oPKIProfile.SaveRequest = true;
    oPKIProfile.RequestFilenameFormat = DER;
    oPKIProfile.RequestFilename = strResumeReqFileName;

    oPKIRequest.PKIProfile = oPKIProfile;

    //Генерация запроса
    var GenResult = oPKIRequest.Generate();
    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( GenResult );
    }

    //Сохранение запроса в файл с установленным именем
    oPKIRequest.Save( oPKIProfile.RequestFilename, oPKIProfile.RequestFilenameFormat );

    return GenResult;
}

//-----------------------------------------------------------------------------

//Создание запроса на обновление ключа сертификата в silent режиме 
//Creating request for renewing certificate in "silent" mode

function CreateRenewingRequest( strRenewReqFileName, bDisplayUI )
{
    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");
    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");
    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");

    //Для выполнения обновления ключа сертификата в КриптоПро УЦ необходимо использовать 2 сертификата: 
    //сертификат, ключи которого обновляются, сертификат подписи запроса. 
    //В данном примере в роли этих сертификатов будет использоваться один и тот же сертификат. 
    //Для выбора сертификата открываем форму выбора сертификата из личного хранилища

    //Открытие хранилища
    oCertStore.Open( CURRENT_USER_STORE, "my" );

    var oCerts;
    if( bDisplayUI )
    {
        // Вызов формы просмотра хранилища, получение коллекции с единственным выбранным сертификатом
        oCerts = oCertStore.Display(1);
    }
    else
    {
        oCerts = oCertStore.Store;
    }

    //Получение сертификата из коллекции и установка его в профиль в качестве приостанавливаемого сертификата, сертификата подписи, сертификата установки ssl соединения
    var oOperationCertificate = oCerts.Item(0);
    var oSignCertificate = oCerts.Item(0);
    oPKIProfile.OperationCertificate = oOperationCertificate;

    //Заполнение шаблона запроса на сертификат
    //Получение данных из выбранного сертификата
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.CRequestTemplate" );
    oRequestTemplate.Certificate = oOperationCertificate;
    oRequestTemplate.ExtendedKeyUsage = oOperationCertificate.EKU;
    oRequestTemplate.CryptoProvider = oOperationCertificate.ProviderName;
    oRequestTemplate.KeyUsage = oOperationCertificate.KU;
    oRequestTemplate.CreateNewKeySet = true;
    oRequestTemplate.MarkKeysExportable = true;
    //Необходимо задавать уникальное имя контейнера
    //oRequestTemplate.Keyset = "NewContainer";
    oRequestTemplate.Keyset = GenerateContainerName();

    //Установим шаблон запроса на сертификат в профиль для формирования PKI запроса
    oPKIProfile.RequestTemplate = oRequestTemplate;
    oPKIProfile.SignCertificate = oSignCertificate;
    oPKIProfile.SignatureCertificatePin = "1";

    //Установка типа запроса - приостановление сертификата, Типа УЦ - КриптоПРО УЦ
    oPKIProfile.RequestType = REQUEST_TYPE_UPDATE_KEY;
    oPKIProfile.CAType = CA_TYPE_CRYPTO_PRO;

    //Установка параметров сохранения запроса
    oPKIProfile.SaveRequest = true;
    oPKIProfile.RequestFilenameFormat = DER;
    oPKIProfile.RequestFilename = strRenewReqFileName;

    //Установка профиля создания PKI запроса
    oPKIRequest.PKIProfile = oPKIProfile;

    //Генерация запроса
    var GenResult = oPKIRequest.Generate();
    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( GenResult );
    }

    //Сохранение запроса в файл с установленным именем
    oPKIRequest.Save( oPKIProfile.RequestFilename, oPKIProfile.RequestFilenameFormat );

    return GenResult;
}

//-----------------------------------------------------------------------------

//Создание запроса на сертификат в silent режиме 
//Creating request for certificate in "silent" mode

function CreatingCertificateRequest(reqFileName)
{
    var resultFile;
    if( typeof(reqFileName) == "string" && (reqFileName != "") )
    {
        resultFile = reqFileName;
    }
    else
    {
        resultFile = "cert_req.pem";
    }

    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");
    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");

    //Заполнение шаблона запроса на сертификат
    //Получение данных из выбранного сертификата
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.CRequestTemplate" );
    oRequestTemplate.CN = "John Doe";
    oRequestTemplate.C = "RU";
    oRequestTemplate.E = "johndoe@example.com";
    oRequestTemplate.L = "City";
    oRequestTemplate.O = "Company";
    oRequestTemplate.OU = "Company unit";
    oRequestTemplate.S = "Region";
    oRequestTemplate.ExtendedKeyUsage = "<keyPurposeId>1.3.6.1.5.5.7.3.1</keyPurposeId>";
    oRequestTemplate.CryptoProvider = "Microsoft Base Cryptographic Provider v1.0";
    oRequestTemplate.CryptoProviderType = 1;
    oRequestTemplate.KeyUsage = 1;
    oRequestTemplate.CreateNewKeySet = true;
    oRequestTemplate.ShowSendRequestWindow = true;
    oRequestTemplate.MarkKeysExportable = true;
    oRequestTemplate.KeyLength = 512;
    //Необходимо задавать уникальное имя контейнера вручную или вовсе не задавать его, тогда имя будет сгенерировано автоматически
    //oRequestTemplate.Keyset = "NewContainer";
    oRequestTemplate.Keyset = GenerateContainerName();

    //Установим шаблон запроса на сертификат в профиль для формирования PKI запроса
    oPKIProfile.RequestTemplate = oRequestTemplate;

    //Установка типа запроса - приостановление сертификата, Типа УЦ - КриптоПРО УЦ
    oPKIProfile.RequestType = REQUEST_TYPE_CERT_PKCS10;
    oPKIProfile.CAType = CA_TYPE_CRYPTO_PRO;

    //Установка параметров сохранения запроса
    oPKIProfile.SaveRequest = true;
    oPKIProfile.RequestFilenameFormat = DER;
    oPKIProfile.RequestFilename = resultFile;

    //Установка профиля создания PKI запроса
    oPKIRequest.PKIProfile = oPKIProfile;

    //Генерация запроса
    var GenResult = oPKIRequest.Generate();
//    var msgBox = new ActiveXObject("wscript.shell");
//    msgBox.Popup( GenResult );

    //Сохранение запроса в файл с установленным именем
    if( !GenResult )
    {
        return null;
    }

    oPKIRequest.Save( oPKIProfile.RequestFilename, oPKIProfile.RequestFilenameFormat );
    return oPKIRequest;
}

//-----------------------------------------------------------------------------

//Создание самоподписанного сертификата в silent режиме
//Creating self signed certificate in "silent" mode

function CreateSelfSignedCertificate()
{
    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");
    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");

    //Заполнение шаблона запроса на сертификат
    //Получение данных из выбранного сертификата
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.CRequestTemplate" );
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
    oRequestTemplate.ShowSendRequestWindow = true;
    oRequestTemplate.MarkKeysExportable = true;
    oRequestTemplate.KeyLength = 512;

    //Необходимо задавать уникальное имя контейнера вручную или вовсе не задавать его, тогда имя будет сгенерировано автоматически
    //oRequestTemplate.Keyset = "NewContainer"
    oRequestTemplate.Keyset = GenerateContainerName();

    //Установим шаблон запроса на сертификат в профиль для формирования PKI запроса
    oPKIProfile.RequestTemplate = oRequestTemplate;

    //Установка типа запроса - приостановление сертификата, Типа УЦ - КриптоПРО УЦ
    oPKIProfile.RequestType = REQUEST_TYPE_SELF_SIGNED_CERT;

    //Установка параметров сохранения
    oPKIProfile.SaveRequest = true;
    oPKIProfile.RequestFilenameFormat = DER;
    oPKIProfile.RequestFilename = "selfsigned_cert.cer";

    //Установка профиля создания PKI запроса
    oPKIRequest.PKIProfile = oPKIProfile;

    //Генерация запроса
    var GenResult = oPKIRequest.Generate();
    var msgBox = new ActiveXObject("wscript.shell");
    msgBox.Popup( GenResult );

//    //Сохранение сертификата
//    oPKIRequest.Save( oPKIProfile.RequestFilename, oPKIProfile.RequestFilenameFormat );
}

//-----------------------------------------------------------------------------

