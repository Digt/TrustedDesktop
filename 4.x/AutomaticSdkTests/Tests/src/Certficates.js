

//enum FORMAT
var BASE64_TYPE = 0;
var DER_TYPE = 1;

var SYSTEM_STORE_MY           = 1
var SYSTEM_STORE_CA           = 2
var SYSTEM_STORE_ROOT         = 4
var SYSTEM_STORE_ADDRESS_BOOK = 16

var LOCAL_MACHINE_STORE = 0
var CURRENT_USER_STORE  = 1

//-----------------------------------------------------------------------------

//Import, export and saving of a certificate
//Импорт, экспорт и сохранение сертификата

function ImportExportAndSaving(certPath, certSavingPath, bDisplayUI)
{
    if( typeof(bDisplayUI) == "undefined" )
    {
        bDisplayUI = true;
    }

    var oCertificate1 = new ActiveXObject( "DigtCrypto.Certificate" );

    // Загружаем сертификат из файла
    oCertificate1.Load( certPath );

    // Экспортируем сертификат в DER формате
    var szData_Der = oCertificate1.Export(DER_TYPE);

    // Вызываем форму просмотра свойств сертификата
    if( bDisplayUI )
    {
        oCertificate1.Display();
    }
    oCertificate1 = null;

    // Импортируем строку с сертификатом в новый объект сертификата
    var oCertificate2 = new ActiveXObject( "DigtCrypto.Certificate" );
    oCertificate2.Import( szData_Der );

    // Сохраняем сертификат
    oCertificate2.Save( certSavingPath, BASE64_TYPE, 0 );

    return oCertificate2;
}

//-----------------------------------------------------------------------------

//Getting properties of a certificate
//Получение свойств сертификата

function GettingCertProperties( strCertificateFileName, bDisplayUI )
{
    if( typeof(bDisplayUI) == "undefined" )
    {
        bDisplayUI = true;
    }

    var oCert;

    //Если файл для сертификата не задан
    if( (typeof(strCertificateFileName)!="string") || (strCertificateFileName == "") )
    {
        var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");
        var oCerts;

        if( bDisplayUI )
        {
            // Вызов формы просмотра хранилища
            oCerts = oCertStore.Display(SYSTEM_STORE_MY || SYSTEM_STORE_ADDRESS_BOOK);
        }
        else
        {
            // Откроем хранилище личных сертификатов ("my") текущего пользователя
            oCertStore.Open( CURRENT_USER_STORE, "my" );
            oCerts = oCertStore.Store;
        }

        // Проверим, извлечен ли сертификат из хранилища
        if( oCerts.Item.Count == 0 )
        {
            if( bDisplayUI )
            {
                return "Ни одного сертификата не было выбрано";
            }
            else
            {
                return "Хранилище не содержит ни одного сертификата"
            }
        }

        // Извлечение первого сертификата из полученного списка (список никак не отсортирован!)
        oCert = oCerts.Item(0);
    }
    else
    {
        //Иначе считаем сертификат из файла
        oCert = new ActiveXObject("DigtCrypto.Certificate");
        oCert.Load( strCertificateFileName );
    }

    // Просмотрим сертификат
    if( bDisplayUI )
    {
        oCert.Display();
    }

    // Получим свойства сертификатов
    sMess = "";
    sMess += "SerialNumber: " + String(oCert.SerialNumber) + "\n";
    sMess += "IssuerName: " + String(oCert.IssuerName) + "\n";
    sMess += "SubjectName: " + String(oCert.SubjectName) + "\n";
    sMess += "validFrom: " + String(oCert.validFrom) + "\n";
    sMess += "validTo: " + String(oCert.validTo) + "\n";
    sMess += "EKU: " + String(oCert.EKU) + "\n";
    sMess += "ID: " + String(oCert.ID) + "\n";
    sMess += "KU: " + String(oCert.KU) + "\n";
    try
    {
        sMess += "ProviderName: " + String(oCert.ProviderName) + "\n"; // Свойство применимо только для сертификатов, имеющих привязку к контейнеру
        sMess += "ContainerName: " + String(oCert.ContainerName) + "\n"; // Свойство применимо только для сертификатов, имеющих привязку к контейнеру
    }
    catch(err)
    {}
    sMess += "CertContext: " + String(oCert.CertContext) + "\n";
    sMess += "ThumbPrint: " + String(oCert.ThumbPrint) + "\n";
    sMess += "PublicKeyAlg: " + String(oCert.PublicKeyAlg) + "\n";
    sMess += "PublicKey(BASE64_TYPE): " + String(oCert.PublicKey(BASE64_TYPE)) + "\n";

    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( sMess, 0, "Свойства сертификата" );
    }

    // Очистим использованные переменные
    oCert = null;
    oCerts = null;
    oCertStore = null;

    return sMess;
}

//-----------------------------------------------------------------------------

//Setting certificate properties
//Setting of a CryptoAPI context of a certificate
//Установка свойств сертификата 
//Пример установки CryptoAPI-контекста сертификата 

function SettingCryptoApiContext( sCertPath, bDisplayUI )
{
    if( typeof(bDisplayUI) == "undefined" )
    {
        bDisplayUI = true;
    }

    var oCert1 = new ActiveXObject("DigtCrypto.Certificate");

    //Загрузим сертификат из которого получим контекст
    oCert1.Load( sCertPath );

    //Получение свойства CertContext
    var sCertContext = oCert1.CertContext;

    //Создадим новый объект Certificate
    var oCert2 = new ActiveXObject("DigtCrypto.Certificate");

    //Установим ему контекст сертификата
    oCert2.CertContext = sCertContext;

    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( String(oCert2.CertContext), 0, "Контекст сертификата" );
    }

    return oCert2;
}

//-----------------------------------------------------------------------------

//Setting certificate properties
//Setting of a serial number and issuer of a certificate
//Установка свойств сертификата 
//Пример установки серийного номера и издателя сертификата

function SettingSerialAndIssuerToCert( sSerialNumber, sIssuer )
{
    //Создадим новый объект Certificate
    var oCert = new ActiveXObject("DigtCrypto.Certificate");

    //Установим ему новые свойства сертификата
    oCert.SerialNumber = sSerialNumber;
    oCert.IssuerName = sIssuer;

    return oCert;
}

//-----------------------------------------------------------------------------

//Certificate comparsion
//Сравнение сертификатов

function CompareCertificates( strCertificateFileName, bDisplayUI )
{
    if( typeof(bDisplayUI) == "undefined" )
    {
        bDisplayUI = true;
    }

    //Выберем первй сертификат из хранилища личных сертфикатов
    //Открываем хранилище для выбора сертификата
    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");
    oCertStore.Open( CURRENT_USER_STORE, "my" );

    // Вызов формы просмотра хранилища и получение коллекции сертификтов
    var oCerts
    if( bDisplayUI )
    {
        oCerts = oCertStore.Display(1);
    }
    else
    {
        oCerts = oCertStore.Store;
    }

    //Получение сертификата из коллекцииSet
    var oCert = oCerts.Item(0);


    //Получим второй сертификат загрузкой из файла
    // Загрузим сертификат для сравнения
    var oCert1 = new ActiveXObject("DigtCrypto.Certificate");
    oCert1.Load( strCertificateFileName );

    //Просмотрим его
    if( bDisplayUI )
    {
        oCert1.Display();
    }

    //Сравнение сертификатов
    var bResult = oCert.isEqual(oCert1);
    var msgBox = new ActiveXObject("wscript.shell");

    if( bDisplayUI )
    {
        if( bResult == true )
        {
            msgBox.Popup( "Сертификаты равны" );
        }
        else
        {
            msgBox.Popup( "Сертификаты не равны" );
        }
    }

    oCertStore.Close();

    return bResult;
}

//-----------------------------------------------------------------------------

//Verifying certificate status
//Проверка статуса сертификата

function VerifyCertificateStatus( bDisplayUI )
{
    if( typeof(bDisplayUI) == "undefined" )
    {
        bDisplayUI = true;
    }

    var POLICY_TYPE_NONE = 0 //Нет политики использования сертификатов

    var VS_CORRECT = 1
    var VS_UNSUFFICIENT_INFO = 2
    var VS_UNCORRECT = 3
    var VS_INVALID_CERTIFICATE_BLOB = 4
    var VS_CERTIFICATE_TIME_EXPIRIED = 5
    var VS_CERTIFICATE_NO_CHAIN = 6
    var VS_CERTIFICATE_CRL_UPDATING_ERROR = 7
    var VS_LOCAL_CRL_NOT_FOUND = 8
    var VS_CRL_TIME_EXPIRIED = 9
    var VS_CERTIFICATE_IN_CRL = 10
    var VS_CERTIFICATE_IN_LOCAL_CRL = 12
    var VS_CERTIFICATE_CORRECT_BY_LOCAL_CRL = 12
    var VS_CERTIFICATE_USING_RESTRICTED = 13


    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");

    // Открываем хранилище
    oCertStore.Open( CURRENT_USER_STORE, "my" );

    // Вызов формы просмотра хранилища и получение сертификата
    var oCerts;
    if( bDisplayUI )
    {
        oCerts = oCertStore.Display(1);
    }
    else
    {
        oCerts = oCertStore.Store;
    }
    var oCert = oCerts.Item(0);

    var status = -1;
    status = oCert.IsValid(POLICY_TYPE_NONE); //Проверим статус сертификата 
    var strStatus = "Status: ";
    switch( status )
    {
        case VS_CORRECT:
            strStatus += "Корректен";
            break;
        case VS_UNSUFFICIENT_INFO:
            strStatus += "Статус неизвестен";
            break;
        case VS_UNCORRECT:
            strStatus += "Некорректен";
            break;
        case VS_INVALID_CERTIFICATE_BLOB:
            strStatus += "Недействительный блоб сертификата";
            break;
        case VS_CERTIFICATE_TIME_EXPIRIED:
            strStatus += "Время действия сертификата истекло или еще не наступило";
            break;
        case VS_CERTIFICATE_NO_CHAIN:
            strStatus += "Невозможно построить цепочку сертификации";
            break;
        case VS_CERTIFICATE_CRL_UPDATING_ERROR:
            strStatus += "Ошибка обновления сертификата";
            break;
        case VS_LOCAL_CRL_NOT_FOUND:
            strStatus += "Не найден локальный СОС";
            break;
        case VS_CRL_TIME_EXPIRIED:
            strStatus += "Истекло время действия СОС";
            break;
        case VS_CERTIFICATE_IN_CRL:
            strStatus += "Сертфикат находится в СОС";
            break;
        case VS_CERTIFICATE_IN_LOCAL_CRL:
            strStatus += "Сертфикат находится в локальном СОС";
            break;
        case VS_CERTIFICATE_CORRECT_BY_LOCAL_CRL:
            strStatus += "Сертификат действителен по локальному СОС";
            break;
        case VS_CERTIFICATE_USING_RESTRICTED:
            strStatus += "Использование сертификата ограничено";
            break;
        default:
            break;
    }

    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("WScript.Shell");
        msgBox.Popup( strStatus );
    }

    oCertStore.Close();
    return strStatus;
}

//-----------------------------------------------------------------------------
