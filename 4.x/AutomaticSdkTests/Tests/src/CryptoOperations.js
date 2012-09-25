
//enum FORMAT
var BASE64_TYPE = 0;
var DER_TYPE = 1;

//enim PROFILESTORETYPE (профили)
var REGISTRY_STORE = 0;

//enum DATATYPE (тип данных)
var DT_PLAIN_DATA = 0;
var DT_SIGNED_DATA = 2;
var DT_ENVELOPED_DATA = 3;
var DT_ENCRYPTED_DATA = 6;

//enum WIZARD_TYPE and RESULTTYPE
var SIGN_WIZARD_TYPE = 1;
var ADD_SIGN_WIZARD_TYPE = 2;
var COSIGN_WIZARD_TYPE = 4;
var ENCRYPT_WIZARD_TYPE = 64;
var DECRYPT_WIZARD_TYPE = 1024

//enum CHECKING_WIZARD
var ALL_OK = 0;

//var CERT_FOR_ENCRYPT = 0;
//var CERT_FOR_DECRYPT = 1;
//var CERT_FOR_SIGN = 2;

//-----------------------------------------------------------------------------

//Signature creation
//Создание подписи

function SignatureCreation( sPlainFileName, sSignFileName, bDisplayUI )
{
    // в хранилище должен быть профиль с заполненными параметрами ЭЦП

    var PLAIN_DATA_FILE = sPlainFileName;
    var OUTPUT_DATA_FILE = sSignFileName;

    //Получим профиль по умолчанию или создадим новый, если его нет

    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open( REGISTRY_STORE ); //Открываем хранилище профилей
    var objProfiles = objProfileStore.Store ;//Получаем коллекцию профилей

    var objProfile;
    if( objProfiles.Count > 0 )
    {
        objProfile = objProfiles.DefaultProfile; //Получим профиль по умолчанию
    }
    else
    {
        objProfile = new ActiveXObject("DigtCrypto.Profile"); //Создадим новый профиль
    }

    //Приступим к получению подписи, используя данные, полученные из профиля
    if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
    {
        objProfile.CollectData( SIGN_WIZARD_TYPE ); //Запустим мастер подписи для сбора данных
    }

    var CheckResult = objProfile.CheckData( SIGN_WIZARD_TYPE ); //Проверим, все ли данные собраны
    if( ALL_OK == CheckResult )
    {
        var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
        oPKCS7Message.Profile = objProfile; //Установим профиль с настройками
        oPKCS7Message.Load( DT_PLAIN_DATA, String(PLAIN_DATA_FILE), "" ); //Загрузим исходные данные
        var signResult = oPKCS7Message.Sign();     //Подпишем данные, используя параметры подписи из профиля
        oPKCS7Message.Save( DT_SIGNED_DATA, BASE64_TYPE, OUTPUT_DATA_FILE ); //Сохраним данные 
        oPKCS7Message = null;

        return true;
    }
    else if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "Профиль некорректно заполнен:\n" + CheckResult );
    }

    return false;
}

//-----------------------------------------------------------------------------

//Adding sugnature
//Добавление подписи 

function AddSignature( sSignFile, sResultSignFile, bDisplayUI )
{
    var SIGN_DATA_FILE = sSignFile;
    var OUTPUT_DATA_FILE = sResultSignFile;

    //Получим профиль по умолчанию или создадим новый, если его нет
    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open( REGISTRY_STORE ); //Открываем хранилище профилей
    var objProfiles = objProfileStore.Store; //Получаем коллекцию профилей

    var objProfile;

    if( objProfiles.Count > 0 )
    {
        objProfile = objProfiles.DefaultProfile; //Получим профиль по умолчанию
    }
    else 
    {
        objProfile = new ActiveXObject("DigtCrypto.Profile"); //Создадим новый профиль
    }

    //загрузим файл присоединенной подписи
    var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
    oPKCS7Message.Load( DT_SIGNED_DATA, String(SIGN_DATA_FILE), "" );

    //Получим коллекцию подписей
    var oSigners = oPKCS7Message.Signers;
    if( oSigners.Count > 0 )
    {
        //Установим профиль с настройками
        oPKCS7Message.Profile = objProfile;

        if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
        {
            objProfile.CollectData( ADD_SIGN_WIZARD_TYPE );
        }

        var CheckResult = -10;
        CheckResult = objProfile.CheckData( ADD_SIGN_WIZARD_TYPE );
        if( ALL_OK == CheckResult )
        {
            //Добавим подпись
            oPKCS7Message.Sign();
            //Сохраним данные
            oPKCS7Message.Save( DT_SIGNED_DATA, BASE64_TYPE, OUTPUT_DATA_FILE );
            oPKCS7Message = null;

            return true;
        }
        else
        {
            var msgBox = new ActiveXObject("wscript.shell");
            msgBox.Popup( "Профиль некорректно заполнен" );
        }
    }
    else
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "Нет подписчиков в сообщении" );
    }

    return false;
}

//-----------------------------------------------------------------------------

//Adding cosignature
//Заверение подписи

function AddCosignature( sSignName, sResultSignName, bDisplayUI )
{
    // в хранилище должен быть профиль с заполненными параметрами ЭЦП

    var SIGN_DATA_FILE = sSignName;
    var OUTPUT_DATA_FILE = sResultSignName;

    //Получим профиль по умолчанию или создадим новый, если его нет

    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open( REGISTRY_STORE ); //Открываем хранилище профилей
    var objProfiles = objProfileStore.Store; //Получаем коллекцию профилей

    var objProfile;
    if( objProfiles.Count > 0 )
    {
        objProfile = objProfiles.DefaultProfile; //Получим профиль по умолчанию
    }
    else
    {
        objProfile = new ActiveXObject("DigtCrypto.Profile"); //Создадим новый профиль
    }

    //загрузим файл присоединенной подписи
    var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
    oPKCS7Message.Load( DT_SIGNED_DATA, String(SIGN_DATA_FILE), "" );

    //Получим коллекцию подписей
    var oSignatures = oPKCS7Message.Signatures;
    if( oSignatures.Count > 0 )
    {
        oPKCS7Message.Profile = objProfile;  //Установим профиль с настройками
        if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
        {
            objProfile.CollectData( COSIGN_WIZARD_TYPE );
        }
        var CheckResult = -10;
        CheckResult = objProfile.CheckData( COSIGN_WIZARD_TYPE );
        if( ALL_OK == CheckResult )
        {
            oPKCS7Message.Profile = objProfile;
            //Заверим подпись c порядковым индексом 0
            oPKCS7Message.Cosign( 0 );
            oPKCS7Message.Save( DT_SIGNED_DATA, BASE64_TYPE, OUTPUT_DATA_FILE ); //Сохраним данные 
            oPKCS7Message = null;

            return true;
        }
        else
        {
            var msgBox = new ActiveXObject("wscript.shell");
            msgBox.Popup( "Профиль некорректно заполнен" );
        }
    }
    else
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "Нет подписчиков в сообщении" );
    }

    return false;
}

//-----------------------------------------------------------------------------

//Opening the window of a management of a PKCS7 data  
//Открытие окна управления PKCS7 данными

function OpenPkcs7MsgManagementWindow( sSignFile, sEncryptedFile )
{
    // в хранилище должен быть профиль с заполненными параметрами ЭЦП

    var SINGED_DATA_FILE = sSignFile;
    var ENCRYPTED_DATA_FILE = sEncryptedFile;

    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open( REGISTRY_STORE ); //Открываем хранилище профилей
    var objProfiles = objProfileStore.Store; //Получаем коллекцию профилей

    var objProfile;
    if( objProfiles.Count > 0 )
    {
        objProfile = objProfiles.DefaultProfile; //Получим профиль по умолчанию
    }
    else
    {
        objProfile = new ActiveXObject("DigtCrypto.Profile"); //Создадим новый профиль
    }

    //загрузим файл подписи
    var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
    oPKCS7Message.Profile = objProfile; //Установим профиль с настройками
    oPKCS7Message.Load( DT_SIGNED_DATA, SINGED_DATA_FILE, "" );
    oPKCS7Message.Display();

    //Загрузим файл шифрованных данных
    oPKCS7Message = null;
    oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
    oPKCS7Message.Profile = objProfile; //Установим профиль с настройками
    oPKCS7Message.Load( DT_ENVELOPED_DATA, ENCRYPTED_DATA_FILE, "" );
    oPKCS7Message.Display();
}

//-----------------------------------------------------------------------------

//Data encryption
//Шифрование данных

function EncryptData( sPlainFile, sEncryptedResultFile, bDisplayUI )
{
    // в хранилище должен быть профиль с заполненными параметрами ЭЦП

    var PLAIN_DATA_FILE = sPlainFile;
    var OUTPUT_DATA_FILE = sEncryptedResultFile;

    //Получим профиль по умолчанию или создадим новый, если его нет

    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open( REGISTRY_STORE ); //Открываем хранилище профилей
    var objProfiles = objProfileStore.Store; //Получаем коллекцию профилей

    var objProfile;
    if( objProfiles.Count > 0 )
    {
        objProfile = objProfiles.DefaultProfile; //Получим профиль по умолчанию
    }
    else
    {
        objProfile = new ActiveXObject("DigtCrypto.Profile"); //Создадим новый профиль
    }

    //Приступим к получению подписи, используя данные, полученные из профиля
    if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
    {
        objProfile.CollectData( ENCRYPT_WIZARD_TYPE ); //Запустим мастер подписи для сбора данных
    }
    var CheckResult = objProfile.CheckData(ENCRYPT_WIZARD_TYPE); //Проверим, все ли данные собраны
    if( ALL_OK == CheckResult )
    {
        var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
        oPKCS7Message.Profile = objProfile; //Установим профиль с настройками
        oPKCS7Message.Load( DT_PLAIN_DATA, PLAIN_DATA_FILE, "" ); //Загрузим исходные данные
        oPKCS7Message.Encrypt();
        oPKCS7Message.Save( DT_ENVELOPED_DATA, BASE64_TYPE, OUTPUT_DATA_FILE ); //Сохраним данные 
        oPKCS7Message = null;

        return true;
    }
    else
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "Профиль некорректно заполнен" );
    }

    return false;
}

//-----------------------------------------------------------------------------

//Encryption of data blocks on the same session key (working with data in memory)
//Example 1 Header of the encrypted message is combined with the first block of ciphertext
//Шифрование блоков данных на одном сессионном ключе (работа с данными в памяти)
//Пример 1 Заголовок шифрованного сообщения совмещен с первым блоком шифрованных данных

function EncryptByBlocksFromMemAttachedHeader( fileNameBlock1, fileNameBlock2 )
{
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );

    //Открытие хранилища профилей и получение профиля по умолчанию
    //В профиле должны быть установлены необходимые для выполнения шифрования данные
    var oProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    var oProfile = new ActiveXObject("DigtCrypto.Profile");
    oProfileStore.Open( REGISTRY_STORE );

    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");

    //Установим профиль, с использованием которого будем производить шифрование
    oMsg.Profile = oProfileStore.Store.DefaultProfile;

    //Инициализация объекта
    oMsg.Open( DT_PLAIN_DATA );

    //Импорт в объект первого блока данных
    oMsg.Import( "Первый блок данных" );

    //Сохранение зашифрованного первого блока данных вместе с  первой порцией (заголовком) шифрованного сообщения в файл
    var strStroka = oMsg.Export();
    var file = fso.CreateTextFile( fileNameBlock1, True );
    file.write(strStroka);
    file.close();

    //Импортируем второй блок данных
    oMsg.Import( "Второй блок данных" );

    //Сохранение зашифрованного второго блока данных в файл
    strStroka = oMsg.Export();

    file = fso.CreateTextFile( fileNameBlock1, True );
    file.write(strStroka);
    file.close();
}

//-----------------------------------------------------------------------------

//Encryption of data blocks on the same session key (working with data in memory)
//Example 2 Header of the encrypted message is stored separately from the first block of ciphertext
//Шифрование блоков данных на одном сессионном ключе (работа с данными в памяти)
//Пример 2 Заголовок шифрованного сообщения сохранен отдельно от первого блока шифрованных данных

function EncryptByBlocksFromMemDetachedHeader( fileNameHeader, fileNameBlock1, fileNameBlock2 )
{
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );

    //Открытие хранилища профилей и получение профиля по умолчанию
    //В профиле должны быть установлены необходимые для выполнения шифрования данные
    var oProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    oProfileStore.Open( REGISTRY_STORE );

    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");

    //Установим профиль, с использованием которого будем производить шифрование
    oMsg.Profile = oProfileStore.Store.DefaultProfile;

    //Инициализация объекта
    oMsg.Open( DT_PLAIN_DATA );
    //Сохранение первой порцией (заголовком) шифрованного сообщения в файл
    var strStroka = oMsg.Export();
    var file = fso.CreateTextFile(fileNameHeader, true);
    file.write(strStroka);
    file.close();

    //Импорт в объект первого блока данных
    oMsg.Import( "Первый блок данных" );

    //Сохранение зашифрованного первого блока данных в файл
    strStroka = oMsg.Export();
    file = fso.CreateTextFile(fileNameBlock1, True);
    file.write(strStroka);
    file.close();

    //Импорт в объект второго блока данных
    oMsg.Import( "Второй блок данных" );

    //Сохранение зашифрованного первого блока данных в файл
    strStroka = oMsg.Export();
    file = fso.CreateTextFile(fileNameBlock1, True);
    file.write(strStroka);
    file.close();
}

//-----------------------------------------------------------------------------

//Encryption of data blocks on the same session key (working with files)
//Example 3 Header of the encrypted message is combined with the first block of ciphertext
//Шифрование блоков данных на одном сессионном ключе (работа с файлами)
//Пример 3 Заголовок шифрованного сообщения совмещен с первым блоком шифрованных данных

function EncryptByBlocksFromFileAttachedHeader( fileNameSourceBlock1, fileNameSourceBlock2, fileNameBlock1, fileNameBlock2 )
{
    var oProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    //Открытие хранилища профилей и получение профиля по умолчанию
    //В профиле по умолчанию должны быть установлены необходимые для выполнения шифрования данные
    oProfileStore.Open( REGISTRY_STORE );

    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");
    //Установим профиль, с использованием которого будем производить шифрование
    oMsg.Profile = oProfileStore.Store.DefaultProfile;

    //Инициализация объекта
    oMsg.Open( DT_PLAIN_DATA );

    //Загрузка первого файла с исходыми данными
    oMsg.Load( fileNameSourceBlock1 );
    //Сохранение шифрованного первого файла с исходыми данными вместе с заголовком шифрованного сообщения
    oMsg.Save( fileNameBlock1 );

    //Загрузка второго файла с исходыми данными
    oMsg.Load( fileNameSourceBlock2 );
    //Сохранение шифрованного первого файла с исходыми данными вместе с заголовком шифрованного сообщения
    oMsg.Save( fileNameBlock2 );
}

//-----------------------------------------------------------------------------

//Encryption of data blocks on the same session key (working with files)
//Example 4 Header of the encrypted message is stored separately from the first block of ciphertext
//Шифрование блоков данных на одном сессионном ключе (работа с файлами)
//Пример 4 Заголовок шифрованного сообщения сохранен отдельно от первого блока шифрованных данных

function EncryptByBlocksFromFileDetachedHeader( fileNameSourceBlock1, fileNameSourceBlock2, fileNameHeader, fileNameBlock1, fileNameBlock2 )
{
    var oProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    //Открытие хранилища профилей и получение профиля по умолчанию
    //В профиле по умолчанию должны быть установлены необходимые для выполнения шифрования данные
    oProfileStore.Open( REGISTRY_STORE );

    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");
    //Установим профиль, с использованием которого будем производить шифрование
    oMsg.Profile = oProfileStore.Store.DefaultProfile;

    //Инициализация объекта
    oMsg.Open( DT_PLAIN_DATA );
    //Сохранение заголовка шифрованного сообщения в отдельный файл
    oMsg.Save( fileNameHeader );
    //Загрузка первого файла с исходыми данными
    oMsg.Load( fileNameSourceBlock1 );
    //Сохранение шифрованного первого файла с исходыми данными 
    oMsg.Save( fileNameBlock1 );
    //Загрузка первого файла с исходыми данными
    oMsg.Load( fileNameSourceBlock2 );
    //Сохранение шифрованного первого файла с исходыми данными 
    oMsg.Save( fileNameBlock2 );
}

//-----------------------------------------------------------------------------

//Data decryption
//Расшифрование данных

function DecryptData( sEncryptedFile, sResultDecryptedFile, bDisplayUI )
{
    // в хранилище должен быть профиль с заполненными параметрами ЭЦП

    var INPUT_DATA_FILE = sEncryptedFile;
    var OUTPUT_DATA_FILE = sResultDecryptedFile;

    //Получим профиль по умолчанию или создадим новый, если его нет

    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open( REGISTRY_STORE ); //Открываем хранилище профилей
    var objProfiles = objProfileStore.Store; //Получаем коллекцию профилей

    var objProfile;
    if( objProfiles.Count > 0 )
    {
        objProfile = objProfiles.DefaultProfile; //Получим профиль по умолчанию
    }
    else
    {
        objProfile = new ActiveXObject("DigtCrypto.Profile"); //Создадим новый профиль
    }

    //Приступим к получению подписи, используя данные, полученные из профиля
    if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
    {
        objProfile.CollectData( DECRYPT_WIZARD_TYPE ); //Запустим мастер расшифрования для сбора данных
    }
    var CheckResult = objProfile.CheckData(DECRYPT_WIZARD_TYPE); //Проверим, все ли данные собраны
    if( ALL_OK == CheckResult )
    {
        var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
        oPKCS7Message.Profile = objProfile; //Установим профиль с настройками
        oPKCS7Message.Load( DT_ENVELOPED_DATA, INPUT_DATA_FILE, "" ); //Загрузим исходные данные
        oPKCS7Message.Decrypt();
        oPKCS7Message.Save( DT_PLAIN_DATA, BASE64_TYPE, OUTPUT_DATA_FILE ); //Сохраним данные 
        oPKCS7Message = null;

        return true;
    }
    else
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "Профиль некорректно заполнен" );
    }

    return false;
}

//-----------------------------------------------------------------------------

//Decryption of data blocks on the same session key (working with memory data)
//Example 1 Header of the encrypted message is combined with the first block of ciphertext
//Расшифрование блоков данных на одном сессионном ключе
//Пример 1 Заголовок шифрованного сообщения совмещен с первым блоком шифрованных данных

function DecryptByBlocksFromMemAttachedHeader( fileNameBlock1, fileNameBlock2, fileNameResultBlock1, fileNameResultBlock2 )
{
    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );

    //Инициализация объекта
    oMsg.Open( DT_ENCRYPTED_DATA, fileNameBlock1 );
    //Экспортируем расшифрованные данные в строку
    strStroka1 = oMsg.Export();

    //Сохраним первый блок расшифрованных данных в файл
    var file = fso.CreateTextFile(fileNameResultBlock1, True);
    file.write(strStroka1);
    file.close();
    strStroka1 = "";

    //Считаем в строку второй блок зашифрованных данных
    var file1 = fso.OpenTextFile(fileNameBlock2, 1);
    strStroka = file1.ReadAll();
    file1.Close();

    //Импортируем в объект второй блок данных
    oMsg.Import( strStroka );
    //Экспортируем расшифрованные данные в строку
    strStroka1 = oMsg.Export();

    //Сохраним расшифрованные данные в файл
    file = fso.CreateTextFile(fileNameResultBlock2, True);
    file.write(strStroka1);
    file.close();
}

//-----------------------------------------------------------------------------

//Decryption of data blocks on the same session key (working with memory data)
//Example 2 Header of the encrypted message is stored separately from the first block of ciphertext
//Расшифрование блоков данных на одном сессионном ключе
//Пример 2 Заголовок шифрованного сообщения сохранен отдельно от первого блока шифрованных данных

function DecryptByBlocksFromMemDetachedHeader( fileNameHeader, fileNameBlock1, fileNameBlock2, fileNameResultBlock1, fileNameResultBlock2 )
{
    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );

    //Инициализация объекта с загрузкой заголовка шифрованного сообщения
    oMsg.Open( DT_ENCRYPTED_DATA, fileNameHeader );

    //Считаем в строку первый блок зашифрованных данных
    var file1 = fso.OpenTextFile(fileNameBlock1, 1);
    var strStroka = file1.ReadAll();
    file1.Close();

    //Импортируем в объект первый блок данных
    oMsg.Import( strStroka );
    //Экспортируем расшифрованные данные в строку
    var strStroka1 = oMsg.Export();

    //Сохраним первый блок расшифрованных данных в файл
    var file = fso.CreateTextFile(fileNameResultBlock1, True);
    file.write(strStroka1);
    file.close();
    strStroka1 = "";

    //Считаем в строку второй блок зашифрованных данных
    file1 = fso.OpenTextFile(fileNameBlock2, 1);
    strStroka = file1.ReadAll();
    file1.Close();

    //Импортируем в объект второй блок данных
    oMsg.Import( strStroka );
    //Экспортируем расшифрованные данные в строку
    strStroka1 = oMsg.Export();

    //Сохраним расшифрованных данные в файл
    file = fso.CreateTextFile(fileNameResultBlock2, True);
    file.write(strStroka1);
    file.close();
}

//-----------------------------------------------------------------------------

//Decryption of data blocks on the same session key (working with files)
//Example 3 Header of the encrypted message is combined with the first block of ciphertext
//Расшифрование блоков данных на одном сессионном ключе
//Пример 3 Заголовок шифрованного сообщения совмещен с первым блоком шифрованных данных

function DecryptByBlocksFromFileAttachedHeader( fileNameBlock1, fileNameBlock2, fileNameResultBlock1, fileNameResultBlock2 )
{
    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");
    //Инициализация объекта
    oMsg.Open( DT_ENCRYPTED_DATA, fileNameBlock1 );
    //Сохраняем первый расшифрованный блок данных в файл
    oMsg.Save( fileNameResultBlock1 );
    //Загрузка второго файла с шифрованными  данными
    oMsg.Load( fileNameBlock2 );
    //Сохранение первого расшифрованного блока данных в файл
    oMsg.Save( fileNameResultBlock2 );
}

//-----------------------------------------------------------------------------

//Decryption of data blocks on the same session key (working with files)
//Example 4 Header of the encrypted message is stored separately from the first block of ciphertext
//Расшифрование блоков данных на одном сессионном ключе
//Пример 4 Заголовок шифрованного сообщения сохранен отдельно от первого блока шифрованных данных

function DecryptByBlocksFromFileDetachedHeader( fileNameHeader, fileNameBlock1, fileNameBlock2, fileNameResultBlock1, fileNameResultBlock2 )
{
    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");
    //Инициализация объекта
    oMsg.Open( DT_ENCRYPTED_DATA, fileNameHeader );
    //Загрузка первого файла с  шифрованными  данными
    oMsg.Load( fileNameBlock1 );
    //Сохранение первого расшифрованного блока данных в файл
    oMsg.Save( fileNameResultBlock1 );
    //Загрузка второго файла с шифрованными  данными
    oMsg.Load( fileNameBlock2 );
    //Сохранение первого расшифрованного блока данных в файл
    oMsg.Save( fileNameResultBlock2 );
}

//-----------------------------------------------------------------------------

