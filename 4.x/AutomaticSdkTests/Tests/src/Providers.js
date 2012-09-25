
//-----------------------------------------------------------------------------

//Connecting removable key carrier
//Подключение отчуждаемого носителя

function ConnectingKeyCarrier()
{
    //Объявление переменных
    var oCryptoProviders = new ActiveXObject("DigtCrypto.CryptoProviders");

    //Вызов мастера подключения отчуждаемого носителяo
    oCryptoProviders.ConnectRemovableKeyCarrierWizard();
}

//-----------------------------------------------------------------------------

//Getting list of supported cryptoproviders
//Получение списка поддерживаемых криптопровайдеров

function GetSupportedProvidersList( bDisplayUI )
{
    var oCryptoProviders = new ActiveXObject("DigtCrypto.CryptoProviders");
    //Получение коллекции поддерживаемых криптопровайдеров
    var oSupportedProviders = oCryptoProviders.SupportedProviders;

    //Получим количество поддерживаемых DigtCrypto криптопровайдеров, установленных в системе
    var count = oSupportedProviders.Count;

    var msgBox;
    if( bDisplayUI )
    {
        msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "Количество поддерживаемых криптопровайдеров, установленных в системе: " + count );
    }

    var arrProvidersNames = new Array();
    for( var index = 0; index < count; index++ )
    {
        arrProvidersNames.push( oSupportedProviders.Item(index).Name );
    }

    if( bDisplayUI )
    {
        msgBox.Popup( arrProvidersNames.join("\n"), 0, "Список поддерживаемых криптопровайдеров" );
    }

    return arrProvidersNames;
}

//-----------------------------------------------------------------------------

//Getting list of system cryptoproviders
//Получение списка системных криптопровайдеров

function GetSystemCryptoProviders( bDisplayUI )
{
    //Создадим коллекцию доступных криптопровайдеров
    var oCryptoProviders = new ActiveXObject("DigtCrypto.CryptoProviders");

    //Получим коллекцию системных криптопровайдеров
    var oSystemProviders= oCryptoProviders.SystemProviders;

    //Получим количество системных криптопровайдеров в коллекции
    var count = oSystemProviders.Count;

    var msgBox;
    if( bDisplayUI )
    {
        msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "Количество системных криптопровайдеров: " + count );
    }

    // Просмотрим наименования криптопровайдеров, установленных в системе
    var arrProvidersNames = new Array();
    for( var index = 0; index < count; index++ )
    {
        arrProvidersNames.push( oSystemProviders.Item(index).Name );
    }

    if( bDisplayUI )
    {
        msgBox.Popup( arrProvidersNames.join("\n"), 0, "Список системных криптопровайдеров" );
    }

    return arrProvidersNames;
}

//-----------------------------------------------------------------------------

