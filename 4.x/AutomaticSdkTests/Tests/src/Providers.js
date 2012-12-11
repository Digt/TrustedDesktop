
//-----------------------------------------------------------------------------

//Connecting removable key carrier
//����������� ������������ ��������

function ConnectingKeyCarrier()
{
    //���������� ����������
    var oCryptoProviders = new ActiveXObject("DigtCrypto.CryptoProviders");

    //����� ������� ����������� ������������ ��������o
    oCryptoProviders.ConnectRemovableKeyCarrierWizard();
}

//-----------------------------------------------------------------------------

//Getting list of supported cryptoproviders
//��������� ������ �������������� �����������������

function GetSupportedProvidersList( bDisplayUI )
{
    var oCryptoProviders = new ActiveXObject("DigtCrypto.CryptoProviders");
    //��������� ��������� �������������� �����������������
    var oSupportedProviders = oCryptoProviders.SupportedProviders;

    //������� ���������� �������������� DigtCrypto �����������������, ������������� � �������
    var count = oSupportedProviders.Count;

    var msgBox;
    if( bDisplayUI )
    {
        msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "���������� �������������� �����������������, ������������� � �������: " + count );
    }

    var arrProvidersNames = new Array();
    for( var index = 0; index < count; index++ )
    {
        arrProvidersNames.push( oSupportedProviders.Item(index).Name );
    }

    if( bDisplayUI )
    {
        msgBox.Popup( arrProvidersNames.join("\n"), 0, "������ �������������� �����������������" );
    }

    return arrProvidersNames;
}

//-----------------------------------------------------------------------------

//Getting list of system cryptoproviders
//��������� ������ ��������� �����������������

function GetSystemCryptoProviders( bDisplayUI )
{
    //�������� ��������� ��������� �����������������
    var oCryptoProviders = new ActiveXObject("DigtCrypto.CryptoProviders");

    //������� ��������� ��������� �����������������
    var oSystemProviders= oCryptoProviders.SystemProviders;

    //������� ���������� ��������� ����������������� � ���������
    var count = oSystemProviders.Count;

    var msgBox;
    if( bDisplayUI )
    {
        msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "���������� ��������� �����������������: " + count );
    }

    // ���������� ������������ �����������������, ������������� � �������
    var arrProvidersNames = new Array();
    for( var index = 0; index < count; index++ )
    {
        arrProvidersNames.push( oSystemProviders.Item(index).Name );
    }

    if( bDisplayUI )
    {
        msgBox.Popup( arrProvidersNames.join("\n"), 0, "������ ��������� �����������������" );
    }

    return arrProvidersNames;
}

//-----------------------------------------------------------------------------

