
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

//�������� ������� �� ���������� � ������������� ������
//Creating certificate request in interactive mode

function CreateCertificateRequestInteractive()
{
    var oRequestCertificate = new ActiveXObject( "DigtCrypto.Request" );

    //����� ������� �������� ������� �� ����������
    oRequestCertificate.Display();

    oRequestCertificate = null;
}

//-----------------------------------------------------------------------------

//�������� ������� � �������������� �����
//Sending request into certification authority (CA)

function SendRequestIntoCA()
{
    //const CA_ADDRESS = "172.17.2.72"
    var CA_ADDRESS = "localhost";
    //var CA_ADDRESS = "http://www.cryptopro.ru/certsrv/";

    //�������� ������ ������� �� ����������
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.RequestTemplate" );
    oRequestTemplate.CN = "John Doe";
    oRequestTemplate.C = "RU";
    oRequestTemplate.E = "johndoe@example.com";
    oRequestTemplate.L = "�����";
    oRequestTemplate.O = "��������";
    oRequestTemplate.OU = "�����";
    oRequestTemplate.S = "������";
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

    //��������� ������ ������� �� ����������
    var oRequest = new ActiveXObject("DigtCrypto.Request");
    oRequest.Template = oRequestTemplate;

    //����������� ������
    oRequest.Generate();

    oRequestTemplate = null;

    //������� ��������� �������
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
    msgBox.Popup( sResult, 0, "��������� ������� ������� �� ����������" );

    //�������� ���������� ��������� ��, ��� ����� �������������� ������ 
    var oIssuerCertificate = new ActiveXObject("DigtCrypto.Certificate");
    oIssuerCertificate.Load( "root.cer" );

    //�������� ������ �� ��������� � ��
    sNewCert = oRequest.Send(CA_ADDRESS, oIssuerCertificate, 1);

    //�������� ��� ������� ��������� �������. ���� ������ �� ����� 5, �� ��������� ������ ������� 
    if( oRequest.CADisposition != 5)
    {
        var oCert = new ActiveXObject("DigtCrypto.Certificate");
        oCert.Import( String(sNewCert) );
        oCert.Display();
    }
    else
    {
        msgBox.Popup( sResult, 0, "�� �������� �� ���������� ��������� �������." );
    }
}

//-----------------------------------------------------------------------------

//�������� ������� �� ���������� � �������������� �������
//������ � ����� template.xml
//Request template creating with template
//Template stored in file template.xml

function CreateRequestWithTemplate( strRequestFileName, strResultReqFile, bDisplayUI )
{
    var oRequestCertificate = new ActiveXObject( "DigtCrypto.Request" );
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.RequestTemplate" );

    //�������� ����� ������� ������� �� ���������� �� XML �����
    oRequestTemplate.LoadXMLTemplate(strRequestFileName);
    
    //Perceptible container name for automatic tests (and manual container removing :] )
    oRequestTemplate.Keyset = GenerateContainerName();

    //����� ������� �������� ������� �� ����������
    oRequestCertificate.Template = oRequestTemplate;

    //����� ������� ��������� ������� �� ����������.
    //��������� ����������������� ���������� � ��������� ����� ����������� ������� �� �������.
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

//�������� ������� �������
//����� Retry (��������� ������� ��������� �� �������������� �������)
//Request status checkup
//Method Retry (checking status of processing by request identifier)

function CheckRequestStatusByID()
{
    var ROOT_ISSUER = "CN=�� 2001,OU=Unit,O=Digt,L=Yoshkar-Ola,S=����� ��,C=RU,E=ca@mail.ru,";
    var CA_ADDRESS = "172.17.2.72";

    //������� ������ ������� �� ����������
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

    //��������� ������ ������� �� ����������
    var oRequest = new ActiveXObject("DigtCrypto.Request");
    oRequest.Template = oRequestTemplate;
    //����������� ������
    oRequest.Generate();

    //�������� ���������� ��������� ��     
    var oIssuerCertificate = new ActiveXObject("DigtCrypto.Certificate");
    oIssuerCertificate.Load( "root.cer" );

    //�������� ������ �� ��������� � �� (�� �������� �� ���������� ��������� �������)
    //� �������� ������ ������� ������������� �������
    sNewCert = oRequest.Send(CA_ADDRESS, oIssuerCertificate, 1);

    //����� ������ ���� ����������� ������ ��������� ������� � ��
    //MsgBox "������������� �������: "+String(sNewCert)+ "��������� ���������� �� ������� ������� � ��"
    var msgBox = new ActiveXObject("wscript.shell");
    msgBox.Popup( "������������� �������: " + String(sNewCert) + "��������� ���������� �� ������� ������� � ��" );

    //�������� ������ ������� � ������� ���������� ����������
    var sNewCert1 = oRequest.Retry(sNewCert, CA_ADDRESS, oIssuerCertificate, 1);

    var oCert = new ActiveXObject("DigtCrypto.Certificate") ;
    oCert.Import( String(sNewCert1) );
    oCert.Display();
}

//-----------------------------------------------------------------------------


//�������� ������� �������
//����� RetrievePending (��������� ������� ��������� �� �������)
//Request status checkup
//Method RetrievePending (checking status of processing by request)

function CheckRequestStatusByRequest()
{
    //�������� ���������� � ��������

    var CURRENT_USER_STORE = 1;

    //������� ��������� �������� � ������� ������ ������, ������ �������� ����� ���������
    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");
    oCertStore.Open( CURRENT_USER_STORE, "request" );
    // ����� ����� ��������� ���������
    var oCerts = oCertStore.Display(32, 0, 0);
    var oCertTemplate = oCerts.Item(0);

    //�������� ������ ��������� ������� � ��
    var oRequest = new ActiveXObject("DigtCrypto.Request");
    var sNewCert1 = oRequest.RetrievePending(oCertTemplate, 1);

    //���� ������ ���������, �� ���������� ���������� ����������
    if( oRequest.CADisposition == 3 )
    {
        var oCert = new ActiveXObject("DigtCrypto.Certificate");
        oCert.Import( String(sNewCert1) );
        oCert.Display();
    }
    else
    {
        //MsgBox "������ �� ���������"
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "������ �� ���������" );
    }
}

//-----------------------------------------------------------------------------

//�������, ������ � ���������� ������� �� ���������� 
//Export, import and saving certificate request

function RequestExportAndImport(reqFileName)
{
    var DER_TYPE = 1;

    //������� ������ ������� �� ����������
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.RequestTemplate" );

    //�������� ������ ������� �� ����������
    oRequestTemplate.CN = "John Doe";
    oRequestTemplate.C = "RU";
    oRequestTemplate.E = "johndoe@example.com";
    oRequestTemplate.ExtendedKeyUsage = "<keyPurposeId>1.3.6.1.5.5.7.3.1</keyPurposeId>";
    oRequestTemplate.CryptoProvider = "Microsoft Base Cryptographic Provider v1.0";
    oRequestTemplate.CryptoProviderType = 1;
    oRequestTemplate.KeyUsage = 1;
    oRequestTemplate.CreateNewKeySet = true;
    oRequestTemplate.Keyset = GenerateContainerName();

    //��������� ������ ������� �� ����������
    var oRequest = new ActiveXObject("DigtCrypto.Request");
    oRequest.Template = oRequestTemplate;

    //��������� � ��������� �������"
    oRequest.Generate();

    //������������ ������ �� ���������� � ������
    szData = oRequest.Export(DER_TYPE);

    oRequest = null;

    //����������� ������ �� ���������� �� ������
    var oRequestImport = new ActiveXObject( "DigtCrypto.Request" );
    oRequestImport.Import( szData );

    //�������� ������ �� ���������� � ������� DER
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

//�������� ���������������� �����������
//Creating self signed certificate

function CreateSelfSignedCert()
{
    //var oCert = new ActiveXObject("DigtCrypto.Certificate");

    //������� ������ ������� �� ����������
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

    //��������� ������ ������� �� ����������
    var oRequest = new ActiveXObject("DigtCrypto.Request");
    oRequest.Template = oRequestTemplate;

    //����������� ��������������� ���������� 
    var oCert = oRequest.CreateSelfSignedCertificate();
    //var oCert = oRequest.GenerateSelfSigned();

    oCert.Display();
}

//-----------------------------------------------------------------------------

