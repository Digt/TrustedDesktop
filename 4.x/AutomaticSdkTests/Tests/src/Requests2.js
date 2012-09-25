
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

// ��� ��������
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

//����� ������� ����� ������ ��� �������� PKI ������� 
//��������, ��������, ���������� ������� � ������������� (GUI) ������.
//Using PKI request creation wizard
//Creating, sending and saving request in interactive mode (GUI).

function CreatePKIRequestWithWizard()
{
    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");
    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");

    //����� ������� ����� ������ ��� �������� PKI �������
    oPKIProfile.CollectData();

    //������������� ����������� ������� ��� �������� PKI �������
    oPKIRequest.PKIProfile = oPKIProfile;

    //��������� ������� �� ��������������� �����������
    var GenResult = oPKIRequest.Generate();

    var msgBox = new ActiveXObject("wscript.shell");
    msgBox.Popup( GenResult );

    //���� � ������� ����������� ���������� ������� � ���� - ���������� ������� � ���� � ������������� ������
    if( oPKIProfile.SaveRequest )
    {
        oPKIRequest.Save( oPKIProfile.RequestFilename );
    }

    //���� � ������� ����������� �������� ������� �� ��. ����� -  ���������� ��������
    if( oPKIProfile.SendRequestByEmail )
    {
        var SendEmailResult = oPKIRequest.SendEmail();       
        msgBox.Popup( SendEmailResult, 0, "�������� �������� ������� �� e-mail" );
    }

    //���� � ������� ����������� �������� ������� � �� -  ���������� ��������
    if( oPKIProfile.SendRequest )
    {
        var SendResult = oPKIRequest.Send();
        msgBox.Popup( SendResult, 0, "�������� �������� ������� � ��" );
    }
}

//-----------------------------------------------------------------------------

//��������, ���������� � �������� PKI ������� � Silent ������ 
//Creating, saving and sending PKI request in "silent" mode

function CreateRequestAndSend()
{
    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");
    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");
    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");

    //��� ���������� ��������������� ����������� ���������� ������������ 3 �����������:
    //������������������ ����������, ���������� ������� �������, ��������� ��� ��������� ssl ����������.
    //� ������ ������� � ���� ���� ������������ ����� �������������� ���� � ��� �� ����������.

    //����������: ������ ���������� ������ ����� ���������� "�������������� �������".
    //��� ������ ����������� ��������� ����� ������ ����������� �� ������� ���������

    //�������� ���������
    oCertStore.Open( CURRENT_USER_STORE, "my" );

    // ����� ����� ��������� ���������, ��������� ��������� � ������������ ��������� ������������
    var oCerts = oCertStore.Display(1);

    //��������� ����������� �� ��������� � ��������� ��� � �������
    //� �������� ������������������� �����������, ����������� �������,
    //����������� ��������� ssl ����������

    var oSSLCertificate = oCerts.Item(0);
    var oOperationCertificate = oCerts.Item(0);
    var oSignCertificate = oCerts.Item(0);

    oPKIProfile.OperationCertificate = oOperationCertificate;
    oPKIProfile.SignCertificate = oSignCertificate;
    oPKIProfile.SignatureCertificatePin = "1";
    //��������� ���� ������� - ��������������� �����������, ���� �� - ��������� ��
    oPKIProfile.RequestType = REQUEST_TYPE_SUSPENDING;
    oPKIProfile.CAType = CA_TYPE_CRYPTO_PRO;
    oPKIProfile.Comment = "���������������";
    oPKIProfile.RevocationReason = REVOCATION_REASON_CERTIFICATE_HOLD;
    oPKIProfile.CertificateHoldDuration = "1-0-0-0-0-0";
    //��������� ���������� ���������� �������
    oPKIProfile.SaveRequest = true;
    oPKIProfile.RequestFilenameFormat = BASE64;
    oPKIProfile.RequestFilename = "Hold.req";
    //��������� ���������� �������� ������� � ��������� ��
    oPKIProfile.SendRequest = true;
    oPKIProfile.ServiceURL = "https://ca-server/ra/ra.wsdl";
    oPKIProfile.SSLCertificate = oSSLCertificate
    //��������� ���������� �������� ������� �� E-Mail
    oPKIProfile.SendRequestByEmail = true;
    oPKIProfile.EmailMessage = "������� ������ �� ���������������";
    oPKIProfile.EmailRecipient = "johndoe@ca.example.com";
    oPKIProfile.EmailSubject = "������ �� ���������������";

    //��������� ������������ ������� ��� �������� �������
    oPKIRequest.PKIProfile = oPKIProfile
    var GenResult = oPKIRequest.Generate();

    var msgBox = new ActiveXObject("wscript.shell");
    msgBox.Popup( GenResult );

    //���� � ������� ����������� ���������� ������� � ���� - ���������� ������� � ���� � ������������� ������
    if( oPKIProfile.SaveRequest )
    {
        oPKIRequest.Save( oPKIProfile.RequestFilename );
    }

    //���� � ������� ����������� �������� ������� �� ��. ����� -  ���������� ��������
    if( oPKIProfile.SendRequestByEmail )
    {
        var SendEmailResult = oPKIRequest.SendEmail();       
        msgBox.Popup( SendEmailResult, 0, "�������� �������� ������� �� e-mail" );
    }

    //���� � ������� ����������� �������� ������� � �� -  ���������� ��������
    if( oPKIProfile.SendRequest )
    {
        var SendResult = oPKIRequest.Send();
        msgBox.Popup( SendResult, 0, "�������� �������� ������� � ��" );
    }
}

//-----------------------------------------------------------------------------

//�������� ������� �� ����� ����������� � silent ������ 
//Creating request for certificate revocation in "silent" mode

function CreateRevocationRequest( strRevokeReqFileName, bDisplayUI )
{
    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");

    //��� �������� ������� �� ����� ����������� ���������� ������������ 2 �����������: 
    //���������� ����������, ���������� ������� �������.
    //� ������ ������� � ���� ���� ������������ ����� �������������� ���� � ��� �� ����������. 

    var oOperationCertificate;
    var oSignCertificate;

    //��� ������ ����������� ��������� ����� ������ ����������� �� ������� ���������

    //�������� ���������
    oCertStore.Open( CURRENT_USER_STORE, "my" );

    var oCerts;
    if( bDisplayUI )
    {
        // ����� ����� ��������� ���������, ��������� ��������� � ������������ ��������� ������������
        oCerts = oCertStore.Display(1);
    }
    else
    {
        oCerts = oCertStore.Store;
    }

    //��������� ����������� �� ��������� � ��������� ��� � ������� � �������� ������������������� �����������, ����������� �������.
    oOperationCertificate = oCerts.Item(0);
    oSignCertificate = oCerts.Item(0);

    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");

    oPKIProfile.OperationCertificate = oOperationCertificate;
    oPKIProfile.SignCertificate = oSignCertificate;
    oPKIProfile.SignatureCertificatePin = "1";

    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");

    //��������� ���� ������� - ��������������� �����������, ���� �� - ��������� ��
    oPKIProfile.RequestType = REQUEST_TYPE_SUSPENDING;
    oPKIProfile.CAType = CA_TYPE_CRYPTO_PRO;
    oPKIProfile.Comment = "�����";
    oPKIProfile.RevocationReason = REVOCATION_REASON_KEY_COMPROMISE;
    //��������� ���������� ���������� �������
    //oPKIProfile.SaveRequest = true;
    oPKIProfile.RequestFilenameFormat = DER;
    oPKIProfile.RequestFilename = strRevokeReqFileName;

    oPKIRequest.PKIProfile = oPKIProfile;

    //��������� �������
    var GenResult = oPKIRequest.Generate();
    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( GenResult );
    }

    //���������� ������� � ���� � ������������� ������
    oPKIRequest.Save( oPKIProfile.RequestFilename, oPKIProfile.RequestFilenameFormat );

    return GenResult;
}

//-----------------------------------------------------------------------------

//�������� ������� �� ������������� ����������� � silent ������ 
//Creating request for resuming certificate

function CreateResumingRequest( strResumeReqFileName, bDisplayUI )
{
    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");
    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");
    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");

    //��� �������� ������� �� ����� ����������� ���������� ������������ 2 �����������: 
    //���������� ����������, ���������� ������� �������.
    //� ������ ������� � ���� ���� ������������ ����� �������������� ���� � ��� �� ����������. 

    //��� ������ ����������� ��������� ����� ������ ����������� �� ������� ���������
    //�������� ���������
    oCertStore.Open( CURRENT_USER_STORE, "my" );

    var oCerts;
    if( bDisplayUI )
    {
        // ����� ����� ��������� ���������, ��������� ��������� � ������������ ��������� ������������
        oCerts = oCertStore.Display(1);
    }
    else
    {
        oCerts = oCertStore.Store;
    }

    //��������� ����������� �� ��������� � ��������� ��� � ������� � �������� ������������������� �����������, ����������� �������.
    var oOperationCertificate = oCerts.Item(0);
    var oSignCertificate = oCerts.Item(0);
    oPKIProfile.OperationCertificate = oOperationCertificate;
    oPKIProfile.SignCertificate = oSignCertificate;
    oPKIProfile.SignatureCertificatePin = "1";

    //��������� ���� ������� - ��������������� �����������, ���� �� - ��������� ��
    oPKIProfile.RequestType = REQUEST_TYPE_RESUMING;
    oPKIProfile.CAType = CA_TYPE_CRYPTO_PRO;
    oPKIProfile.Comment = "�������������";
    //��������� ���������� ���������� �������
    //oPKIProfile.SaveRequest = true;
    oPKIProfile.RequestFilenameFormat = DER;
    oPKIProfile.RequestFilename = strResumeReqFileName;

    oPKIRequest.PKIProfile = oPKIProfile;

    //��������� �������
    var GenResult = oPKIRequest.Generate();
    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( GenResult );
    }

    //���������� ������� � ���� � ������������� ������
    oPKIRequest.Save( oPKIProfile.RequestFilename, oPKIProfile.RequestFilenameFormat );

    return GenResult;
}

//-----------------------------------------------------------------------------

//�������� ������� �� ���������� ����� ����������� � silent ������ 
//Creating request for renewing certificate in "silent" mode

function CreateRenewingRequest( strRenewReqFileName, bDisplayUI )
{
    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");
    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");
    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");

    //��� ���������� ���������� ����� ����������� � ��������� �� ���������� ������������ 2 �����������: 
    //����������, ����� �������� �����������, ���������� ������� �������. 
    //� ������ ������� � ���� ���� ������������ ����� �������������� ���� � ��� �� ����������. 
    //��� ������ ����������� ��������� ����� ������ ����������� �� ������� ���������

    //�������� ���������
    oCertStore.Open( CURRENT_USER_STORE, "my" );

    var oCerts;
    if( bDisplayUI )
    {
        // ����� ����� ��������� ���������, ��������� ��������� � ������������ ��������� ������������
        oCerts = oCertStore.Display(1);
    }
    else
    {
        oCerts = oCertStore.Store;
    }

    //��������� ����������� �� ��������� � ��������� ��� � ������� � �������� ������������������� �����������, ����������� �������, ����������� ��������� ssl ����������
    var oOperationCertificate = oCerts.Item(0);
    var oSignCertificate = oCerts.Item(0);
    oPKIProfile.OperationCertificate = oOperationCertificate;

    //���������� ������� ������� �� ����������
    //��������� ������ �� ���������� �����������
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.CRequestTemplate" );
    oRequestTemplate.Certificate = oOperationCertificate;
    oRequestTemplate.ExtendedKeyUsage = oOperationCertificate.EKU;
    oRequestTemplate.CryptoProvider = oOperationCertificate.ProviderName;
    oRequestTemplate.KeyUsage = oOperationCertificate.KU;
    oRequestTemplate.CreateNewKeySet = true;
    oRequestTemplate.MarkKeysExportable = true;
    //���������� �������� ���������� ��� ����������
    //oRequestTemplate.Keyset = "NewContainer";
    oRequestTemplate.Keyset = GenerateContainerName();

    //��������� ������ ������� �� ���������� � ������� ��� ������������ PKI �������
    oPKIProfile.RequestTemplate = oRequestTemplate;
    oPKIProfile.SignCertificate = oSignCertificate;
    oPKIProfile.SignatureCertificatePin = "1";

    //��������� ���� ������� - ��������������� �����������, ���� �� - ��������� ��
    oPKIProfile.RequestType = REQUEST_TYPE_UPDATE_KEY;
    oPKIProfile.CAType = CA_TYPE_CRYPTO_PRO;

    //��������� ���������� ���������� �������
    oPKIProfile.SaveRequest = true;
    oPKIProfile.RequestFilenameFormat = DER;
    oPKIProfile.RequestFilename = strRenewReqFileName;

    //��������� ������� �������� PKI �������
    oPKIRequest.PKIProfile = oPKIProfile;

    //��������� �������
    var GenResult = oPKIRequest.Generate();
    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( GenResult );
    }

    //���������� ������� � ���� � ������������� ������
    oPKIRequest.Save( oPKIProfile.RequestFilename, oPKIProfile.RequestFilenameFormat );

    return GenResult;
}

//-----------------------------------------------------------------------------

//�������� ������� �� ���������� � silent ������ 
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

    //���������� ������� ������� �� ����������
    //��������� ������ �� ���������� �����������
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
    //���������� �������� ���������� ��� ���������� ������� ��� ����� �� �������� ���, ����� ��� ����� ������������� �������������
    //oRequestTemplate.Keyset = "NewContainer";
    oRequestTemplate.Keyset = GenerateContainerName();

    //��������� ������ ������� �� ���������� � ������� ��� ������������ PKI �������
    oPKIProfile.RequestTemplate = oRequestTemplate;

    //��������� ���� ������� - ��������������� �����������, ���� �� - ��������� ��
    oPKIProfile.RequestType = REQUEST_TYPE_CERT_PKCS10;
    oPKIProfile.CAType = CA_TYPE_CRYPTO_PRO;

    //��������� ���������� ���������� �������
    oPKIProfile.SaveRequest = true;
    oPKIProfile.RequestFilenameFormat = DER;
    oPKIProfile.RequestFilename = resultFile;

    //��������� ������� �������� PKI �������
    oPKIRequest.PKIProfile = oPKIProfile;

    //��������� �������
    var GenResult = oPKIRequest.Generate();
//    var msgBox = new ActiveXObject("wscript.shell");
//    msgBox.Popup( GenResult );

    //���������� ������� � ���� � ������������� ������
    if( !GenResult )
    {
        return null;
    }

    oPKIRequest.Save( oPKIProfile.RequestFilename, oPKIProfile.RequestFilenameFormat );
    return oPKIRequest;
}

//-----------------------------------------------------------------------------

//�������� ���������������� ����������� � silent ������
//Creating self signed certificate in "silent" mode

function CreateSelfSignedCertificate()
{
    var oPKIProfile = new ActiveXObject("DigtCrypto.PKIProfile");
    var oPKIRequest = new ActiveXObject("DigtCrypto.PKIRequest");

    //���������� ������� ������� �� ����������
    //��������� ������ �� ���������� �����������
    var oRequestTemplate = new ActiveXObject( "DigtCrypto.CRequestTemplate" );
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
    oRequestTemplate.ShowSendRequestWindow = true;
    oRequestTemplate.MarkKeysExportable = true;
    oRequestTemplate.KeyLength = 512;

    //���������� �������� ���������� ��� ���������� ������� ��� ����� �� �������� ���, ����� ��� ����� ������������� �������������
    //oRequestTemplate.Keyset = "NewContainer"
    oRequestTemplate.Keyset = GenerateContainerName();

    //��������� ������ ������� �� ���������� � ������� ��� ������������ PKI �������
    oPKIProfile.RequestTemplate = oRequestTemplate;

    //��������� ���� ������� - ��������������� �����������, ���� �� - ��������� ��
    oPKIProfile.RequestType = REQUEST_TYPE_SELF_SIGNED_CERT;

    //��������� ���������� ����������
    oPKIProfile.SaveRequest = true;
    oPKIProfile.RequestFilenameFormat = DER;
    oPKIProfile.RequestFilename = "selfsigned_cert.cer";

    //��������� ������� �������� PKI �������
    oPKIRequest.PKIProfile = oPKIProfile;

    //��������� �������
    var GenResult = oPKIRequest.Generate();
    var msgBox = new ActiveXObject("wscript.shell");
    msgBox.Popup( GenResult );

//    //���������� �����������
//    oPKIRequest.Save( oPKIProfile.RequestFilename, oPKIProfile.RequestFilenameFormat );
}

//-----------------------------------------------------------------------------

