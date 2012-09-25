

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
//������, ������� � ���������� �����������

function ImportExportAndSaving(certPath, certSavingPath, bDisplayUI)
{
    if( typeof(bDisplayUI) == "undefined" )
    {
        bDisplayUI = true;
    }

    var oCertificate1 = new ActiveXObject( "DigtCrypto.Certificate" );

    // ��������� ���������� �� �����
    oCertificate1.Load( certPath );

    // ������������ ���������� � DER �������
    var szData_Der = oCertificate1.Export(DER_TYPE);

    // �������� ����� ��������� ������� �����������
    if( bDisplayUI )
    {
        oCertificate1.Display();
    }
    oCertificate1 = null;

    // ����������� ������ � ������������ � ����� ������ �����������
    var oCertificate2 = new ActiveXObject( "DigtCrypto.Certificate" );
    oCertificate2.Import( szData_Der );

    // ��������� ����������
    oCertificate2.Save( certSavingPath, BASE64_TYPE, 0 );

    return oCertificate2;
}

//-----------------------------------------------------------------------------

//Getting properties of a certificate
//��������� ������� �����������

function GettingCertProperties( strCertificateFileName, bDisplayUI )
{
    if( typeof(bDisplayUI) == "undefined" )
    {
        bDisplayUI = true;
    }

    var oCert;

    //���� ���� ��� ����������� �� �����
    if( (typeof(strCertificateFileName)!="string") || (strCertificateFileName == "") )
    {
        var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");
        var oCerts;

        if( bDisplayUI )
        {
            // ����� ����� ��������� ���������
            oCerts = oCertStore.Display(SYSTEM_STORE_MY || SYSTEM_STORE_ADDRESS_BOOK);
        }
        else
        {
            // ������� ��������� ������ ������������ ("my") �������� ������������
            oCertStore.Open( CURRENT_USER_STORE, "my" );
            oCerts = oCertStore.Store;
        }

        // ��������, �������� �� ���������� �� ���������
        if( oCerts.Item.Count == 0 )
        {
            if( bDisplayUI )
            {
                return "�� ������ ����������� �� ���� �������";
            }
            else
            {
                return "��������� �� �������� �� ������ �����������"
            }
        }

        // ���������� ������� ����������� �� ����������� ������ (������ ����� �� ������������!)
        oCert = oCerts.Item(0);
    }
    else
    {
        //����� ������� ���������� �� �����
        oCert = new ActiveXObject("DigtCrypto.Certificate");
        oCert.Load( strCertificateFileName );
    }

    // ���������� ����������
    if( bDisplayUI )
    {
        oCert.Display();
    }

    // ������� �������� ������������
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
        sMess += "ProviderName: " + String(oCert.ProviderName) + "\n"; // �������� ��������� ������ ��� ������������, ������� �������� � ����������
        sMess += "ContainerName: " + String(oCert.ContainerName) + "\n"; // �������� ��������� ������ ��� ������������, ������� �������� � ����������
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
        msgBox.Popup( sMess, 0, "�������� �����������" );
    }

    // ������� �������������� ����������
    oCert = null;
    oCerts = null;
    oCertStore = null;

    return sMess;
}

//-----------------------------------------------------------------------------

//Setting certificate properties
//Setting of a CryptoAPI context of a certificate
//��������� ������� ����������� 
//������ ��������� CryptoAPI-��������� ����������� 

function SettingCryptoApiContext( sCertPath, bDisplayUI )
{
    if( typeof(bDisplayUI) == "undefined" )
    {
        bDisplayUI = true;
    }

    var oCert1 = new ActiveXObject("DigtCrypto.Certificate");

    //�������� ���������� �� �������� ������� ��������
    oCert1.Load( sCertPath );

    //��������� �������� CertContext
    var sCertContext = oCert1.CertContext;

    //�������� ����� ������ Certificate
    var oCert2 = new ActiveXObject("DigtCrypto.Certificate");

    //��������� ��� �������� �����������
    oCert2.CertContext = sCertContext;

    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( String(oCert2.CertContext), 0, "�������� �����������" );
    }

    return oCert2;
}

//-----------------------------------------------------------------------------

//Setting certificate properties
//Setting of a serial number and issuer of a certificate
//��������� ������� ����������� 
//������ ��������� ��������� ������ � �������� �����������

function SettingSerialAndIssuerToCert( sSerialNumber, sIssuer )
{
    //�������� ����� ������ Certificate
    var oCert = new ActiveXObject("DigtCrypto.Certificate");

    //��������� ��� ����� �������� �����������
    oCert.SerialNumber = sSerialNumber;
    oCert.IssuerName = sIssuer;

    return oCert;
}

//-----------------------------------------------------------------------------

//Certificate comparsion
//��������� ������������

function CompareCertificates( strCertificateFileName, bDisplayUI )
{
    if( typeof(bDisplayUI) == "undefined" )
    {
        bDisplayUI = true;
    }

    //������� ����� ���������� �� ��������� ������ �����������
    //��������� ��������� ��� ������ �����������
    var oCertStore = new ActiveXObject("DigtCrypto.CertificateStore");
    oCertStore.Open( CURRENT_USER_STORE, "my" );

    // ����� ����� ��������� ��������� � ��������� ��������� �����������
    var oCerts
    if( bDisplayUI )
    {
        oCerts = oCertStore.Display(1);
    }
    else
    {
        oCerts = oCertStore.Store;
    }

    //��������� ����������� �� ���������Set
    var oCert = oCerts.Item(0);


    //������� ������ ���������� ��������� �� �����
    // �������� ���������� ��� ���������
    var oCert1 = new ActiveXObject("DigtCrypto.Certificate");
    oCert1.Load( strCertificateFileName );

    //���������� ���
    if( bDisplayUI )
    {
        oCert1.Display();
    }

    //��������� ������������
    var bResult = oCert.isEqual(oCert1);
    var msgBox = new ActiveXObject("wscript.shell");

    if( bDisplayUI )
    {
        if( bResult == true )
        {
            msgBox.Popup( "����������� �����" );
        }
        else
        {
            msgBox.Popup( "����������� �� �����" );
        }
    }

    oCertStore.Close();

    return bResult;
}

//-----------------------------------------------------------------------------

//Verifying certificate status
//�������� ������� �����������

function VerifyCertificateStatus( bDisplayUI )
{
    if( typeof(bDisplayUI) == "undefined" )
    {
        bDisplayUI = true;
    }

    var POLICY_TYPE_NONE = 0 //��� �������� ������������� ������������

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

    // ��������� ���������
    oCertStore.Open( CURRENT_USER_STORE, "my" );

    // ����� ����� ��������� ��������� � ��������� �����������
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
    status = oCert.IsValid(POLICY_TYPE_NONE); //�������� ������ ����������� 
    var strStatus = "Status: ";
    switch( status )
    {
        case VS_CORRECT:
            strStatus += "���������";
            break;
        case VS_UNSUFFICIENT_INFO:
            strStatus += "������ ����������";
            break;
        case VS_UNCORRECT:
            strStatus += "�����������";
            break;
        case VS_INVALID_CERTIFICATE_BLOB:
            strStatus += "���������������� ���� �����������";
            break;
        case VS_CERTIFICATE_TIME_EXPIRIED:
            strStatus += "����� �������� ����������� ������� ��� ��� �� ���������";
            break;
        case VS_CERTIFICATE_NO_CHAIN:
            strStatus += "���������� ��������� ������� ������������";
            break;
        case VS_CERTIFICATE_CRL_UPDATING_ERROR:
            strStatus += "������ ���������� �����������";
            break;
        case VS_LOCAL_CRL_NOT_FOUND:
            strStatus += "�� ������ ��������� ���";
            break;
        case VS_CRL_TIME_EXPIRIED:
            strStatus += "������� ����� �������� ���";
            break;
        case VS_CERTIFICATE_IN_CRL:
            strStatus += "��������� ��������� � ���";
            break;
        case VS_CERTIFICATE_IN_LOCAL_CRL:
            strStatus += "��������� ��������� � ��������� ���";
            break;
        case VS_CERTIFICATE_CORRECT_BY_LOCAL_CRL:
            strStatus += "���������� ������������ �� ���������� ���";
            break;
        case VS_CERTIFICATE_USING_RESTRICTED:
            strStatus += "������������� ����������� ����������";
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
