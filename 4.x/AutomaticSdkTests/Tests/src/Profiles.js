
var CERT_FOR_ENCRYPT = 0;
var CERT_FOR_DECRYPT = 1;
var CERT_FOR_SIGN = 2;
var BASE64 = 0;
var DER = 0;
var REGISTRY_STORE = 0;
var CURRENT_USER_STORE = 1;
var LOCAL_MACHINE_STORE = 0;
// enum CERTIFICATE_VERIFY_LEVEL
var CERTIFICATE_VERIFY_LOCAL_CRL = 1;
var CERTIFICATE_VERIFY_ONLINE_CRL = 2;
//var CERTIFICATE_VERIFY_OCSP = 3 // reserved
//enum SILENT_LEVEL
var SILENT_LEVEL_SERVER = 1;
var SILENT_LEVEL_REQUIRED = 2;
var SILENT_LEVEL_WINDOWS = 3;
var SILENT_LEVEL_INTERACTIVE = 4;
// CERTIFICATE_PUPROSE
var CERTIFICATE_FOR_NEW_SIGNATURE = 1;
var CERTIFICATE_FOR_SIGNATURE_VERIFYING = 2;
var CERTIFICATE_FOR_ENCIPHER = 3;
var CERTIFICATE_FOR_DECIPHER = 4;
// CERTIFICATE_WORKABILITY
var CERTIFICATE_ALLOWED = 1;
var CERTIFICATE_RESTRICTED = 2;
var CERTIFICATE_UNKNOWN = 3;

//-----------------------------------------------------------------------------

//Creating of a new profile
//�������� ������ ������� 

function CreateNewProfile( strCertFileName, bDisplayUI )
{
    var CERTIFICATE_PATH = strCertFileName;

    //������������ �������� ���������� �������, ��������, ����� �������� � ���������

    var objProfile = new ActiveXObject("DigtCrypto.Profile");
    //����� ��������� ��������� �������
    objProfile.Provider = "Provider";
    objProfile.HashAlg = "Hashalg";
    objProfile.EncAlg = "EncAlg";
    objProfile.SignFinalWindow = false;
    objProfile.SignStartWindow = false;
    objProfile.SignFormatWindow = false;
    objProfile.EncryptFinalWindow = true;
    objProfile.EncryptStartWindow = true;
    objProfile.EncryptFormatWindow = true;
    objProfile.SignIncludeBase64Headers = true;
    objProfile.EncryptIncludeBase64Headers  = true;
    objProfile.SignExitFormat = DER;
    objProfile.EncryptExitFormat = DER;
    objProfile.UseCertificateForEncrypt = true;
    objProfile.Detach = true;
    objProfile.DefaultPath = "def/path";
    objProfile.IncludeSignatureTime = true;
    objProfile.ShowStatusReport = true;
    objProfile.EncryptToSenderAddress = true;
    objProfile.Resource = "Resource";
    objProfile.SignPropertiesWindow = true;
    objProfile.SignCertificateWindow = true;
    objProfile.EncryptPropertiesWindow = true;
    objProfile.EncryptRecipientsWindow = true;
    objProfile.Name = "Test profile";
    objProfile.Comment = "Comment";
    objProfile.Description = "Remove this test profile";
    objProfile.SilentLevel = SILENT_LEVEL_SERVER;

    //��������� �������� ������������� ������������

    //�������� ����� OID
    var oOID = new ActiveXObject("DigtCrypto.OID");
    oOID.FriendlyName = "�������������� ������������";
    oOID.Value = "1.3.6.1.5.5.7.3.2";

    var msgBox;
    if( bDisplayUI )
    {
        msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "FriendlyName = " + String(oOID.FriendlyName) + "\n" + "Value = " + String(oOID.Value), 0, "������������� ��������" );
    }

    //������� ���� OID � ��������� OIDs
    var oOIDs = new ActiveXObject("DigtCrypto.OIDs");
    oOIDs.Add( oOID )

    //�������� �������� SignatureRequiredOIDs
    var oPolicyProfile = new ActiveXObject("DigtCrypto.PolicyProfile");
    oPolicyProfile.SignatureRequiredOIDs = oOIDs;
    oOIDs = null;

    //�������� �������� ������������� ����������� � �������
    objProfile.Policy = oPolicyProfile;

    var oCertificate = new ActiveXObject("DigtCrypto.Certificate");
    var oRecipients = new ActiveXObject("DigtCrypto.Certificates");
    var oCertificateStore = new ActiveXObject("DigtCrypto.CertificateStore");


    //��������� ����������� ������� � ������� (����������� ���������� � �������������
    //      ��������������� ����� �� �������, � ���������������� �����������)
    oCertificate.Load( CERTIFICATE_PATH );
    //��������, ������������� �� ���������� ��������
    var lWorkability = -1
    lWorkability = objProfile.CheckCertificate(oCertificate, CERTIFICATE_FOR_NEW_SIGNATURE);
    if( lWorkability == 1 )
    {
        if( bDisplayUI )
        {
            msgBox.Popup( oCertificate.SerialNumber );
        }
        objProfile.SetCertificate( CERT_FOR_SIGN, "", oCertificate );
        oRecipients.Add( oCertificate ); // �������� ��������� ������������ �����������
    }
    else if( bDisplayUI )
    {
        msgBox.Popup( "���������� �� �������� �� �������� �������������" );
    }

    //��������� ��������� ������������ ����������� � �������
    objProfile.Recipients = oRecipients;

    //��������� ��������� ������������ ��� ��������� �������
    objProfile.SetVerifiedCertificates( CERTIFICATE_VERIFY_ONLINE_CRL , oRecipients );

    //�������� ������������� ������ ������ �������
    var sProfileID = objProfile.ID;

    //�������� ��������� ������� � ���������
    var objProfiles = new ActiveXObject("DigtCrypto.Profiles");
    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");

    objProfileStore.Open(REGISTRY_STORE); // ��������� ��������� �������� � ������� ����������
    var objProfiles = objProfileStore.Store; //  �������� ��������� �������� �� ���������
    objProfiles.Add(objProfile); //  ��������� ������� � ��������� ��������
    objProfileStore.Store = objProfiles; // ������������� ��������� �������� � ���������
    objProfileStore.Save();  // ��������� ��������� ��������
    objProfileStore.Close(); // ��������� ���������

    objProfile = null;
    objProfileStore = null;
    objProfiles = null;

    return sProfileID;
}

//-----------------------------------------------------------------------------

//Getting profile parameters
//��������� ���������� �������

function GetProfileParameters( strProfileid, bDisplayUI )
{
    var objProfiles = new ActiveXObject("DigtCrypto.Profiles");
    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open(REGISTRY_STORE);
    var objProfiles = objProfileStore.Store;

    var objProfile;
    if( typeof(strProfileid) == "undefined" || String(strProfileid).length == 0)
    {
        //������� ������� �� ���������
        objProfile =  objProfiles.DefaultProfile;
    }
    else
    {
        //������� ������� �� ��������������
        objProfile =  objProfiles.Profile( strProfileid );
    }

    //������� ��������� ��������� �������
    var strProfileProps = "";
    strProfileProps += "ID=" + objProfile.ID + "\n";
    strProfileProps += "SignStartWindow="  + String(objProfile.SignStartWindow) + "\n";
    strProfileProps += "Name="  + objProfile.Name + "\n";
    strProfileProps += "Description="  + objProfile.Description + "\n";
    strProfileProps += "Comment="  + objProfile.Comment + "\n";
    strProfileProps += "SilentLevel="  + String(objProfile.SilentLevel) + "\n"    ;

    //��������� ����������� ���������� �������� �� ������� ����������� �������
    var objCertificate = objProfile.GetCertificate(CERT_FOR_SIGN);
    strProfileProps += "Signature certificate=" + objCertificate.SerialNumber + "\n";
    //��������� ��� ���� ����������� �������
    objProfile.GetPin(CERT_FOR_SIGN);
    strProfileProps += "Signature certificate PIN=" + objProfile.GetPin(CERT_FOR_SIGN) + "\n";

    //��������� ��������� ������������ �� ������� ����������� ������������ ���������
    objCertificates = objProfile.Recipients;
    for( var i = 0; i < objCertificates.Count; i++ )  //����� ������������ �����������
    {
        var oCertificate = objCertificates.Item(i);
        strProfileProps += "Recipient certificate " + String(i+1) + "=" + oCertificate.SerialNumber + "\n";
    }

    if( bDisplayUI )
    {
        msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( strProfileProps );
    }

    objProfile = null;
    objProfileStore = null;
    objProfiles = null;

    return strProfileProps;
}

//-----------------------------------------------------------------------------

//Removing profile
//�������� �������

function RemovingProfile( strProfileId )
{
    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");

    objProfileStore.Open(REGISTRY_STORE);

    //�������� �������� ��������
    var objProfiles = objProfileStore.Store;

    //������ ������ ���������� �������
    var index = -1;
    for( var i = 0; i < objProfiles.Count; i++ )
    {
        if( objProfiles.Item(i).ID == strProfileId )
        {
            index = i;
            break;
        }
    }

    //�������� ����������: ���� ������� ������� �� ������
    if( index == -1 )
    {
        return false;
    }

    objProfiles.Remove( index );
    objProfileStore.Store = objProfiles;
    objProfileStore.Save();
    objProfileStore.Close();
    objProfiles = null;

    return true;
}

//-----------------------------------------------------------------------------

//Displaying of a profile manager
//������ ������ ��������� ��������

function DisplayProfileManager()
{
    var oProfileStore = new ActiveXObject( "DigtCrypto.ProfileStore" );
    oProfileStore.Open( REGISTRY_STORE );

    oProfileStore.Display();
}

//-----------------------------------------------------------------------------

//Coping a profile
//�������� ����� �������

function CopyProfile( strProfileName, strProfileDescription, bDisplayUI )
{
    var objProfile = new ActiveXObject("DigtCrypto.Profile");
    objProfile.Name = strProfileName;
    objProfile.Description = strProfileDescription;

    var objProfileClone = objProfile.Clone();
    if( bDisplayUI )
    {
        var strResult = "��� ��������� �������:  " + objProfile.Name;
        strResult += "\n�������� ��������� �������:  " + objProfile.Description;
        strResult += "\n\n��� �������������� �������:  " + objProfileClone.Name;
        strResult += "\n�������� �������������� �������:  " + objProfileClone.Description;

        msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( strResult );
    }

    return objProfileClone;
}

//-----------------------------------------------------------------------------

//Loading profile store from web-service
//�������� ��������� �������� c ���-�������

/*

//���������� ����������

Const LOG_LEVEL_DISABLED = 0 //����������� ��������� 
Const LOG_LEVEL_ERROR = 1 //������ 
Const LOG_LEVEL_WARN = 2 //�������������� 
Const LOG_LEVEL_INFO = 3 //���������� 
Const LOG_LEVEL_DEBUG = 4 //������� 
Const LOG_LEVEL_TRACE = 5 //����������� 

Const strPattern = "log4cplus.appender.toDSL.layout.ConversionPattern=%d [%5t] [%p] [%c] [%x] - %m%n"

Dim oUtil , sLoggerID, strLoggerStatement
Set oUtil = new ActiveXObject("DigtCrypto.Util")


sLoggerID = oUtil.StartupLogger(strPattern, LOG_LEVEL_TRACE, 500*1024, "DSS" , 0)
MsgBox sLoggerID, , "sLoggerID"

strLoggerStatement = oUtil.GetLoggerStatements(sLoggerID )
MsgBox strLoggerStatement
oUtil.ShutdownLogger sLoggerID 


*/

//-----------------------------------------------------------------------------
