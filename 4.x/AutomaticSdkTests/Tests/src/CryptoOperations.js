
//enum FORMAT
var BASE64_TYPE = 0;
var DER_TYPE = 1;

//enim PROFILESTORETYPE (�������)
var REGISTRY_STORE = 0;

//enum DATATYPE (��� ������)
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
//�������� �������

function SignatureCreation( sPlainFileName, sSignFileName, bDisplayUI )
{
    // � ��������� ������ ���� ������� � ������������ ����������� ���

    var PLAIN_DATA_FILE = sPlainFileName;
    var OUTPUT_DATA_FILE = sSignFileName;

    //������� ������� �� ��������� ��� �������� �����, ���� ��� ���

    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open( REGISTRY_STORE ); //��������� ��������� ��������
    var objProfiles = objProfileStore.Store ;//�������� ��������� ��������

    var objProfile;
    if( objProfiles.Count > 0 )
    {
        objProfile = objProfiles.DefaultProfile; //������� ������� �� ���������
    }
    else
    {
        objProfile = new ActiveXObject("DigtCrypto.Profile"); //�������� ����� �������
    }

    //��������� � ��������� �������, ��������� ������, ���������� �� �������
    if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
    {
        objProfile.CollectData( SIGN_WIZARD_TYPE ); //�������� ������ ������� ��� ����� ������
    }

    var CheckResult = objProfile.CheckData( SIGN_WIZARD_TYPE ); //��������, ��� �� ������ �������
    if( ALL_OK == CheckResult )
    {
        var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
        oPKCS7Message.Profile = objProfile; //��������� ������� � �����������
        oPKCS7Message.Load( DT_PLAIN_DATA, String(PLAIN_DATA_FILE), "" ); //�������� �������� ������
        var signResult = oPKCS7Message.Sign();     //�������� ������, ��������� ��������� ������� �� �������
        oPKCS7Message.Save( DT_SIGNED_DATA, BASE64_TYPE, OUTPUT_DATA_FILE ); //�������� ������ 
        oPKCS7Message = null;

        return true;
    }
    else if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "������� ����������� ��������:\n" + CheckResult );
    }

    return false;
}

//-----------------------------------------------------------------------------

//Adding sugnature
//���������� ������� 

function AddSignature( sSignFile, sResultSignFile, bDisplayUI )
{
    var SIGN_DATA_FILE = sSignFile;
    var OUTPUT_DATA_FILE = sResultSignFile;

    //������� ������� �� ��������� ��� �������� �����, ���� ��� ���
    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open( REGISTRY_STORE ); //��������� ��������� ��������
    var objProfiles = objProfileStore.Store; //�������� ��������� ��������

    var objProfile;

    if( objProfiles.Count > 0 )
    {
        objProfile = objProfiles.DefaultProfile; //������� ������� �� ���������
    }
    else 
    {
        objProfile = new ActiveXObject("DigtCrypto.Profile"); //�������� ����� �������
    }

    //�������� ���� �������������� �������
    var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
    oPKCS7Message.Load( DT_SIGNED_DATA, String(SIGN_DATA_FILE), "" );

    //������� ��������� ��������
    var oSigners = oPKCS7Message.Signers;
    if( oSigners.Count > 0 )
    {
        //��������� ������� � �����������
        oPKCS7Message.Profile = objProfile;

        if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
        {
            objProfile.CollectData( ADD_SIGN_WIZARD_TYPE );
        }

        var CheckResult = -10;
        CheckResult = objProfile.CheckData( ADD_SIGN_WIZARD_TYPE );
        if( ALL_OK == CheckResult )
        {
            //������� �������
            oPKCS7Message.Sign();
            //�������� ������
            oPKCS7Message.Save( DT_SIGNED_DATA, BASE64_TYPE, OUTPUT_DATA_FILE );
            oPKCS7Message = null;

            return true;
        }
        else
        {
            var msgBox = new ActiveXObject("wscript.shell");
            msgBox.Popup( "������� ����������� ��������" );
        }
    }
    else
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "��� ����������� � ���������" );
    }

    return false;
}

//-----------------------------------------------------------------------------

//Adding cosignature
//��������� �������

function AddCosignature( sSignName, sResultSignName, bDisplayUI )
{
    // � ��������� ������ ���� ������� � ������������ ����������� ���

    var SIGN_DATA_FILE = sSignName;
    var OUTPUT_DATA_FILE = sResultSignName;

    //������� ������� �� ��������� ��� �������� �����, ���� ��� ���

    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open( REGISTRY_STORE ); //��������� ��������� ��������
    var objProfiles = objProfileStore.Store; //�������� ��������� ��������

    var objProfile;
    if( objProfiles.Count > 0 )
    {
        objProfile = objProfiles.DefaultProfile; //������� ������� �� ���������
    }
    else
    {
        objProfile = new ActiveXObject("DigtCrypto.Profile"); //�������� ����� �������
    }

    //�������� ���� �������������� �������
    var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
    oPKCS7Message.Load( DT_SIGNED_DATA, String(SIGN_DATA_FILE), "" );

    //������� ��������� ��������
    var oSignatures = oPKCS7Message.Signatures;
    if( oSignatures.Count > 0 )
    {
        oPKCS7Message.Profile = objProfile;  //��������� ������� � �����������
        if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
        {
            objProfile.CollectData( COSIGN_WIZARD_TYPE );
        }
        var CheckResult = -10;
        CheckResult = objProfile.CheckData( COSIGN_WIZARD_TYPE );
        if( ALL_OK == CheckResult )
        {
            oPKCS7Message.Profile = objProfile;
            //������� ������� c ���������� �������� 0
            oPKCS7Message.Cosign( 0 );
            oPKCS7Message.Save( DT_SIGNED_DATA, BASE64_TYPE, OUTPUT_DATA_FILE ); //�������� ������ 
            oPKCS7Message = null;

            return true;
        }
        else
        {
            var msgBox = new ActiveXObject("wscript.shell");
            msgBox.Popup( "������� ����������� ��������" );
        }
    }
    else
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "��� ����������� � ���������" );
    }

    return false;
}

//-----------------------------------------------------------------------------

//Opening the window of a management of a PKCS7 data  
//�������� ���� ���������� PKCS7 �������

function OpenPkcs7MsgManagementWindow( sSignFile, sEncryptedFile )
{
    // � ��������� ������ ���� ������� � ������������ ����������� ���

    var SINGED_DATA_FILE = sSignFile;
    var ENCRYPTED_DATA_FILE = sEncryptedFile;

    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open( REGISTRY_STORE ); //��������� ��������� ��������
    var objProfiles = objProfileStore.Store; //�������� ��������� ��������

    var objProfile;
    if( objProfiles.Count > 0 )
    {
        objProfile = objProfiles.DefaultProfile; //������� ������� �� ���������
    }
    else
    {
        objProfile = new ActiveXObject("DigtCrypto.Profile"); //�������� ����� �������
    }

    //�������� ���� �������
    var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
    oPKCS7Message.Profile = objProfile; //��������� ������� � �����������
    oPKCS7Message.Load( DT_SIGNED_DATA, SINGED_DATA_FILE, "" );
    oPKCS7Message.Display();

    //�������� ���� ����������� ������
    oPKCS7Message = null;
    oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
    oPKCS7Message.Profile = objProfile; //��������� ������� � �����������
    oPKCS7Message.Load( DT_ENVELOPED_DATA, ENCRYPTED_DATA_FILE, "" );
    oPKCS7Message.Display();
}

//-----------------------------------------------------------------------------

//Data encryption
//���������� ������

function EncryptData( sPlainFile, sEncryptedResultFile, bDisplayUI )
{
    // � ��������� ������ ���� ������� � ������������ ����������� ���

    var PLAIN_DATA_FILE = sPlainFile;
    var OUTPUT_DATA_FILE = sEncryptedResultFile;

    //������� ������� �� ��������� ��� �������� �����, ���� ��� ���

    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open( REGISTRY_STORE ); //��������� ��������� ��������
    var objProfiles = objProfileStore.Store; //�������� ��������� ��������

    var objProfile;
    if( objProfiles.Count > 0 )
    {
        objProfile = objProfiles.DefaultProfile; //������� ������� �� ���������
    }
    else
    {
        objProfile = new ActiveXObject("DigtCrypto.Profile"); //�������� ����� �������
    }

    //��������� � ��������� �������, ��������� ������, ���������� �� �������
    if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
    {
        objProfile.CollectData( ENCRYPT_WIZARD_TYPE ); //�������� ������ ������� ��� ����� ������
    }
    var CheckResult = objProfile.CheckData(ENCRYPT_WIZARD_TYPE); //��������, ��� �� ������ �������
    if( ALL_OK == CheckResult )
    {
        var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
        oPKCS7Message.Profile = objProfile; //��������� ������� � �����������
        oPKCS7Message.Load( DT_PLAIN_DATA, PLAIN_DATA_FILE, "" ); //�������� �������� ������
        oPKCS7Message.Encrypt();
        oPKCS7Message.Save( DT_ENVELOPED_DATA, BASE64_TYPE, OUTPUT_DATA_FILE ); //�������� ������ 
        oPKCS7Message = null;

        return true;
    }
    else
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "������� ����������� ��������" );
    }

    return false;
}

//-----------------------------------------------------------------------------

//Encryption of data blocks on the same session key (working with data in memory)
//Example 1 Header of the encrypted message is combined with the first block of ciphertext
//���������� ������ ������ �� ����� ���������� ����� (������ � ������� � ������)
//������ 1 ��������� ������������ ��������� �������� � ������ ������ ����������� ������

function EncryptByBlocksFromMemAttachedHeader( fileNameBlock1, fileNameBlock2 )
{
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );

    //�������� ��������� �������� � ��������� ������� �� ���������
    //� ������� ������ ���� ����������� ����������� ��� ���������� ���������� ������
    var oProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    var oProfile = new ActiveXObject("DigtCrypto.Profile");
    oProfileStore.Open( REGISTRY_STORE );

    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");

    //��������� �������, � �������������� �������� ����� ����������� ����������
    oMsg.Profile = oProfileStore.Store.DefaultProfile;

    //������������� �������
    oMsg.Open( DT_PLAIN_DATA );

    //������ � ������ ������� ����� ������
    oMsg.Import( "������ ���� ������" );

    //���������� �������������� ������� ����� ������ ������ �  ������ ������� (����������) ������������ ��������� � ����
    var strStroka = oMsg.Export();
    var file = fso.CreateTextFile( fileNameBlock1, True );
    file.write(strStroka);
    file.close();

    //����������� ������ ���� ������
    oMsg.Import( "������ ���� ������" );

    //���������� �������������� ������� ����� ������ � ����
    strStroka = oMsg.Export();

    file = fso.CreateTextFile( fileNameBlock1, True );
    file.write(strStroka);
    file.close();
}

//-----------------------------------------------------------------------------

//Encryption of data blocks on the same session key (working with data in memory)
//Example 2 Header of the encrypted message is stored separately from the first block of ciphertext
//���������� ������ ������ �� ����� ���������� ����� (������ � ������� � ������)
//������ 2 ��������� ������������ ��������� �������� �������� �� ������� ����� ����������� ������

function EncryptByBlocksFromMemDetachedHeader( fileNameHeader, fileNameBlock1, fileNameBlock2 )
{
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );

    //�������� ��������� �������� � ��������� ������� �� ���������
    //� ������� ������ ���� ����������� ����������� ��� ���������� ���������� ������
    var oProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    oProfileStore.Open( REGISTRY_STORE );

    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");

    //��������� �������, � �������������� �������� ����� ����������� ����������
    oMsg.Profile = oProfileStore.Store.DefaultProfile;

    //������������� �������
    oMsg.Open( DT_PLAIN_DATA );
    //���������� ������ ������� (����������) ������������ ��������� � ����
    var strStroka = oMsg.Export();
    var file = fso.CreateTextFile(fileNameHeader, true);
    file.write(strStroka);
    file.close();

    //������ � ������ ������� ����� ������
    oMsg.Import( "������ ���� ������" );

    //���������� �������������� ������� ����� ������ � ����
    strStroka = oMsg.Export();
    file = fso.CreateTextFile(fileNameBlock1, True);
    file.write(strStroka);
    file.close();

    //������ � ������ ������� ����� ������
    oMsg.Import( "������ ���� ������" );

    //���������� �������������� ������� ����� ������ � ����
    strStroka = oMsg.Export();
    file = fso.CreateTextFile(fileNameBlock1, True);
    file.write(strStroka);
    file.close();
}

//-----------------------------------------------------------------------------

//Encryption of data blocks on the same session key (working with files)
//Example 3 Header of the encrypted message is combined with the first block of ciphertext
//���������� ������ ������ �� ����� ���������� ����� (������ � �������)
//������ 3 ��������� ������������ ��������� �������� � ������ ������ ����������� ������

function EncryptByBlocksFromFileAttachedHeader( fileNameSourceBlock1, fileNameSourceBlock2, fileNameBlock1, fileNameBlock2 )
{
    var oProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    //�������� ��������� �������� � ��������� ������� �� ���������
    //� ������� �� ��������� ������ ���� ����������� ����������� ��� ���������� ���������� ������
    oProfileStore.Open( REGISTRY_STORE );

    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");
    //��������� �������, � �������������� �������� ����� ����������� ����������
    oMsg.Profile = oProfileStore.Store.DefaultProfile;

    //������������� �������
    oMsg.Open( DT_PLAIN_DATA );

    //�������� ������� ����� � �������� �������
    oMsg.Load( fileNameSourceBlock1 );
    //���������� ������������ ������� ����� � �������� ������� ������ � ���������� ������������ ���������
    oMsg.Save( fileNameBlock1 );

    //�������� ������� ����� � �������� �������
    oMsg.Load( fileNameSourceBlock2 );
    //���������� ������������ ������� ����� � �������� ������� ������ � ���������� ������������ ���������
    oMsg.Save( fileNameBlock2 );
}

//-----------------------------------------------------------------------------

//Encryption of data blocks on the same session key (working with files)
//Example 4 Header of the encrypted message is stored separately from the first block of ciphertext
//���������� ������ ������ �� ����� ���������� ����� (������ � �������)
//������ 4 ��������� ������������ ��������� �������� �������� �� ������� ����� ����������� ������

function EncryptByBlocksFromFileDetachedHeader( fileNameSourceBlock1, fileNameSourceBlock2, fileNameHeader, fileNameBlock1, fileNameBlock2 )
{
    var oProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    //�������� ��������� �������� � ��������� ������� �� ���������
    //� ������� �� ��������� ������ ���� ����������� ����������� ��� ���������� ���������� ������
    oProfileStore.Open( REGISTRY_STORE );

    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");
    //��������� �������, � �������������� �������� ����� ����������� ����������
    oMsg.Profile = oProfileStore.Store.DefaultProfile;

    //������������� �������
    oMsg.Open( DT_PLAIN_DATA );
    //���������� ��������� ������������ ��������� � ��������� ����
    oMsg.Save( fileNameHeader );
    //�������� ������� ����� � �������� �������
    oMsg.Load( fileNameSourceBlock1 );
    //���������� ������������ ������� ����� � �������� ������� 
    oMsg.Save( fileNameBlock1 );
    //�������� ������� ����� � �������� �������
    oMsg.Load( fileNameSourceBlock2 );
    //���������� ������������ ������� ����� � �������� ������� 
    oMsg.Save( fileNameBlock2 );
}

//-----------------------------------------------------------------------------

//Data decryption
//������������� ������

function DecryptData( sEncryptedFile, sResultDecryptedFile, bDisplayUI )
{
    // � ��������� ������ ���� ������� � ������������ ����������� ���

    var INPUT_DATA_FILE = sEncryptedFile;
    var OUTPUT_DATA_FILE = sResultDecryptedFile;

    //������� ������� �� ��������� ��� �������� �����, ���� ��� ���

    var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
    objProfileStore.Open( REGISTRY_STORE ); //��������� ��������� ��������
    var objProfiles = objProfileStore.Store; //�������� ��������� ��������

    var objProfile;
    if( objProfiles.Count > 0 )
    {
        objProfile = objProfiles.DefaultProfile; //������� ������� �� ���������
    }
    else
    {
        objProfile = new ActiveXObject("DigtCrypto.Profile"); //�������� ����� �������
    }

    //��������� � ��������� �������, ��������� ������, ���������� �� �������
    if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
    {
        objProfile.CollectData( DECRYPT_WIZARD_TYPE ); //�������� ������ ������������� ��� ����� ������
    }
    var CheckResult = objProfile.CheckData(DECRYPT_WIZARD_TYPE); //��������, ��� �� ������ �������
    if( ALL_OK == CheckResult )
    {
        var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
        oPKCS7Message.Profile = objProfile; //��������� ������� � �����������
        oPKCS7Message.Load( DT_ENVELOPED_DATA, INPUT_DATA_FILE, "" ); //�������� �������� ������
        oPKCS7Message.Decrypt();
        oPKCS7Message.Save( DT_PLAIN_DATA, BASE64_TYPE, OUTPUT_DATA_FILE ); //�������� ������ 
        oPKCS7Message = null;

        return true;
    }
    else
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( "������� ����������� ��������" );
    }

    return false;
}

//-----------------------------------------------------------------------------

//Decryption of data blocks on the same session key (working with memory data)
//Example 1 Header of the encrypted message is combined with the first block of ciphertext
//������������� ������ ������ �� ����� ���������� �����
//������ 1 ��������� ������������ ��������� �������� � ������ ������ ����������� ������

function DecryptByBlocksFromMemAttachedHeader( fileNameBlock1, fileNameBlock2, fileNameResultBlock1, fileNameResultBlock2 )
{
    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );

    //������������� �������
    oMsg.Open( DT_ENCRYPTED_DATA, fileNameBlock1 );
    //������������ �������������� ������ � ������
    strStroka1 = oMsg.Export();

    //�������� ������ ���� �������������� ������ � ����
    var file = fso.CreateTextFile(fileNameResultBlock1, True);
    file.write(strStroka1);
    file.close();
    strStroka1 = "";

    //������� � ������ ������ ���� ������������� ������
    var file1 = fso.OpenTextFile(fileNameBlock2, 1);
    strStroka = file1.ReadAll();
    file1.Close();

    //����������� � ������ ������ ���� ������
    oMsg.Import( strStroka );
    //������������ �������������� ������ � ������
    strStroka1 = oMsg.Export();

    //�������� �������������� ������ � ����
    file = fso.CreateTextFile(fileNameResultBlock2, True);
    file.write(strStroka1);
    file.close();
}

//-----------------------------------------------------------------------------

//Decryption of data blocks on the same session key (working with memory data)
//Example 2 Header of the encrypted message is stored separately from the first block of ciphertext
//������������� ������ ������ �� ����� ���������� �����
//������ 2 ��������� ������������ ��������� �������� �������� �� ������� ����� ����������� ������

function DecryptByBlocksFromMemDetachedHeader( fileNameHeader, fileNameBlock1, fileNameBlock2, fileNameResultBlock1, fileNameResultBlock2 )
{
    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );

    //������������� ������� � ��������� ��������� ������������ ���������
    oMsg.Open( DT_ENCRYPTED_DATA, fileNameHeader );

    //������� � ������ ������ ���� ������������� ������
    var file1 = fso.OpenTextFile(fileNameBlock1, 1);
    var strStroka = file1.ReadAll();
    file1.Close();

    //����������� � ������ ������ ���� ������
    oMsg.Import( strStroka );
    //������������ �������������� ������ � ������
    var strStroka1 = oMsg.Export();

    //�������� ������ ���� �������������� ������ � ����
    var file = fso.CreateTextFile(fileNameResultBlock1, True);
    file.write(strStroka1);
    file.close();
    strStroka1 = "";

    //������� � ������ ������ ���� ������������� ������
    file1 = fso.OpenTextFile(fileNameBlock2, 1);
    strStroka = file1.ReadAll();
    file1.Close();

    //����������� � ������ ������ ���� ������
    oMsg.Import( strStroka );
    //������������ �������������� ������ � ������
    strStroka1 = oMsg.Export();

    //�������� �������������� ������ � ����
    file = fso.CreateTextFile(fileNameResultBlock2, True);
    file.write(strStroka1);
    file.close();
}

//-----------------------------------------------------------------------------

//Decryption of data blocks on the same session key (working with files)
//Example 3 Header of the encrypted message is combined with the first block of ciphertext
//������������� ������ ������ �� ����� ���������� �����
//������ 3 ��������� ������������ ��������� �������� � ������ ������ ����������� ������

function DecryptByBlocksFromFileAttachedHeader( fileNameBlock1, fileNameBlock2, fileNameResultBlock1, fileNameResultBlock2 )
{
    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");
    //������������� �������
    oMsg.Open( DT_ENCRYPTED_DATA, fileNameBlock1 );
    //��������� ������ �������������� ���� ������ � ����
    oMsg.Save( fileNameResultBlock1 );
    //�������� ������� ����� � ������������  �������
    oMsg.Load( fileNameBlock2 );
    //���������� ������� ��������������� ����� ������ � ����
    oMsg.Save( fileNameResultBlock2 );
}

//-----------------------------------------------------------------------------

//Decryption of data blocks on the same session key (working with files)
//Example 4 Header of the encrypted message is stored separately from the first block of ciphertext
//������������� ������ ������ �� ����� ���������� �����
//������ 4 ��������� ������������ ��������� �������� �������� �� ������� ����� ����������� ������

function DecryptByBlocksFromFileDetachedHeader( fileNameHeader, fileNameBlock1, fileNameBlock2, fileNameResultBlock1, fileNameResultBlock2 )
{
    var oMsg = new ActiveXObject("DigtCrypto.EncryptedMessage");
    //������������� �������
    oMsg.Open( DT_ENCRYPTED_DATA, fileNameHeader );
    //�������� ������� ����� �  ������������  �������
    oMsg.Load( fileNameBlock1 );
    //���������� ������� ��������������� ����� ������ � ����
    oMsg.Save( fileNameResultBlock1 );
    //�������� ������� ����� � ������������  �������
    oMsg.Load( fileNameBlock2 );
    //���������� ������� ��������������� ����� ������ � ����
    oMsg.Save( fileNameResultBlock2 );
}

//-----------------------------------------------------------------------------

