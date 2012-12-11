
//-----------------------------------------------------------------------------

//Signature status verification
//Проверка статуса подписи

function SignatureStatusVerification( sFileName, bDisplayUI )
{
    var CERT_AND_SIGN = 0;
    var SIGN_ONLY = 1;
    var DT_SIGNED_DATA = 2;

    var oPKCS7Message = new ActiveXObject( "DigtCrypto.PKCS7Message" );
    oPKCS7Message.Load( DT_SIGNED_DATA, sFileName, "" );
    //Получим коллекцию ЭЦП
    var oSignatures = oPKCS7Message.Signatures;

    //oPKCS7Message.Display();
    //Получим количество подписей в коллекции
    var n = oSignatures.Count;
    //MsgBox n, , "Количество подписей"
    var oSignature;

    var msgBox = new ActiveXObject("wscript.shell");
    var resultArray = new Array();
    for( var i = 0; i < n; i++ )
    {
        //Получим указатель на элемент коллекции подписей
        oSignature = oSignatures.Item(i);
        var Status = oSignature.Verify(SIGN_ONLY, "test.txt");
        if( (typeof(bDisplayUI)=="undefined") || bDisplayUI )
        {
            msgBox.Popup( "Status: " + String(Status) );
        }
        resultArray.push( Status );
    }

    return resultArray;
}

//-----------------------------------------------------------------------------

//Getting properties of a signature
//Получение свойств подписи

function GettingSignaturesProperties( sSignFileName, bDisplayUI )
{
    var DT_SIGNED_DATA = 2;

    var oPKCS7Message = new ActiveXObject( "DigtCrypto.PKCS7Message" );
    //Загрузка присоединенной подписи
    oPKCS7Message.Load( DT_SIGNED_DATA, sSignFileName, "" );

    //Получим коллекцию ЭЦП
    var oSignatures = oPKCS7Message.Signatures;
    //Получим количество подписей
    var n = oSignatures.Count;
    //var oSigns = new ActiveXObject("DigtCrypto.Signatures");

    var oSignature;
    var oSign;
    var oCertificate;
    var sResult = "";

    var msgBox = new ActiveXObject("wscript.shell");

    var returnResult = "";

    for( var i = 0; i < n; i++ )
    {
        //Получим свойства для каждой подписи  
        oSignature = oSignatures.Item(i);

        if( bDisplayUI )
        {
            //Просмотрим подпись
            oSignature.Display();
        }

        //Получаем индекс подписчика
        var uLongSignerIndex = oSignature.SignerIndex;
        sResult += "Индекс подписчика  " + String(uLongSignerIndex) + "\n";

        //Получаем комментарий подписи
        var sComments = oSignature.Comments;
        sResult += "Комментарий подписи  " + String(sComments) + "\n";

        //Получаем время подписи
        var sSigningTime = oSignature.SigningTime;
        sResult += "Время подписи  " + String(sSigningTime) + "\n";

        //Получаем идентификатор ресурса подписи
        var sResource = oSignature.Resource;
        sResult += "Идентификатор ресурса подписи  " + String(sResource) + "\n";

        //Получаем хэш алгоритм подписи
        var sHashAlg = oSignature.HashAlg;
        sResult += "Хэш алгоритм подписи  " + String(sHashAlg) + "\n";

        //Получаем алгоритм подписи ЭЦП
        var sHashEncAlg = oSignature.HashEncAlg;
        sResult += "Алгоритм подписи ЭЦП  " + String(sHashEncAlg) + "\n";

        //Получаем тип эцп: 1-отсоединенная, 0-присоединенная
        var sDetached = oSignature.Detached;
        sResult += "Тип эцп  " + String(sDetached) + "\n";

        //Получаем номер версии протокола CMS
        var lCMSVersion = oSignature.CMSVersion;
        sResult += "Номер версии протокола CMS  " + String(lCMSVersion) + "\n";

        //Получаем Тип Содержимого сообщения
        var sContentType = oSignature.ContentType;
        sResult += "Тип содержимого сообщения  " + String(sContentType) + "\n";

        if( bDisplayUI )
        {
            msgBox.Popup( sResult, 0, "Свойства подписи" );
        }

        //Получаем указатель на сертификат подписи
        oCertificate = oSignature.Certificate;
        if( bDisplayUI )
        {
            oCertificate.Display();
        }

        returnResult += "Подпись " + String(i+1) + "\n\n" + sResult;
        returnResult += "Сертификат подписи: " + oCertificate.SubjectName + "\n\n\n";

        sResult = "";
    }

    return returnResult;
}

//-----------------------------------------------------------------------------

//Getting collection of cosignatures
//Получение коллекции заверяющих подписей

function GettingCollectionOfCosigns( sSignFile, bDisplayUI )
{
    var DT_SIGNED_DATA = 2;

    var oPKCS7Message = new ActiveXObject( "DigtCrypto.PKCS7Message" );
    //Загрузка присоединенной подписи
    oPKCS7Message.Load( DT_SIGNED_DATA, sSignFile, "" );

    //Получим коллекцию ЭЦП
    var oSignatures = oPKCS7Message.Signatures;
    //Получим количество подписей
    var n = oSignatures.Count;
    //var oSigns = new ActiveXObject("DigtCrypto.Signatures");

    var oSignature;
    //var oSigns;
    var sResult = "";

    for( var i = 0; i < n; i++ )
    {
        //Получим свойства для каждой подписи  
        oSignature = oSignatures.Item(i);

        //Получим подписи коллекцию встречных (заверяющих) подписей
        oSigns = oSignature.Cosignature;
        sResult += "Количество встречных подписей у подписи #" + String(i+1) + " = " + String(oSigns.Count) + "\n";
    } 

    if( bDisplayUI )
    {
        var msgBox = new ActiveXObject("wscript.shell");
        msgBox.Popup( sResult );
    }

    return sResult;
}

//-----------------------------------------------------------------------------

//Signing the form content
//Подпись содержимого формы

/*

<html>
<head>
<title>Signing test</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<script language="JavaScript">
<!--

function OnOK()
{
    var form = document.form;

    var str = form.text_to_sign.value;
    if (str == "")
    {
        alert("No text found!");
        return;
    }

    var signed_text = SignText(str);
    if (signed_text == "")
    {
        alert("No signature found!");
        return;
    }

    alert(signed_text);

    //form.signature.value = signed_text;
    //form.submit();
}

*/

function SignText(text)
{
    var retStr = "";

    try
    {
        var oCertificate = new ActiveXObject("DigtCrypto.Certificate");
        var oPKCS7Message = new ActiveXObject("DigtCrypto.PKCS7Message");
        var oProfile = new ActiveXObject("DigtCrypto.Profile");

        if (text != "")
        {
            var DT_PLAIN_DATA = 0;
            var DT_SIGNED_DATA = 2;
            var BASE64_TYPE = 0;
            var SIGN_WIZARD_TYPE = 1;
            var SILENT_LEVEL_REQUIRED = 3;

            var isDetachedSign = true; // generate detached (true) or attached (false) signature

            oProfile.SilentLevel = SILENT_LEVEL_REQUIRED;
            oProfile.DisableInputFilesWindow = true;
            if ( oProfile.CollectData(SIGN_WIZARD_TYPE) ) // it returns false when user cancel wizard
            {
                oProfile.SignIncludeBase64Headers = false;
                oProfile.Detach = isDetachedSign;

                //if ( oProfile.CheckData(SIGN_WIZARD_TYPE) ) // may be skipped if is called after .CollectData()
                //{
                //    throw "Not enough data (code: " + oProfile.CheckData(SIGN_WIZARD_TYPE) + ")!";
                //}

                oPKCS7Message.Import(DT_PLAIN_DATA, text); // loading of data

                oPKCS7Message.Profile = oProfile;
                oPKCS7Message.Sign(); // signing

                retStr = oPKCS7Message.Export(DT_SIGNED_DATA, BASE64_TYPE); // exporting of signed document
            }
        }
    }
    catch(e) 
    { 
        alert("Exception catched: " + ((typeof e == "object") ? e.description : e));
    }

    return retStr;
}

/*

//--> 
</script>
</head>
<body class="popup" onload="window.focus();"> 
<form name="form" method="post" action="sign_form.asp">
    <input type="hidden" name="signature" value="">
    <table border="0" cellspacing="0" cellpadding="0" class="popup">
        <tr><td>Enter a text to sign</td></tr>
        <tr><td>
            <input type="text" name="text_to_sign" value="" size="40">
            <input type="button" name="btnOK" value="Submit" OnClick="OnOK()">
        </td></tr>
    </table>
</form>
</body>
</html>


*/

//-----------------------------------------------------------------------------

