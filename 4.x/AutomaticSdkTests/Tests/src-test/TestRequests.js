
TestCase("RequestsManagement",{
    // Objects
    fso: null,

    // Files
    strInXmlTemplateFileName:"",
    strOutRequest1FileName:"",
    strOitRequesr2FileName:"",

    // Stored constants
    strTmpFolder:"",
    strRequestTemplateContent:"<?xml version=\"1.0\" encoding=\"windows-1251\"?><CertificateRequestTemplate><TemplateInformation><SuppotedCA></SuppotedCA><FriendlyName>Шаблон по умолчанию</FriendlyName><Description></Description></TemplateInformation><Subject><DirectoryName><RDNSequence><RDNEntry readonly=\"false\" mandatory=\"true\" hidden=\"false\"><OID>CN</OID><Name>Идентификатор (CN)</Name><Value>John Doe</Value><Length>64</Length></RDNEntry><RDNEntry><OID>2.5.4.10</OID><Name>Организация</Name><Length>64</Length><Value>Рога и копыта</Value></RDNEntry><RDNEntry><OID>L</OID><Name>Город</Name><Length>64</Length></RDNEntry><RDNEntry><OID>S</OID><Name>Область</Name><Length>64</Length></RDNEntry><RDNEntry><OID>C</OID><Name>Страна</Name><Value>RU</Value></RDNEntry><RDNEntry><OID>1.2.840.113549.1.9.1</OID><Name>E-mail</Name><Length>128</Length><Value>johndoe@example.com</Value></RDNEntry><RDNEntry><OID>1.2.643.3.131.1.1</OID><Name>ИНН</Name><Length>12</Length></RDNEntry></RDNSequence></DirectoryName></Subject><Extensions><Extension><OID>2.5.29.15</OID><Critical>True</Critical><Value><KeyUsage><Bits><KeyAgreement/><DataEncipherment/><NonRepudiation/><DigitalSignature/></Bits></KeyUsage></Value></Extension><Extension><OID>2.5.29.37</OID><Value><ExtendedKeyUsage><KeyPurposeId>1.3.6.1.5.5.7.3.2</KeyPurposeId><KeyPurposeId>1.3.6.1.5.5.7.3.4</KeyPurposeId></ExtendedKeyUsage></Value></Extension></Extensions><Provider><Name>Microsoft Base Cryptographic Provider v1.0</Name><Type>1</Type></Provider><Keyset><CreateNew>true</CreateNew><Keyspec>3</Keyspec><MarkExportable>false</MarkExportable></Keyset></CertificateRequestTemplate>",

    setUp:function()
    {
        // Init file system object
        this.fso = new ActiveXObject("Scripting.FileSystemObject");

        // Init temporary folder
        this.strTmpFolder = this.fso.GetSpecialFolder(2) + "\\";

        // Init files names and paths
        this.strInXmlTemplateFileName = this.strTmpFolder + "ReqTemplate.xml";
        this.strOutRequest1FileName = this.strTmpFolder + "outReq1-from-template.pem";
        this.strOitRequesr2FileName = this.strTmpFolder + "ReqFileName.pem";

        // Save template file
        var tmpltFile = this.fso.CreateTextFile( this.strInXmlTemplateFileName );
        tmpltFile.Write(this.strRequestTemplateContent);
        tmpltFile.Close();
    },

    tearDown:function()
    {
        this.fso.DeleteFile( this.strInXmlTemplateFileName );

        //TODO: Remove requests key containers
        if( this.fso.FileExists( this.strOutRequest1FileName ) )
        {
            this.fso.DeleteFile( this.strOutRequest1FileName );
        }

        if( this.fso.FileExists( this.strOutRequest2FileName ) )
        {
            this.fso.DeleteFile( this.strOutRequest2FileName );
        }
    },

    testRequestsManagement:function()
    {
        expectAsserts(10);

        //CreateRequestWithTemplate
        {
            var oReqFromTemplate = CreateRequestWithTemplate( this.strInXmlTemplateFileName, this.strOutRequest1FileName, false );

            assertObject( "CreateRequestWithTemplate function result", oReqFromTemplate );
            assertTrue( "Request saving result (file exists)", this.fso.FileExists(this.strOutRequest1FileName) );

            assertEquals( "DN, loaded from template", "CN=John Doe,O=Рога и копыта,C=RU,E=johndoe@example.com", oReqFromTemplate.Template.DN );
            assertEquals( "Cryptoprovider", "Microsoft Base Cryptographic Provider v1.0", oReqFromTemplate.Template.CryptoProvider );
        }//4 asserts

        //RequestExportAndImport
        {
            var oRequest = RequestExportAndImport(this.strOutRequest2FileName);

            assertNotNull( "Returned object", oRequest );

            assertEquals( "Distinguished name", "CN=John Doe,C=RU,E=johndoe@example.com", oRequest.Template.DN );
            assertEquals( "Common name", "John Doe", oRequest.Template.CN );
            assertEquals( "e-mail", "johndoe@example.com", oRequest.Template.E );
            assertEquals( "Country", "RU", oRequest.Template.C );
            //assertEquals( "Cryptoprovider", "Microsoft Base Cryptographic Provider v1.0", oRequest.Template.CryptoProvider );
            assertEquals( "Extended key usage", "<keyPurposeId>1.3.6.1.5.5.7.3.1</keyPurposeId>", oRequest.Template.ExtendedKeyUsage );
        }//6 asserts

        //jstestdriver.console.log("Done", "");
    }
});

TestCase("RequestsCAClient",{
    // Objects
    fso: null,

    // Files
    strOutReq1Revoke:"",
    strOutReq2Resume:"",
    strOutReq3Renew:"",
    strOutReq4NewCert:"",

    // Stored constants
    strTmpFolder:"",


    setUp:function()
    {
        // Init file system object
        this.fso = new ActiveXObject("Scripting.FileSystemObject");

        // Init temporary folder
        this.strTmpFolder = this.fso.GetSpecialFolder(2) + "\\";

        // Init files names and paths
        this.strOutReq1Revoke = this.strTmpFolder + "Revoke.pem";
        this.strOutReq2Resume = this.strTmpFolder + "Resume.pem";
        this.strOutReq3Renew = this.strTmpFolder + "key_update.pem";
        this.strOutReq4NewCert = this.strTmpFolder + "cert_req.pem";
    },

    tearDown:function()
    {
        //TODO: Remove requests key containers
        if( this.fso.FileExists( this.strOutReq1Revoke ) )
        {
            this.fso.DeleteFile( this.strOutReq1Revoke );
        }

        if( this.fso.FileExists( this.strOutReq2Resume ) )
        {
            this.fso.DeleteFile( this.strOutReq2Resume );
        }

        if( this.fso.FileExists( this.strOutReq3Renew ) )
        {
            this.fso.DeleteFile( this.strOutReq3Renew );
        }

        if( this.fso.FileExists( this.strOutReq4NewCert ) )
        {
            this.fso.DeleteFile( this.strOutReq4NewCert );
        }
    },

    testRequestsCAClient:function()
    {
        expectAsserts(9);

        //CreateRevocationRequest
        {
            assertEquals( "CreateRevocationRequest function returned value", REQUEST_GENERATION_SUCCESS, CreateRevocationRequest( this.strOutReq1Revoke, false ) );
            assertTrue( "Result file exists", this.fso.FileExists(this.strOutReq1Revoke));
        }//2 asserts

        //CreateResumingRequest
        {
            assertEquals( "CreateResumingRequest function returned value", REQUEST_GENERATION_SUCCESS, CreateResumingRequest( this.strOutReq2Resume, false ) );
            assertTrue( "Result file exists", this.fso.FileExists(this.strOutReq2Resume));
        }//2 asserts

        //CreateRenewingRequest
        {
            assertEquals( "CreateRenewingRequest function returned value", REQUEST_GENERATION_SUCCESS, CreateRenewingRequest( this.strOutReq3Renew, false ) );
            assertTrue( "Result file exists", this.fso.FileExists(this.strOutReq3Renew));
        }//2 assert

        //CreatingCertificateRequest
        {
            var oPkiReq = CreatingCertificateRequest(this.strOutReq4NewCert);
            assertNotNull( "CreatingCertificateRequest function returned value", oPkiReq );
            assertObject( "CreatingCertificateRequest function returned value", oPkiReq );

            assertTrue( "Result file exists", this.fso.FileExists(this.strOutReq4NewCert));
            //assertEquals( "Distinguished name", "CN=John Doe,C=RU,E=johndoe@example.com", oPkiReq.RequestTemplate.DN );
        }//3 asserts

        //jstestdriver.console.log("Done", "");
    }
});

