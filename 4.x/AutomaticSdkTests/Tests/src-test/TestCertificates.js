
// To prepare test stand:
//  - Must ensure that at least one certificate is available in the personal store
//    (self signed is also suitable).
//  - After 14.10.2014 replace value of a strCertString member with a new certificate
//  - If status of the first certificate in the store is not "correct", in the test code
//    for "VerifyCertificateStatus" replace expecting status to the actual value.

TestCase("Certificates",{
    // Objects
    fso: null,

    // Files
    inCertName:"",
    outCertName:"",
    certForComparing:"",

    // Stored constants
    strTmpFolder:"",
    strCertString:"-----BEGIN CERTIFICATE-----\nMIICQzCCAfCgAwIBAgIQaYQDKGqmWbpGNWItSd5f0zAKBgYqhQMCAgMFADBlMSAw\nHgYJKoZIhvcNAQkBFhFpbmZvQGNyeXB0b3Byby5ydTELMAkGA1UEBhMCUlUxEzAR\nBgNVBAoTCkNSWVBUTy1QUk8xHzAdBgNVBAMTFlRlc3QgQ2VudGVyIENSWVBUTy1Q\nUk8wHhcNMDkwNDA3MTIwMjE1WhcNMTQxMDA0MDcwOTQxWjBlMSAwHgYJKoZIhvcN\nAQkBFhFpbmZvQGNyeXB0b3Byby5ydTELMAkGA1UEBhMCUlUxEzARBgNVBAoTCkNS\nWVBUTy1QUk8xHzAdBgNVBAMTFlRlc3QgQ2VudGVyIENSWVBUTy1QUk8wYzAcBgYq\nhQMCAhMwEgYHKoUDAgIjAQYHKoUDAgIeAQNDAARAAuT/0ab2nICa2ux/SnjBzC3T\n5Zbqy+0iMnmyAuLGfDXmdGQbCXcRjGc/D9DoI6Z+bTt/xMQo/SscaAEgoFzYeaN4\nMHYwCwYDVR0PBAQDAgHGMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFG2PXgXZ\nX6yRF5QelZoFMDg3ehAqMBIGCSsGAQQBgjcVAQQFAgMCAAIwIwYJKwYBBAGCNxUC\nBBYEFHrJxwnbIByWlC/8Rq1tk9BeaRIOMAoGBiqFAwICAwUAA0EAWHPSk7xjIbEO\nc3Lu8XK1G4u7yTsIu0xa8uGlNU+ZxNVSUnAm3a7QqSfptlt9b0T9Jk39oWN0XHTY\nSXMKd3djTQ==\n-----END CERTIFICATE-----\n",

    setUp:function()
    {
        // Init file system object
        this.fso = new ActiveXObject("Scripting.FileSystemObject");

        // Init temporary folder
        this.strTmpFolder = this.fso.GetSpecialFolder(2) + "\\";

        // Init files names and paths
        this.inCertName = this.strTmpFolder + "in_cert.cer";
        this.outCertName = this.strTmpFolder + "out_cert.cer";
        this.certForComparing = this.strTmpFolder + "cmp_cert.cer";

        // Save certificate file
        var certFile = this.fso.CreateTextFile( this.inCertName );
        certFile.Write(this.strCertString);
        certFile.Close();

        // Extract certificate from store to disk
        {
            var certStore = new ActiveXObject("DigtCrypto.CertificateStore");
            certStore.Open( 1, "my" );
            certStore.Store.Item(0).Save(this.certForComparing,0,true);
        }
    },

    tearDown:function()
    {
        this.fso.DeleteFile( this.inCertName );
        if( this.fso.FileExists(this.outCertName) )
        {
            this.fso.DeleteFile( this.outCertName );
        }
        this.fso.DeleteFile( this.certForComparing );
    },

    testCertificates:function()
    {
        expectAsserts(15);

        //ImportExportAndSaving
        {
            var resultObject = ImportExportAndSaving(this.inCertName, this.outCertName, false);
            assertObject( "ImportExportAndSaving function result", resultObject );
            assertTrue( "\"Certificate successfully saved (file exists)\"", this.fso.FileExists(this.outCertName) );
            assertEquals( "Issuer name of certificate", "CN=Test Center CRYPTO-PRO,O=CRYPTO-PRO,C=RU,E=info@cryptopro.ru", resultObject.IssuerName );
            assertEquals( "Serial number of certificate", "698403286AA659BA4635622D49DE5FD3", resultObject.SerialNumber );
        }//4 asserts

        //GettingCertProperties        
        {
            var strResult = GettingCertProperties( this.inCertName, false );
            assertString( "GettingCertProperties function result", strResult );

            //regexp assert fails :(
            //assertMatch( "Issuer name", ".*", strResult );
            {
                var regExp = new RegExp( ".*IssuerName: CN=Test Center CRYPTO-PRO,O=CRYPTO-PRO,C=RU,E=info@cryptopro.ru.*" );
                assertTrue( "Issuer name", regExp.test(strResult) );
            }
            //assertMatch( "Serial number", ".*", strResult );
            {
                var regExp = new RegExp( ".*SerialNumber: 698403286AA659BA4635622D49DE5FD3.*" );
                assertTrue( "Serial number", regExp.test(strResult) );
            }
            //assertMatch( "Public key alg", ".*", strResult );
            {
                var regExp = new RegExp( ".*PublicKeyAlg: 1.2.643.2.2.19.*" );
                assertTrue( "Public key alg", regExp.test(strResult) );
            }

            //TODO: Add testing certificate with "sofisticated" DN
        }//4 asserts

        //SettingSerialAndIssuerToCert
        {
            var strIssuer = "CN=John Doe, E=johndoe@example.com";
            var strSerial = "Serial-number";

            var resultCert = SettingSerialAndIssuerToCert( strSerial, strIssuer );

            assertObject( "Returned certificate", resultCert );
            assertEquals( "Issuer name", strIssuer, resultCert.IssuerName );
            assertEquals( "Serial number", strSerial, resultCert.SerialNumber );
        }//3 asserts

        //CompareCertificates
        {
            assertFalse( "Compare result of unequal certificates", CompareCertificates( this.inCertName, false ) );
            assertTrue( "Compare result of equal certificates", CompareCertificates( this.certForComparing, false ) );
        }//2 asserts

        //VerifyCertificateStatus
        {
            var strCertStatus = VerifyCertificateStatus(false);
            assertString( "VerifyCertificateStatus function returned value", strCertStatus );
            assertEquals( "Certificate status string value", "Status: Корректен", strCertStatus );

            //TODO: Add certificates with different status (uncorrect, without CRL, etc.)
        }//2 asserts

        //jstestdriver.console.log("Done", "");
    } //,

// Testing of the CryptoAPI example is moved into a separate function and commented out,
// because this property is obsolete and causes errors

//    testCryptoApiContext:function()
//    {
//        expectAsserts(4);
//
//        //SettingCryptoApiContext
//        {
//            var resultCert = SettingCryptoApiContext( this.inCertName, false );
//            assertObject( "Returning value of SettingCryptoApiContext function", resultCert );
//            assertNotEquals( "Context type", 0, resultCert.CertContext );
//
//            assertEquals( "Issuer name of certificate", "CN=Test Center CRYPTO-PRO,O=CRYPTO-PRO,C=RU,E=info@cryptopro.ru", resultCert.IssuerName );
//            jstestdriver.console.log("Issuer:", resultCert.IssuerName);
//            assertEquals( "Serial number of certificate", "698403286AA659BA4635622D49DE5FD3", resultCert.SerialNumber );
//            jstestdriver.console.log("Serial:", resultCert.SerialNumber);
//        }//4 assert
//
//        //jstestdriver.console.log("Done", "");
//    }
});

