
// To prepare test stand:
//  - Default profile must exists in the profile store
//  - Parameters of signing must be setted
//  - Parameters of encryption must be setted

TestCase("Signatures",{
    // Objects
    fso: null,

    // Files
    inputSign:"",

    // Stored constants
    strTmpFolder:"",
    signContent:"----- BEGIN PKCS7 SIGNED -----\nMIAGCSqGSIb3DQEHAqCAMIACAQExCzAJBgUrDgMCGgUAMIAGCSqGSIb3DQEHAaCA\nJIAEgfYNCjxodG1sPg0KPGhlYWQ+DQo8c2NyaXB0IHNyYz0iUmVxdWVzdHMxLmpz\nIj48L3NjcmlwdD4NCjwvaGVhZD4NCjxib2R5Pg0KPGJ1dHRvbiBvbmNsaWNrPSJy\ndW5UZXN0KCk7Ij5QdXNoPC9idXR0b24+DQoNCjwhLS0gZm9ybSBhY3Rpb249IkRp\nc3BsYXlNZXNzYWdlKCdCdXR0b24gcHJlc3NlZCcpOyI+DQogICAgPGlucHV0IHR5\ncGU9ImJ1dHRvbiIgdmFsdWU9IlB1c2giPg0KPC9mb3JtIC0tPg0KDQo8L2JvZHk+\nDQo8L2h0bWw+DQoAAAAAAACgggfGMIIB/TCCAaegAwIBAgIIXCDaXRYZRu4wDQYJ\nKoZIhvcNAQEFBQAwMTELMAkGA1UEBhMCUlUxETAPBgNVBAoTCHRlc3Qgb3JnMQ8w\nDQYDVQQDEwZ0ZXN0MDEwHhcNMTIwNTEyMDgxMTA0WhcNMTMwNTEyMDgxMTA0WjAx\nMQswCQYDVQQGEwJSVTERMA8GA1UEChMIdGVzdCBvcmcxDzANBgNVBAMTBnRlc3Qw\nMTBcMA0GCSqGSIb3DQEBAQUAA0sAMEgCQQDC15AkG43uIFLdXjn0sI8XP0kTaYBH\nsdGA3hyTAd9f1Cto5p0ka9tHtS/XxPgIfoIj0v4xPNRf2URyqfqXV1tlAgMBAAGj\ngaIwgZ8wHQYDVR0OBBYEFGyl/tabOvdDeMX8ulDkeXe4K29NMAsGA1UdDwQEAwID\n2DAPBgNVHRMECDAGAQH/AgEBMGAGA1UdIwRZMFeAFGyl/tabOvdDeMX8ulDkeXe4\nK29NoTWkMzAxMQswCQYDVQQGEwJSVTERMA8GA1UEChMIdGVzdCBvcmcxDzANBgNV\nBAMTBnRlc3QwMYIIXCDaXRYZRu4wDQYJKoZIhvcNAQEFBQADQQAMji+vwBk/RHyc\nGbeoY+KZgUnr4QQviqEiN9zar4aCtGO5h9Cpin3ANvXaoDbppSpnd7h2l9O+zA6X\nxwn+MBP1MIICljCCAkCgAwIBAgIIXAkQk1va8jkwDQYJKoZIhvcNAQEFBQAwYzEi\nMCAGCSqGSIb3DQEJARYTdGVzdG9yZ0BleGFtcGxlLmNvbTELMAkGA1UEBhMCUlUx\n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\n----- END PKCS7 SIGNED -----\n",

    setUp:function()
    {
        // Init file system object
        this.fso = new ActiveXObject("Scripting.FileSystemObject");

        // Init temporary folder
        this.strTmpFolder = this.fso.GetSpecialFolder(2) + "\\";
        
        // Init files names and paths
        this.inputSign = this.strTmpFolder + "test.txt.sig";

        // Save signature file
        var signFile = this.fso.CreateTextFile( this.inputSign );
        signFile.Write(this.signContent);
        signFile.Close();
    },

    tearDown:function()
    {
        this.fso.DeleteFile( this.inputSign );
    },

    testSignatures:function()
    {
        expectAsserts(12);

        //SignatureStatusVerification
        {
            var arrStatuses = SignatureStatusVerification(this.inputSign, false);

            assertArray( "SignatureStatusVerification returned value", arrStatuses );

            assertEquals( "Number of verifyed signatures", 2, arrStatuses.length );
            assertEquals( "First signature status", 1, arrStatuses[0] );
            assertEquals( "Second signature status", 1, arrStatuses[1] );
        }//4 asserts

        //GettingSignaturesProperties
        {
            var strResult = GettingSignaturesProperties(this.inputSign, false);
            assertString( "GettingSignaturesProperties returned value", strResult );
            assertNotEquals( "GettingSignaturesProperties returned string length", 0, strResult.length );

            //regexp assert fails :(
            //assertMatch( "", ".*", strResult );
            {
                var regExp = new RegExp( ".*CN=User User.*" );
                assertTrue( "First signature certificate owner", regExp.test(strResult) );
            }
            //assertMatch( "", ".*", strResult );
            {
                var regExp = new RegExp( ".*CN=RSA MSCSP Certificate.*" );
                assertTrue( "Second signature certificate owner", regExp.test(strResult) );
            }
        }//4 asserts

        //GettingCollectionOfCosigns
        {
            var strResult = GettingCollectionOfCosigns(this.inputSign, false);
            assertString( "GettingCollectionOfCosigns returned value", strResult );
            assertNotEquals( "GettingCollectionOfCosigns returned string length", 0, strResult );

            //regexp assert fails :(
            //assertMatch( "", ".*", strResult );
            {
                var regExp = new RegExp( ".*#1 = 0.*" );
                assertTrue( "First signature cosign count", regExp.test(strResult) );
            }
            //assertMatch( "", ".*", strResult );
            {
                var regExp = new RegExp( ".*#2 = 1.*" );
                assertTrue( "Second signature cosign count", regExp.test(strResult) );
            }

            //TODO: Check propreties of cosigns
        }//4 asserts

        //jstestdriver.console.log("Done", "");
    }
});

TestCase("CryptoOperations",{
    // Objects
    fso: null,

    // Files
    inPlainFile:"",
    outSignFile:"",
    outSignWithAddedSign:"",
    outSignWithCosign:"",
    outEncryptedFile:"",
    outDecryptedFile:"",

    // Stored constants
    strTmpFolder:"",

    setUp:function()
    {
        // Init file system object
        this.fso = new ActiveXObject("Scripting.FileSystemObject");

        // Init temporary folder
        this.strTmpFolder = this.fso.GetSpecialFolder(2) + "\\";

        // Init files names and paths
        this.inPlainFile = this.strTmpFolder + "test.txt";
        this.outSignFile = this.strTmpFolder + "test.txt.sig";
        this.outSignWithAddedSign = this.strTmpFolder + "test-with-added-sign.txt.sig";
        this.outSignWithCosign = this.strTmpFolder + "test-with-cosig.txt.sig";
        this.outEncryptedFile = this.strTmpFolder + "test.txt.enc";
        this.outDecryptedFile = this.strTmpFolder + "test-decrypted.txt";

        // Save plain file
        var plainFile = this.fso.CreateTextFile( this.inPlainFile );
        plainFile.Write("Plain file content");
        plainFile.Close();
    },

    tearDown:function()
    {
        this.fso.DeleteFile( this.inPlainFile );

        if( this.fso.FileExists( this.outSignFile ) )
        {
            this.fso.DeleteFile( this.outSignFile );
        }

        if( this.fso.FileExists( this.outSignWithAddedSign ) )
        {
            this.fso.DeleteFile( this.outSignWithAddedSign );
        }

        if( this.fso.FileExists( this.outSignWithCosign ) )
        {
            this.fso.DeleteFile( this.outSignWithCosign );
        }

        if( this.fso.FileExists( this.outEncryptedFile ) )
        {
            this.fso.DeleteFile( this.outEncryptedFile );
        }

        if( this.fso.FileExists( this.outDecryptedFile ) )
        {
            this.fso.DeleteFile( this.outDecryptedFile );
        }
    },

    testCryptoOperations:function()
    {
        expectAsserts(26);

        //SignatureCreation
        {
            var bSignResult = SignatureCreation(this.inPlainFile, this.outSignFile, false);
            assertTrue( "SignatureCreation function result", bSignResult );

            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.outSignFile));

            var oNewMessage = new ActiveXObject("DigtCrypto.PKCS7Message");
            assertTrue( "SignatureCreation exit file loading result", oNewMessage.Load( 2, this.outSignFile ) );
            assertObject( "Signatures collection type", oNewMessage.Signatures );
            assertEquals( "Signatures count", 1, oNewMessage.Signatures.Count );
            assertEquals( "First sign status", 1, oNewMessage.Signatures.Item(0).Verify(0,"") );
        }//6 asserts

        //AddSignature
        {
            var bResult = AddSignature(this.outSignFile, this.outSignWithAddedSign, false);
            assertTrue( "AddSignature function result", bResult );

            assertTrue( "AddSignature result file exists", this.fso.FileExists(this.outSignWithAddedSign));

            var oNewMessage = new ActiveXObject("DigtCrypto.PKCS7Message");
            assertTrue( "AddSignature exit file loading result", oNewMessage.Load( 2, this.outSignWithAddedSign ) );
            assertObject( "Signatures collection type", oNewMessage.Signatures );
            assertEquals( "Signatures count", oNewMessage.Signatures.Count, 2 );
            assertEquals( "First sign status", 1, oNewMessage.Signatures.Item(0).Verify(0,"") );
            assertEquals( "Second sign status", 1, oNewMessage.Signatures.Item(1).Verify(0,"") );
        }//7 asserts

        //AddCosignature
        {
            var bResult = AddCosignature(this.outSignWithAddedSign, this.outSignWithCosign, false);
            assertTrue( "AddCosignature signatureCreation function result", bResult );

            assertTrue( "AddCosignature result file exists", this.fso.FileExists(this.outSignWithCosign));

            var oNewMessage = new ActiveXObject("DigtCrypto.PKCS7Message");
            assertTrue( "AddCosignature exit file loading result", oNewMessage.Load( 2, this.outSignWithCosign ) );
            assertObject( "Signatures collection type", oNewMessage.Signatures );
            assertEquals( "Signatures count", 2, oNewMessage.Signatures.Count );

            var signWithCosign;
            var cosignCount = 0;
            for( var i = 0; i < oNewMessage.Signatures.Count; i++ )
            {
                var count = oNewMessage.Signatures.Item(i).Cosignature.Count;
                if( count )
                {
                    cosignCount += count;
                    signWithCosign = i;
                }
            }

            assertEquals( "Total number of cosigns", 1, cosignCount );
            assertEquals( "Cosign status", 1, oNewMessage.Signatures.Item(signWithCosign).Cosignature.Item(0).Verify(0,"") );
        }//7 asserts

        //EncryptData
        {
            var bResult = EncryptData(this.inPlainFile, this.outEncryptedFile, false);
            assertTrue( "EncryptData function result", bResult );

            assertTrue( "EncryptData result file exists", this.fso.FileExists(this.outEncryptedFile));

            var oNewMessage = new ActiveXObject("DigtCrypto.PKCS7Message");
            oNewMessage.Load( -1, this.outEncryptedFile );
            assertEquals( "Data type fo a loaded exit file of a EncryptData function", 3, oNewMessage.ContentType );

            //releasing ecnrypted file, otherwise test can finished with error
            //      (permission on encrypted file is denied)
            oNewMessage.Import( 0, "anything", "" );
            oNewMessage = null;
        }//3 asserts

        //DecryptData
        {
            var bResult = DecryptData(this.outEncryptedFile, this.outDecryptedFile, false);
            assertTrue( "DecryptData function result", bResult );

            assertTrue( "DecryptData result file exists", this.fso.FileExists(this.outDecryptedFile));

            var file = this.fso.OpenTextFile(this.inPlainFile, 1, false);
            var sourceFileData = String(file.ReadAll());
            file.Close();

            file = this.fso.OpenTextFile(this.outDecryptedFile, 1, false);
            var decryptedFileData = String(file.ReadAll());
            file.Close();

            assertTrue( "Result of comparing source and decrypted data", decryptedFileData == sourceFileData );
        }//3 asserts

        //jstestdriver.console.log("Done", "");
    }
});

/*
TestCase("CryptoOperationByBlocks",{
    // Objects
    fso: null,

    // Files
    // Encrypt from memory
    file01WithAttachedHeader:"",
    file02WithAttachedHeaderBlock2:"",

    file03DetachedHeader:"",
    file04DetachedHeaderBlock1:"",
    file05DetachedHeaderBlock2:"",

    // Encrypt from file
    sourceData1FileName:"",
    sourceData2FileName:"",

    file06WithAttachedHeader:"",
    file07WithAttachedHeaderBlock2:"",

    file08DetachedHeader:"",
    file09DetachedHeaderBlock1:"",
    file10DetachedHeaderBlock2:"",

    // Decrypt result
    outFile1Block1:"",
    outFile1Block2:"",
    outFile2Block1:"",
    outFile2Block2:"",
    outFile3Block1:"",
    outFile3Block2:"",
    outFile4Block1:"",
    outFile4Block2:"",


    // Stored constants
    strTmpFolder:"",

//    data1BlockWithAttachedHeader:"",
//    data2Block2:"",
//    data3DetachedHeader:"",
//    data4Block1WithDetachedHeader:"",
//    data5Block2:"",

    setUp:function()
    {
        // Init file system object
        this.fso = new ActiveXObject("Scripting.FileSystemObject");

        // Init temporary folder
        this.strTmpFolder = this.fso.GetSpecialFolder(2) + "\\";

        // Init files names and paths
        this.file01WithAttachedHeader = this.strTmpFolder + this.fso.GetTempName();
        this.file02WithAttachedHeaderBlock2 = this.strTmpFolder + this.fso.GetTempName();
        this.file03DetachedHeader = this.strTmpFolder + this.fso.GetTempName();
        this.file04DetachedHeaderBlock1 = this.strTmpFolder + this.fso.GetTempName();
        this.file05DetachedHeaderBlock2 = this.strTmpFolder + this.fso.GetTempName();
        this.sourceData1FileName = this.strTmpFolder + this.fso.GetTempName();
        this.sourceData2FileName = this.strTmpFolder + this.fso.GetTempName();
        this.file06WithAttachedHeader = this.strTmpFolder + this.fso.GetTempName();
        this.file07WithAttachedHeaderBlock2 = this.strTmpFolder + this.fso.GetTempName();
        this.file08DetachedHeader = this.strTmpFolder + this.fso.GetTempName();
        this.file09DetachedHeaderBlock1 = this.strTmpFolder + this.fso.GetTempName();
        this.file10DetachedHeaderBlock2 = this.strTmpFolder + this.fso.GetTempName();
        this.outFile1Block1 = this.strTmpFolder + this.fso.GetTempName();
        this.outFile1Block2 = this.strTmpFolder + this.fso.GetTempName();
        this.outFile2Block1 = this.strTmpFolder + this.fso.GetTempName();
        this.outFile2Block2 = this.strTmpFolder + this.fso.GetTempName();
        this.outFile3Block1 = this.strTmpFolder + this.fso.GetTempName();
        this.outFile3Block2 = this.strTmpFolder + this.fso.GetTempName();
        this.outFile4Block1 = this.strTmpFolder + this.fso.GetTempName();
        this.outFile4Block2 = this.strTmpFolder + this.fso.GetTempName();

        // Save plain file
        var plainFile = this.fso.CreateTextFile( this.sourceData1FileName );
        plainFile.Write("Data block 1 content");
        plainFile.Close();
        
        plainFile = this.fso.CreateTextFile( this.sourceData2FileName );
        plainFile.Write("Content of data block 2");
        plainFile.Close();
    },

    tearDown:function()
    {
        if( this.fso.FileExists( this.file01WithAttachedHeader ) )
        {
            this.fso.DeleteFile( this.file01WithAttachedHeader );
        }
        if( this.fso.FileExists( this.file02WithAttachedHeaderBlock2 ) )
        {
            this.fso.DeleteFile( this.file02WithAttachedHeaderBlock2 );
        }
        if( this.fso.FileExists( this.file03DetachedHeader ) )
        {
            this.fso.DeleteFile( this.file03DetachedHeader );
        }
        if( this.fso.FileExists( this.file04DetachedHeaderBlock1 ) )
        {
            this.fso.DeleteFile( this.file04DetachedHeaderBlock1 );
        }
        if( this.fso.FileExists( this.file05DetachedHeaderBlock2 ) )
        {
            this.fso.DeleteFile( this.file05DetachedHeaderBlock2 );
        }

        this.fso.DeleteFile( this.sourceData1FileName );
        this.fso.DeleteFile( this.sourceData2FileName );

        if( this.fso.FileExists( this.file06WithAttachedHeader ) )
        {
            this.fso.DeleteFile( this.file06WithAttachedHeader );
        }
        if( this.fso.FileExists( this.file07WithAttachedHeaderBlock2 ) )
        {
            this.fso.DeleteFile( this.file07WithAttachedHeaderBlock2 );
        }
        if( this.fso.FileExists( this.file08DetachedHeader ) )
        {
            this.fso.DeleteFile( this.file08DetachedHeader );
        }
        if( this.fso.FileExists( this.file09DetachedHeaderBlock1 ) )
        {
            this.fso.DeleteFile( this.file09DetachedHeaderBlock1 );
        }
        if( this.fso.FileExists( this.file10DetachedHeaderBlock2 ) )
        {
            this.fso.DeleteFile( this.file10DetachedHeaderBlock2 );
        }
        if( this.fso.FileExists( this.outFile1Block1 ) )
        {
            this.fso.DeleteFile( this.outFile1Block1 );
        }
        if( this.fso.FileExists( this.outFile1Block2 ) )
        {
            this.fso.DeleteFile( this.outFile1Block2 );
        }
        if( this.fso.FileExists( this.outFile2Block1 ) )
        {
            this.fso.DeleteFile( this.outFile2Block1 );
        }
        if( this.fso.FileExists( this.outFile2Block2 ) )
        {
            this.fso.DeleteFile( this.outFile2Block2 );
        }
        if( this.fso.FileExists( this.outFile3Block1 ) )
        {
            this.fso.DeleteFile( this.outFile3Block1 );
        }
        if( this.fso.FileExists( this.outFile3Block2 ) )
        {
            this.fso.DeleteFile( this.outFile3Block2 );
        }
        if( this.fso.FileExists( this.outFile4Block1 ) )
        {
            this.fso.DeleteFile( this.outFile4Block1 );
        }
        if( this.fso.FileExists( this.outFile4Block2 ) )
        {
            this.fso.DeleteFile( this.outFile4Block2 );
        }
    },

    testCryptoOperationByBlocks:function()
    {
        expectAsserts(26);

        var file;
        var readedData;

        //EncryptByBlocksFromMemAttachedHeader
        {
            EncryptByBlocksFromMemAttachedHeader( this.file01WithAttachedHeader, this.file02WithAttachedHeaderBlock2 );

            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.file01WithAttachedHeader));
            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.file02WithAttachedHeader));
        }//2 asserts

        //EncryptByBlocksFromMemDetachedHeader
        {
            EncryptByBlocksFromMemDetachedHeader( this.file03DetachedHeader, this.file04DetachedHeaderBlock1, this.file05DetachedHeaderBlock2 );

            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.file03DetachedHeader));
            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.file04DetachedHeaderBlock1));
            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.file05DetachedHeaderBlock2));
        }//3 asserts

        //EncryptByBlocksFromFileAttachedHeader
        {
            EncryptByBlocksFromFileAttachedHeader( this.sourceData1FileName, this.sourceData2FileName, this.file06WithAttachedHeader, this.file07WithAttachedHeaderBlock2 );

            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.file06WithAttachedHeader));
            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.file07WithAttachedHeaderBlock2));
        }//2 asserts

        //EncryptByBlocksFromFileDetachedHeader
        {
            EncryptByBlocksFromFileDetachedHeader( this.sourceData1FileName, this.sourceData2FileName, this.file08DetachedHeader, this.file09DetachedHeaderBlock1, this.file10DetachedHeaderBlock2 );

            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.file08DetachedHeader));
            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.file09DetachedHeaderBlock1));
            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.file10DetachedHeaderBlock2));
        }//3 asserts


        //DecryptByBlocksFromMemAttachedHeader
        {
            DecryptByBlocksFromMemAttachedHeader( this.file01WithAttachedHeader, this.file02WithAttachedHeaderBlock2, this.outFile1Block1, this.outFile1Block2 );

            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.outFile1Block1));
            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.outFile1Block2));

            file = fso.OpenTextFile( this.outFile1Block1, 1);
            readedData = file.ReadAll();
            file.Close();
            assertEquals( "First decrypted block (1)", "Первый блок данных", readedData );

            file = fso.OpenTextFile( this.outFile1Block2, 1);
            readedData = file.ReadAll();
            file.Close();
            assertEquals( "Second decrypted block (1)", "Второй блок данных", readedData );
        }//4 asserts

        //DecryptByBlocksFromMemDetachedHeader
        {
            DecryptByBlocksFromMemDetachedHeader( this.file03DetachedHeader, this.file04DetachedHeaderBlock1, this.file05DetachedHeaderBlock2, this.outFile2Block1, this.outFile2Block2 );

            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.outFile2Block1));
            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.outFile2Block2));

            file = fso.OpenTextFile( this.outFile2Block1, 1);
            readedData = file.ReadAll();
            file.Close();
            assertEquals( "First decrypted block (2)", "Первый блок данных", readedData );

            file = fso.OpenTextFile( this.outFile2Block2, 1);
            readedData = file.ReadAll();
            file.Close();
            assertEquals( "Second decrypted block (2)", "Второй блок данных", readedData );
        }//4 asserts

        //DecryptByBlocksFromFileAttachedHeader
        {
            DecryptByBlocksFromFileAttachedHeader( this.file06WithAttachedHeader, this.file07WithAttachedHeaderBlock2, this.outFile3Block1, this.outFile3Block2 );

            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.outFile3Block1));
            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.outFile3Block2));

            file = fso.OpenTextFile( this.outFile3Block1, 1);
            readedData = file.ReadAll();
            file.Close();
            assertEquals( "First decrypted block (3)", "Data block 1 content", readedData );

            file = fso.OpenTextFile( this.outFile3Block2, 1);
            readedData = file.ReadAll();
            file.Close();
            assertEquals( "Second decrypted block (3)", "Content of data block 2", readedData );
        }//4 asserts

        //DecryptByBlocksFromFileDetachedHeader
        {
            DecryptByBlocksFromFileDetachedHeader( this.file08DetachedHeader, this.file09DetachedHeaderBlock1, this.file10DetachedHeaderBlock2, this.outFile4Block1, this.outFile4Block2 );

            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.outFile4Block1));
            assertTrue( "SignatureCreation result file exists", this.fso.FileExists(this.outFile4Block2));

            file = fso.OpenTextFile( this.outFile4Block1, 1);
            readedData = file.ReadAll();
            file.Close();
            assertEquals( "First decrypted block (4)", "Data block 1 content", readedData );

            file = fso.OpenTextFile( this.outFile4Block2, 1);
            readedData = file.ReadAll();
            file.Close();
            assertEquals( "Second decrypted block (4)", "Content of data block 2", readedData );
        }//4 asserts
    }
});
*/
