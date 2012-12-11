
TestCase("Cryptoproviders",{
    testCryptoproviders:function()
    {
        expectAsserts(4);

        //GetSupportedProvidersList
        {
            var arrProvidersList = GetSupportedProvidersList(false);
            assertNotEquals( "Count of returned providers", 0, arrProvidersList.length );

            //regexp assert fails :(
            //assertMatch( "List contains MS CSP", ".*Microsoft Base Cryptographic Provider.*", arrProvidersList.join(" ") );
            {
                var regExp = new RegExp( ".*Microsoft Base Cryptographic Provider.*" );
                assertTrue( "List contains MS CSP", regExp.test(arrProvidersList.join(" ")) );
            }
        }//2 asserts

        //GetSystemCryptoProviders
        {
            var arrProvidersList = GetSystemCryptoProviders(false);
            assertNotEquals( "Count of returned providers", 0, arrProvidersList.length );

            //regexp assert fails :(
            //assertMatch( "List contains MS CSP", ".*Microsoft Base Cryptographic Provider.*", arrProvidersList.join(" ") );
            {
                var regExp = new RegExp( ".*Microsoft Base Cryptographic Provider.*" );
                assertTrue( "List contains MS CSP", regExp.test(arrProvidersList.join(" ")) );
            }
        }//2 asserts

        //jstestdriver.console.log("Done", "");
    }
});

TestCase("Profiles",{
    // Objects
    fso: null,

    // Files
    inCertificateForProfile:"",

    // Stored constants and values
    strTmpFolder:"",
    profileId:"",

    certForProfileContent:"-----BEGIN CERTIFICATE-----\nMIICljCCAkCgAwIBAgIIXAkQk1va8jkwDQYJKoZIhvcNAQEFBQAwYzEiMCAGCSqG\nSIb3DQEJARYTdGVzdG9yZ0BleGFtcGxlLmNvbTELMAkGA1UEBhMCUlUxEDAOBgNV\nBAoTB3Rlc3RvcmcxHjAcBgNVBAMTFVJTQSBNU0NTUCBDZXJ0aWZpY2F0ZTAeFw0x\nMjA4MzAxMzI5NDFaFw0xMzA4MzAxMzI5NDFaMGMxIjAgBgkqhkiG9w0BCQEWE3Rl\nc3RvcmdAZXhhbXBsZS5jb20xCzAJBgNVBAYTAlJVMRAwDgYDVQQKEwd0ZXN0b3Jn\nMR4wHAYDVQQDExVSU0EgTVNDU1AgQ2VydGlmaWNhdGUwXDANBgkqhkiG9w0BAQEF\nAANLADBIAkEA3TvNszBchezHHt8+qjTaKLq6s6JQZu23Qcq+nGgokd/YhAvAkV8n\nKi4fwOfb+NPXcslXYGd4ORjit86f7x9zewIDAQABo4HXMIHUMB0GA1UdDgQWBBR7\nxMi48uabg/dHiF0Lcb69xEIvgjALBgNVHQ8EBAMCAtwwDwYDVR0TBAgwBgEB/wIB\nATCBlAYDVR0jBIGMMIGJgBR7xMi48uabg/dHiF0Lcb69xEIvgqFnpGUwYzEiMCAG\nCSqGSIb3DQEJARYTdGVzdG9yZ0BleGFtcGxlLmNvbTELMAkGA1UEBhMCUlUxEDAO\nBgNVBAoTB3Rlc3RvcmcxHjAcBgNVBAMTFVJTQSBNU0NTUCBDZXJ0aWZpY2F0ZYII\nXAkQk1va8jkwDQYJKoZIhvcNAQEFBQADQQBnxIQGpVje2V0e+k40HNTCc23IBwWZ\nM18X7FpdzeYpbNr1/wLhJkOSHssWAJqwrElFsJSEA6zpqM8atTpUpAg4\n-----END CERTIFICATE-----\n",

    setUp:function()
    {
        // Init file system object
        this.fso = new ActiveXObject("Scripting.FileSystemObject");

        // Init temporary folder
        this.strTmpFolder = this.fso.GetSpecialFolder(2) + "\\";

        // Init files names and paths
        this.inCertificateForProfile = this.strTmpFolder + "certForProfile.cer";

        // Save template file
        var tmpltFile = this.fso.CreateTextFile( this.inCertificateForProfile );
        tmpltFile.Write(this.certForProfileContent);
        tmpltFile.Close();
    },

    tearDown:function()
    {
        this.fso.DeleteFile( this.inCertificateForProfile );

        if( this.profileId != "" )
        {
            jstestdriver.console.log("", "");
            jstestdriver.console.log("Test profile with ID \"" + this.profileId + "\" is not removed from store.", "");
            jstestdriver.console.log("Trying to ermove...", "");

            var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
            objProfileStore.Open(0);

            //var objNewProfile = null;

            try
            {
                for( var i = 0; i < objProfileStore.Store.Count; i++ )
                {
                    if( objProfileStore.Store.Item(i).ID == this.profileId )
                    {
                        objProfileStore.Store.Remove( i );
                        break;
                    }
                }
                objProfileStore.Close();

                jstestdriver.console.log("\t", "Profile successfully removed");
            }
            catch(err)
            {
                jstestdriver.console.log("\t", "Error while removing profile");
                jstestdriver.console.log("\t", String(err));
            }
        }
    },

    testProfiles:function()
    {
        expectAsserts(9);

        //CreateNewProfile
        {
            var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
            objProfileStore.Open(0);
            var profilesCount = objProfileStore.Store.Count;
            objProfileStore.Close();

            this.profileId = CreateNewProfile( this.inCertificateForProfile, false );

            objProfileStore.Open(0);
            assertEquals( "Increased profile count after function CreateNewProfile", profilesCount+1, objProfileStore.Store.Count );

            var objNewProfile = null;
            for( var i = 0; i < objProfileStore.Store.Count; i++ )
            {
                if( objProfileStore.Store.Item(i).ID == this.profileId )
                {
                    objNewProfile = objProfileStore.Store.Item(i);
                }
            }
            objProfileStore.Close();

            assertObject( "New found profile", objNewProfile );

            assertEquals( "Profile name", "Test profile", objNewProfile.Name );
            assertEquals( "Profile description", "Remove this test profile", objNewProfile.Description );
        }//4 asserts

        //GetProfileParameters
        {
            var strProfileProperties = GetProfileParameters( this.profileId, false );
            //jstestdriver.console.log("Profile properties:\n\n", strProfileProperties);

            //regexp assert fails :(
            //assertMatch( "Profile ID matching", ".*ID=" + this.profileId + ".*", strProfileProperties );
            {
                var regExp = new RegExp( "[.\n]*ID=" + this.profileId + "[.\n]*" );
                assertTrue( "Profile ID matching", regExp.test(strProfileProperties) );
            }
            //assertMatch( "Profile name matching", "[.\n]*Name=Test profile[.\n]*", strProfileProperties );
            {
                var regExp = new RegExp( "[.\n]*Name=Test profile[.\n]*" );
                assertTrue( "Profile name matching", regExp.test(strProfileProperties) );
            }
            //assertMatch( "Profile description matching", "[.\n]*Description=Remove this test profile[.\n]*", strProfileProperties );
            {
                var regExp = new RegExp( "[.\n]*Description=Remove this test profile[.\n]*" );
                assertTrue( "Profile description matching", regExp.test(strProfileProperties) );
            }
        }//3 asserts

        //RemovingProfile
        {
            var objProfileStore = new ActiveXObject("DigtCrypto.ProfileStore");
            objProfileStore.Open(0);
            var profilesCount = objProfileStore.Store.Count;
            objProfileStore.Close();

            assertTrue( "Remove profile result", RemovingProfile( this.profileId ) );

            objProfileStore.Open(0);
            assertEquals( "Decreased profile count", profilesCount-1, objProfileStore.Store.Count );
            objProfileStore.Close();

            this.profileId = "";
        }//2 asserts

        //jstestdriver.console.log("Done", "");
    }
});

