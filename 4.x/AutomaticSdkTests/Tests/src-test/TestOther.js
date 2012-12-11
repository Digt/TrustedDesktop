
// To prepare test stand:
//  - For new CryptoARM release update version in the testing code part of GetDigtCryptoVersion function

TestCase("Other",{
    //stored values
    strInstalledLicense:"",

    tearDown:function()
    {
        if( this.strInstalledLicense != "" )
        {
            jstestdriver.console.log("", "");
            jstestdriver.console.log("It seems that license \"" + this.profileId + "\" is removed from registry", "");
            jstestdriver.console.log("Trying to restore...", "");

            try
            {
                var oUtil = new ActiveXObject("DigtCrypto.Util");
                oUtil.License = this.strInstalledLicense;

                jstestdriver.console.log("\t", "License successfully restored.");
            }
            catch(err)
            {
                jstestdriver.console.log("\t", "Failed to restore license. You must to do it manually.");
                jstestdriver.console.log("\t", String(err));
            }
        }
    },

    testOther:function()
    {
        expectAsserts(8);

        //GetDigtCryptoVersion
        {
            var sVersion = GetDigtCryptoVersion( false );
            assertString( "Returned value of the GetDigtCryptoVersion function", sVersion );
            var arrVersionparts = sVersion.split( ".", 3 );
            assertEquals( "Version, returned by GetDigtCryptoVersion function", "4.7.0", arrVersionparts.join(".") );
        }//2 assert

        //SetUpLicense
        {
            this.strInstalledLicense = SetUpLicense( "TD4AB-YRVAL-GABYR-VALGA-BYRVA-LGABY-RVALG", false );
            assertString( "Returned license value", this.strInstalledLicense );
            assertNotEquals( "Returned license value", this.strInstalledLicense, "" );
            assertEquals( "Product id from license", "TD4", this.strInstalledLicense.substr(0,3) );

            assertEquals( "Returned pseudo license", "TD4AB-YRVAL-GABYR-VALGA-BYRVA-LGABY-RVALG", SetUpLicense( this.strInstalledLicense, false ) );
            this.strInstalledLicense = "";
        }//4 assert

        //StartUpAndShutDownLogger
        {
            var sLogMessages = StartUpAndShutDownLogger(false);
            assertString( "StartUpAndShutDownLogger function returned value", sLogMessages );
            assertNotEquals( "Count of returned providers", String(sLogMessages).length, 0 );
        }//2 asserts

        //jstestdriver.console.log("Done", "");
    }
});

