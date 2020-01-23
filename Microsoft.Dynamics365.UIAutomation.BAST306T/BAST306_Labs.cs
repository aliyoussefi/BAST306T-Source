// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.ApplicationInsights.DataContracts;
using System.Diagnostics;

namespace Microsoft.Dynamics365.UIAutomation.BAST306T
{
    [TestClass]
    public class BAST306_Labs
    {
        private static TestContext _context;
        private static string _username = "";
        private static string _password = "";
        private static BrowserType _browserType;
        private static Uri _xrmUri;
        private static string _azureKey = "";

        [TestMethod]
        [TestCategory("Pass")]
        public void SharedTestUploadTelemetryUCI()
        {
            //Setting the options here to demonstrate what needs to 
            //be set for tracking performance telemetry in UCI
            var options = TestSettings.Options;
            options.AppInsightsKey = _azureKey;
            options.UCIPerformanceMode = true;


            var client = new WebClient(options);
            using (var xrmApp = new XrmApp(client))
            {
                try
                {
                    xrmApp.OnlineLogin.Login(_xrmUri, _username.ToSecureString(), _password.ToSecureString());

                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                    //Track Performance Center Telemetry
                    //----------------------------------------------------------------
                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    xrmApp.Telemetry.TrackPerformanceEvents();
                    //----------------------------------------------------------------


                    //Track Performance Center Telemetry with additional data
                    //Add additional information to telemtry to specify which entity was opened
                    //----------------------------------------------------------------
                    xrmApp.Grid.OpenRecord(0);

                    var props = new Dictionary<string, string>();
                    props.Add("Entity", "Account");

                    var metrics = new Dictionary<string, double>();
                    var commandResult = xrmApp.CommandResults.FirstOrDefault(x => x.CommandName == "Open Grid Record");

                    metrics.Add(commandResult.CommandName, commandResult.ExecutionTime);

                    //xrmApp.Telemetry.TrackPerformanceEvents(props, metrics);
                    //----------------------------------------------------------------

                    //Track Command Results Events 
                    //----------------------------------------------------------------
                    //xrmApp.Telemetry.TrackCommandEvents();
                    //----------------------------------------------------------------

                    //Track Browser Performance Telemetry
                    //----------------------------------------------------------------
                   // xrmApp.Telemetry.TrackBrowserEvents(Api.UCI.Telemetry.BrowserEventType.Resource, null, true);
                    //xrmApp.Telemetry.TrackBrowserEvents(Api.UCI.Telemetry.BrowserEventType.Navigation);
                    //xrmApp.Telemetry.TrackBrowserEvents(Api.UCI.Telemetry.BrowserEventType.Server);
                    //----------------------------------------------------------------

                }
                catch (Exception)
                {
                    //client.Browser.TakeWindowScreenShot(Assembly.GetExecutingAssembly().Location, OpenQA.Selenium.ScreenshotImageFormat.Png);
                    throw;
                }
               
            }
        }

        #region ctor
        [ClassInitialize]
        public static void LabConstructor(TestContext context)
        {
            _context = context;
            _username = context.Properties["OnlineUsername"].ToString();
            _password = context.Properties["OnlinePassword"].ToString();
            _xrmUri = new Uri(context.Properties["OnlineCrmUrl"].ToString());
            _browserType = (BrowserType)Enum.Parse(typeof(BrowserType), context.Properties["BrowserType"].ToString());
            _azureKey = context.Properties["AzureKey"].ToString();
        }
        #endregion
        #region Application Insights Tests
        [TestMethod]
        [TestCategory("ApplicationInsights")]
        [TestCategory("ExceptionTests")]
        public void Test_TrackException()
        {
            ApplicationInsights.TelemetryClient telemetry = new Microsoft.ApplicationInsights.TelemetryClient { InstrumentationKey = _azureKey };
            for (int i = 0; i < 10; i++)
            {
                ExceptionTelemetry ExceptionTelemetry_telemetry = new ExceptionTelemetry();
                DateTime dt = DateTime.Now;
                ExceptionTelemetry_telemetry.SeverityLevel = SeverityLevel.Information;
                ExceptionTelemetry_telemetry.Message = "This is from the Test_TrackException Unit Test";
                ExceptionTelemetry_telemetry.Sequence = i.ToString();
                Debug.WriteLine("Debug for BAST306T Lab");
                telemetry.Context.Operation.Name = "Test_TrackException Unit TEST";
                ExceptionTelemetry_telemetry.Exception = new Exception(String.Format("{0} Exception inside of Test_TrackException Unit Test", i));
                telemetry.TrackException(ExceptionTelemetry_telemetry);
            }
            telemetry.Flush();
            
        }
        #endregion
        #region Pass Tests
        [TestMethod]
        [TestCategory("Pass")]
        public void CreateAccount_Pass()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username.ToSecureString(), _password.ToSecureString());

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                xrmApp.CommandBar.ClickCommand("New");

                xrmApp.Entity.SetValue("name", TestSettings.GetRandomString(5,15));

                xrmApp.Entity.Save();
                
            }
            
        }
        #endregion
        #region Fail Tests
        [TestMethod]
        [TestCategory("Fail")]
        public void CreateAccount_Fail()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username.ToSecureString(), _password.ToSecureString());

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                xrmApp.CommandBar.ClickCommand("New");

                xrmApp.Entity.SetValue("name", TestSettings.GetRandomString(5, 15));

                xrmApp.Entity.Save();
                throw new Exception("Error thrown in Lab");
            }

        }
        #endregion
    }
}