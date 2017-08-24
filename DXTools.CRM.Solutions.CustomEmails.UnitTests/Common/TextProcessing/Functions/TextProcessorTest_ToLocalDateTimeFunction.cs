using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Fakes;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Fakes;
using Microsoft.QualityTools.Testing.Fakes;
using System.Collections.Generic;
using System.Linq;

namespace DXTools.CRM.Solutions.CustomEmails.UnitTests.Common.TextProcessing.Functions
{
    [TestClass]
    public class TextProcessorTest_ToLocalDateTimeFunction
    {
        /// <summary>
        /// Tests the text processor_ to local date time function.
        /// </summary>
        [TestMethod]
        public void TestTextProcessor_ToLocalDateTimeFunction()
        {
            using (ShimsContext.Create())
            {

                DateTime createdOnAttributeESTServer = new DateTime(2015, 04, 27, 11, 01, 00);                    
                string easternZoneId = "Eastern Standard Time";   
                TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById(easternZoneId);
                
                //In UTC that EST Time = 15:01 
                DateTime createdOnAttribute= TimeZoneInfo.ConvertTimeToUtc(createdOnAttributeESTServer, easternZone);
                 

                Entity recordContext = new Entity("dxtools_payment");
                recordContext["createdon"] = createdOnAttribute;
                recordContext.Id = Guid.NewGuid();

                IOrganizationService stubOrganisationService = StubOrganizationServiceFactory.SetupIOrganisationService(recordContext, "createdon", AttributeTypeCode.DateTime);

                ITextProcessor dynTextProcessor = new DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.TextProcessor(stubOrganisationService);

                string inputText = "Your payment was created on {ToString(ToLocalDateTime(createdon,\"GMT Standard Time\"),\"dd/MM/yyyy HH:mm\")} or {createdon}, thanks. ";                
                string expected = "Your payment was created on " + "27/04/2015 16:01" + " or " + createdOnAttribute + ", thanks. ";


                string resultOuput = dynTextProcessor.Process(inputText, recordContext.ToEntityReference());
                //in London should be 27/04/2015 16:01
                Assert.AreEqual(expected, resultOuput);
            }
        }

        ///// <summary>
        ///// Notice quote characters have been encoded where " is now &quot;
        ///// 
        ///// This is required because CRM template content is enconded.
        ///// </summary>
        [TestMethod]
        public void TestTextProcessor_ToLocalDateTimeFunctionEncoded()
        {
            using (ShimsContext.Create())
            {
                DateTime createdOnAttributeESTServer = new DateTime(2015, 04, 27, 11, 01, 00);
                string easternZoneId = "Eastern Standard Time";
                TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById(easternZoneId);

                //In UTC that EST Time = 15:01 
                DateTime createdOnAttribute = TimeZoneInfo.ConvertTimeToUtc(createdOnAttributeESTServer, easternZone);


                Entity recordContext = new Entity("dxtools_payment");
                recordContext["createdon"] = createdOnAttribute;                
                recordContext.Id = Guid.NewGuid();

                IOrganizationService stubOrganisationService = StubOrganizationServiceFactory.SetupIOrganisationService(recordContext, "createdon", AttributeTypeCode.DateTime);

                ITextProcessor dynTextProcessor = new DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.TextProcessor(stubOrganisationService);

                string inputText = "Your payment was created on {ToString(ToLocalDateTime(createdon,&quot;GMT Standard Time&quot;),&quot;dd/MM/yyyy HH:mm&quot;)} or {createdon}, thanks. ";
                string expected = "Your payment was created on " + "27/04/2015 16:01" + " or " + createdOnAttribute + ", thanks. ";

                string resultOuput = dynTextProcessor.Process(inputText, recordContext.ToEntityReference());

                Assert.AreEqual(expected, resultOuput);
            }
        }

        /////// <summary>
        /////// Notice how the format parameter of the ToString function is given using single quotes such as '
        /////// </summary>
        [TestMethod]
        public void TestTextProcessor_ToLocalDateTimeFunctionSingleQuotes()
        {
            using (ShimsContext.Create())
            {
                DateTime createdOnAttributeESTServer = new DateTime(2015, 04, 27, 11, 01, 00);
                string easternZoneId = "Eastern Standard Time";
                TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById(easternZoneId);

                //In UTC that EST Time = 15:01 
                DateTime createdOnAttribute = TimeZoneInfo.ConvertTimeToUtc(createdOnAttributeESTServer, easternZone);


                Entity recordContext = new Entity("dxtools_payment");
                recordContext["createdon"] = createdOnAttribute;
                recordContext.Id = Guid.NewGuid();

                IOrganizationService stubOrganisationService = StubOrganizationServiceFactory.SetupIOrganisationService(recordContext, "createdon", AttributeTypeCode.DateTime);

                ITextProcessor dynTextProcessor = new DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.TextProcessor(stubOrganisationService);

                string inputText = "Your payment was created on {ToString(ToLocalDateTime(createdon,'GMT Standard Time'),'dd/MM/yyyy HH:mm')} or {createdon}, thanks. ";
                string expected = "Your payment was created on " + "27/04/2015 16:01" + " or " + createdOnAttribute + ", thanks. ";

                string resultOuput = dynTextProcessor.Process(inputText, recordContext.ToEntityReference());

                Assert.AreEqual(expected, resultOuput);
            }
        }




    }
}
