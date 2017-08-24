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
    public class TextProcessorTest_ToStringFunction
    {
        [TestMethod]
        public void TestTextProcessor_ToStringFunction()
        {
            using (ShimsContext.Create())
            {
                DateTime createdOnAttribute = DateTime.Now;

                Entity recordContext = new Entity("dxtools_payment");
                recordContext["createdon"] = createdOnAttribute;
                recordContext.Id = Guid.NewGuid();

                IOrganizationService stubOrganisationService = StubOrganizationServiceFactory.SetupIOrganisationService(recordContext, "createdon", AttributeTypeCode.DateTime);

                ITextProcessor dynTextProcessor = new DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.TextProcessor(stubOrganisationService);

                string inputText = "Your payment was created on {ToString(createdon,\"dd/MM/yyyy\")} or {createdon}, thanks. ";
                string expected = "Your payment was created on " + DateTime.Now.ToString("dd/MM/yyyy") + " or " + createdOnAttribute + ", thanks. ";

                string resultOuput = dynTextProcessor.Process(inputText, recordContext.ToEntityReference());

                Assert.AreEqual(expected, resultOuput);
            }
        }

        /// <summary>
        /// Notice quote characters have been encoded where " is now &quot;
        /// 
        /// This is required because CRM template content is enconded.
        /// </summary>
        [TestMethod]
        public void TestTextProcessor_ToStringFunctionEncoded()
        {
            using (ShimsContext.Create())
            {
                DateTime createdOnAttribute = DateTime.Now;

                Entity recordContext = new Entity("dxtools_payment");
                recordContext["createdon"] = createdOnAttribute;
                recordContext.Id = Guid.NewGuid();

                IOrganizationService stubOrganisationService = StubOrganizationServiceFactory.SetupIOrganisationService(recordContext, "createdon", AttributeTypeCode.DateTime);

                ITextProcessor dynTextProcessor = new DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.TextProcessor(stubOrganisationService);

                string inputText = "Your payment was created on {ToString(createdon,&quot;dd/MM/yyyy&quot;)} or {createdon}, thanks. ";
                string expected = "Your payment was created on " + DateTime.Now.ToString("dd/MM/yyyy") + " or " + createdOnAttribute + ", thanks. ";

                string resultOuput = dynTextProcessor.Process(inputText, recordContext.ToEntityReference());

                Assert.AreEqual(expected, resultOuput);
            }
        }

        /// <summary>
        /// Notice how the format parameter of the ToString function is given using single quotes such as '
        /// </summary>
        [TestMethod]
        public void TestTextProcessor_ToStringFunctionSingleQuotes()
        {
            using (ShimsContext.Create())
            {
                DateTime createdOnAttribute = DateTime.Now;

                Entity recordContext = new Entity("dxtools_payment");
                recordContext["createdon"] = createdOnAttribute;
                recordContext.Id = Guid.NewGuid();

                IOrganizationService stubOrganisationService = StubOrganizationServiceFactory.SetupIOrganisationService(recordContext, "createdon", AttributeTypeCode.DateTime);

                ITextProcessor dynTextProcessor = new DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.TextProcessor(stubOrganisationService);

                string inputText = "Your payment was created on {ToString(createdon,'dd/MM/yyyy')} or {createdon}, thanks. ";
                string expected = "Your payment was created on " + DateTime.Now.ToString("dd/MM/yyyy") + " or " + createdOnAttribute + ", thanks. ";

                string resultOuput = dynTextProcessor.Process(inputText, recordContext.ToEntityReference());

                Assert.AreEqual(expected, resultOuput);
            }
        }


        /// <summary>
        /// Test Currency type in a function using a lookup attribute as function parameter
        /// </summary>
        [TestMethod]
        public void TestTextProcessor_ToStringFunction_Currency()
        {
            using (ShimsContext.Create())
            {
                Entity parentPayment = new Entity("dxtools_payment");
                parentPayment["dxtools_amount"] = new Money(150);
                parentPayment.Id = Guid.NewGuid();
                StubData parentRecord = new StubData();
                parentRecord.Record = parentPayment;
                parentRecord.AttributesMetadata = new AttributeMetadata[] { 
                        new StubAttributeMetadata(AttributeTypeCode.Money)
                        {
                            LogicalName = "dxtools_amount"
                        }
                };
                
                Entity childPayment = new Entity("dxtools_payment");
                childPayment["dxtools_parentpaymentid"] = parentPayment.ToEntityReference();
                childPayment["dxtools_amount"] = new Decimal(85);
                childPayment.Id = Guid.NewGuid();
                StubData childRecord = new StubData();
                childRecord.Record = childPayment;
                childRecord.AttributesMetadata = new AttributeMetadata[] { 
                        new StubAttributeMetadata(AttributeTypeCode.Money)
                        {
                            LogicalName = "dxtools_amount"
                        },
                        new StubAttributeMetadata(AttributeTypeCode.Lookup)
                        {
                            LogicalName = "dxtools_parentpaymentid"
                        } 
                };

                List<StubData> stubData = new List<StubData>();
                stubData.Add(parentRecord);
                stubData.Add(childRecord);

                IOrganizationService stubOrganisationService = StubOrganizationServiceFactory.SetupIOrganisationService(stubData);

                ITextProcessor dynTextProcessor = new DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.TextProcessor(stubOrganisationService);

                string inputText = "Parent Payment amount: {ToString(dxtools_parentpaymentid.dxtools_amount,'0.00')}";
                string expected = "Parent Payment amount: " + ((Money)parentPayment["dxtools_amount"]).Value.ToString("0.00");

                string resultOuput = dynTextProcessor.Process(inputText, childPayment.ToEntityReference());

                Assert.AreEqual(expected, resultOuput);
            }
        }

    }
}
