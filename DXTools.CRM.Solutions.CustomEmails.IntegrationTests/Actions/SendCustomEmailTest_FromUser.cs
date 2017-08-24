using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Xrm.Framework.Test.Integration;
using Microsoft.Crm.Sdk.Messages;
using DXTools.CRM.Solutions.CustomEmails.Common;
using DXTools.CRM.Solutions.CustomEmails.Common.Messages;


namespace DXTools.CRM.Solutions.CustomEmails.IntegrationTests.Actions
{
    /// <summary>
    /// Send a custom email where the sender will be a user
    /// </summary>
    [TestClass]
    public class SendCustomEmailTest_FromUser : XrmIntegrationTest
    {
        #region Setup

        protected override OrganizationRequest SetupTriggerRequest()
        {
            Entity contact = CreateContact();

            Entity payment = CreatePayment(contact);
            EntityReference paymentReference = payment.ToEntityReference();

            WhoAmIResponse whoAmI = OrganizationService.Execute(new WhoAmIRequest()) as WhoAmIResponse;
            Entity whoAmIUser = new Entity("systemuser");
            whoAmIUser.Id = whoAmI.UserId;

            EntityCollection to = new EntityCollection();
            to.Entities.Add(contact);

            return new dxtools_SendCustomEmailRequest()
            {
                EmailTemplateName = "DXTools Payment Notification Template"
                ,
                ToRecipient = to
                ,
                FromSender = new EntityReference("systemuser", whoAmI.UserId)
                ,
                Regarding = paymentReference
                ,
                RecordContext = paymentReference
                ,
                Attachments = "DXTools Payment Attachment Sample.docx,DXTools Sample T&C.pdf"
            };
        }

        private Entity CreateContact()
        {
            Entity newRecord = new Entity("contact");
            newRecord["firstname"] = "Mr Test";
            newRecord["lastname"] = "Emails";
            newRecord["emailaddress1"] = "test.emails@testing.com";
            newRecord["parentcustomerid"] = CreateAccount().ToEntityReference();

            newRecord.Id = OrganizationService.Create(newRecord);

            return newRecord;
        }

        private Entity CreateAccount()
        {
            Entity newRecord = new Entity("account");
            newRecord["name"] = "Test LTD";

            newRecord.Id = OrganizationService.Create(newRecord);

            return newRecord;
        }

        private Entity CreatePayment(Entity contact)
        {
            Entity newParentRecord = new Entity("dxtools_payment");
            newParentRecord["dxtools_paymentreference"] = "DXT-" + DateTime.Now.Ticks;
            newParentRecord["dxtools_subject"] = "Integration Test - SendCustomEmailTest_Default at " + DateTime.Now;
            newParentRecord["dxtools_amount"] = new Microsoft.Xrm.Sdk.Money(new Decimal(50.3));
            newParentRecord["dxtools_paymentdirection"] = new OptionSetValue(503530000);
            newParentRecord["dxtools_contactid"] = contact.ToEntityReference();

            newParentRecord.Id = OrganizationService.Create(newParentRecord);

            Entity newRecord = new Entity("dxtools_payment");
            newRecord["dxtools_paymentreference"] = "DXT-" + DateTime.Now.Ticks;
            newRecord["dxtools_subject"] = "Integration Test - SendCustomEmailTest_Default at " + DateTime.Now;
            newRecord["dxtools_amount"] = new Microsoft.Xrm.Sdk.Money(new Decimal(101.5));
            newRecord["dxtools_paymentdirection"] = new OptionSetValue(503530000);
            newRecord["dxtools_contactid"] = contact.ToEntityReference();
            newRecord["dxtools_parentpaymentid"] = newParentRecord.ToEntityReference();

            newRecord.Id = OrganizationService.Create(newRecord);

            return newRecord;
        }

        #endregion

        #region Test

        [TestMethod]
        public void Test_SendCustomEmailTest_Default()
        {
            base.Test();
        }

        #endregion

        #region Verify

        protected override void Verify()
        {
            Assert.IsNull(Error);

            dxtools_SendCustomEmailResponse response = TriggerResponse as dxtools_SendCustomEmailResponse;
            
            Assert.IsNotNull(response);
        }

        #endregion
    }
}
