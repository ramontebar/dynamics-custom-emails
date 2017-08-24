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
    /// Send a custom email where the sender will be a queue
    /// </summary>
    [TestClass]
    public class SendCustomEmailTest_FromQueue : XrmIntegrationTest
    {
        #region Instance variables

        private EntityReference senderQueue;

        private EntityReference newEmail;

        #endregion

        #region Setup

        protected override OrganizationRequest SetupTriggerRequest()
        {
            Entity contact = CreateContact();

            Entity payment = CreatePayment(contact);
            EntityReference paymentReference = payment.ToEntityReference();

            senderQueue = CreateQueue();

            EntityCollection to = new EntityCollection();
            to.Entities.Add(contact);

            return new dxtools_SendCustomEmailRequest()
            {
                EmailTemplateName = "DXTools Payment Notification Template"
                ,
                ToRecipient = to
                ,
                FromSender = senderQueue
                ,
                Regarding = paymentReference
                ,
                RecordContext = paymentReference
                ,
                Attachments = "DXTools Payment Attachment Sample.docx,DXTools Sample T&C.pdf"
            };
        }

        private EntityReference CreateQueue()
        {
            Entity newRecord = new Entity("queue");
            newRecord["name"] = "Dx Tools Test Queue " + DateTime.Now;
            newRecord["emailaddress"] = "test.emails@testing.com";

            newRecord.Id = OrganizationService.Create(newRecord);

            return newRecord.ToEntityReference();
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
            Entity newRecord = new Entity("dxtools_payment");
            newRecord["dxtools_paymentreference"] = "DXT-" + DateTime.Now.Ticks;
            newRecord["dxtools_subject"] = "Integration Test - SendCustomEmailTest_Default at " + DateTime.Now;
            newRecord["dxtools_amount"] = new Microsoft.Xrm.Sdk.Money(new Decimal(101.5));
            newRecord["dxtools_paymentdirection"] = new OptionSetValue(503530000);
            newRecord["dxtools_contactid"] = contact.ToEntityReference();

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

            Assert.IsNotNull(response.NewCreatedEmail);

            newEmail = response.NewCreatedEmail;
        }

        #endregion

        #region CleanUp

        protected override void CleanUp()
        {
            base.CleanUp();

            this.OrganizationService.Delete(newEmail.LogicalName, newEmail.Id);

            this.OrganizationService.Delete(senderQueue.LogicalName, senderQueue.Id);
        }

        #endregion
    }
}
