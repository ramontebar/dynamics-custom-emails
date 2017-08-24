using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Client;
using DXTools.CRM.Solutions.CustomEmails.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Metadata;
using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing;

namespace DXTools.CRM.Solutions.CustomEmails.IntegrationTests.Common
{
    /// <summary>
    /// Test using "DXTools Payment Notification Template" template which is part of the DXTools Sample solution
    /// </summary>
    [TestClass]
    public class CreateCustomEmailTest
    {
        protected OrganizationService CrmOrganisationService { get; set; }

        [TestInitialize]
        public void InitialiseTests()
        {
            CrmConnection crmConnection = new CrmConnection("Xrm");
            CrmOrganisationService = new OrganizationService(crmConnection);
        }

        [TestMethod]
        public void TestCreateCustomEmail()
        {
            Entity contact = CreateContact();

            Entity payment = CreatePayment(contact);
            EntityReference paymentReference = payment.ToEntityReference();

            WhoAmIResponse whoAmI = this.CrmOrganisationService.Execute(new WhoAmIRequest()) as WhoAmIResponse;
            Entity whoAmIUser = new Entity("systemuser");
            whoAmIUser.Id = whoAmI.UserId;

            EntityReferenceCollection to = new EntityReferenceCollection();
            to.Add(contact.ToEntityReference());

            ITextProcessor crmTextProcessor = new TextProcessor(this.CrmOrganisationService);

            CreateCustomEmail createCustomEmail = new CreateCustomEmail(this.CrmOrganisationService, crmTextProcessor);
            createCustomEmail.CreateEmail(
                "DXTools Payment Notification Template"
                , whoAmI.UserId
                , paymentReference
                , paymentReference
                , new EntityReference("systemuser", whoAmI.UserId)
                , to
                , new string[]{"DXTools Payment Attachment Sample.docx", "DXTools Sample T&C.pdf"}
                , false
                , false);
        }

        private Entity CreateContact()
        {
            Entity newRecord = new Entity("contact");
            newRecord["firstname"] = "Mr Test";
            newRecord["lastname"] = "Emails";
            newRecord["emailaddress1"] = "test.emails@testing.com";
            newRecord["parentcustomerid"] = CreateAccount().ToEntityReference();

            newRecord.Id = CrmOrganisationService.Create(newRecord);

            return newRecord;
        }

        private Entity CreateAccount()
        {
            Entity newRecord = new Entity("account");
            newRecord["name"] = "Test LTD";

            newRecord.Id = CrmOrganisationService.Create(newRecord);

            return newRecord;
        }

        private Entity CreatePayment(Entity contact)
        {
            Entity newRecord = new Entity("dxtools_payment");
            newRecord["dxtools_paymentreference"] = "DXT-" + DateTime.Now.Ticks;
            newRecord["dxtools_subject"] = "Testing email template";
            newRecord["dxtools_amount"] = new Microsoft.Xrm.Sdk.Money(new Decimal(101.5));
            newRecord["dxtools_paymentdirection"] = new OptionSetValue(503530000);
            newRecord["dxtools_contactid"] = contact.ToEntityReference();

            newRecord.Id = CrmOrganisationService.Create(newRecord);

            return newRecord;
        }
    }
}
