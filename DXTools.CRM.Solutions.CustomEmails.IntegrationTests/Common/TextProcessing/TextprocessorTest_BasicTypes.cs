using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Sdk;
using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing;
using Microsoft.Xrm.Sdk.Query;

namespace DXTools.CRM.Solutions.CustomEmails.IntegrationTests.Common.TextProcessing
{
    [TestClass]
    public class TextprocessorTest_BasicTypes
    {
        protected OrganizationService CrmOrganisationService { get; set; }

        [TestInitialize]
        public void InitialiseTests()
        {
            CrmConnection crmConnection = new CrmConnection("Xrm");
            crmConnection.ProxyTypesEnabled = true;
            CrmOrganisationService = new OrganizationService(crmConnection);
        }

        [TestMethod]
        public void TestTextprocessor_BasicTypes()
        {
            Entity contact = CreateContact();

            Entity payment = CreatePayment(contact);

            Entity paymentTask = CreatePaymentTask(payment, contact);

            ISyntacticParser syntacticParser = new SyntacticParser(new LexicalParser());
            ITextProcessor textProcessor = new TextProcessor(this.CrmOrganisationService);

            //Test DateTime type
            Assert.AreEqual(((DateTime)payment["createdon"]).ToString(), textProcessor.Process("{createdon}", payment.ToEntityReference()));

            //Test Owner type
            Assert.AreEqual(((EntityReference)payment["ownerid"]).Name.ToString(), textProcessor.Process("{ownerid}", payment.ToEntityReference()));
            Assert.AreEqual(((EntityReference)payment["ownerid"]).Name.ToString(), textProcessor.Process("{ownerid.Name}", payment.ToEntityReference()));
            Assert.AreEqual(((EntityReference)payment["ownerid"]).Id.ToString(), textProcessor.Process("{ownerid.Id}", payment.ToEntityReference()));

            //Test OptionSet type
            Assert.AreEqual(Enum.GetName(typeof(PaymentDirections), PaymentDirections.IN), textProcessor.Process("{dxtools_paymentdirection}", payment.ToEntityReference()));
            Assert.AreEqual(Enum.GetName(typeof(PaymentDirections), PaymentDirections.IN), textProcessor.Process("{dxtools_paymentdirection.Label}", payment.ToEntityReference()));
            Assert.AreEqual(((OptionSetValue)payment["dxtools_paymentdirection"]).Value.ToString(), textProcessor.Process("{dxtools_paymentdirection.Value}", payment.ToEntityReference()));

            //Test ActivityParty type
            Assert.AreEqual("Microsoft.Xrm.Sdk.EntityCollection", textProcessor.Process("{customers}", paymentTask.ToEntityReference()));

            //Test Money type
            Assert.AreEqual(((Money)payment["dxtools_amount"]).Value.ToString(), textProcessor.Process("{dxtools_amount}", payment.ToEntityReference()));

            //Test State (statecode) type
            Assert.AreEqual(Enum.GetName(typeof(State), State.Active), textProcessor.Process("{statecode}", payment.ToEntityReference()));
        }

        private Entity CreateContact()
        {
            Entity newRecord = new Entity("contact");
            newRecord["firstname"] = "Mr Test";
            newRecord["lastname"] = "Emails";
            newRecord["emailaddress1"] = "test.emails@testing.com";
            newRecord["parentcustomerid"] = CreateParentOrganisation().ToEntityReference();

            newRecord.Id = CrmOrganisationService.Create(newRecord);

            return newRecord;
        }

        private Entity CreateParentOrganisation()
        {
            Entity newRecord = new Entity("account");
            newRecord["name"] = "Test Organisation LTD";

            newRecord.Id = CrmOrganisationService.Create(newRecord);

            return newRecord;
        }

        private Entity CreatePayment(Entity contact)
        {
            Entity newRecord = new Entity("dxtools_payment");
            newRecord["dxtools_paymentreference"] = "DXT-" + DateTime.Now.Ticks;
            newRecord["dxtools_subject"] = "Testing email template";
            newRecord["dxtools_amount"] = new Money(new Decimal(101.50));
            newRecord["dxtools_paymentdirection"] = new OptionSetValue((int)PaymentDirections.IN); //IN
            newRecord["dxtools_contactid"] = contact.ToEntityReference();

            newRecord.Id = CrmOrganisationService.Create(newRecord);

            newRecord = CrmOrganisationService.Retrieve(
                newRecord.LogicalName
                , newRecord.Id
                , new ColumnSet(true));

            return newRecord;
        }

        private Entity CreatePaymentTask(Entity payment, Entity contact)
        {
            Entity newRecord = new Entity("dxtools_paymenttask");
            newRecord["subject"] = "DXT-Test" + DateTime.Now.Ticks;
            newRecord["regardingobjectid"] = payment.ToEntityReference();

            Entity customerParty = new Entity("activityparty");
            customerParty["partyid"] = contact.ToEntityReference();
            newRecord["customers"] = new Entity[] { customerParty };

            newRecord.Id = CrmOrganisationService.Create(newRecord);

            return newRecord;
        }
    }
}
