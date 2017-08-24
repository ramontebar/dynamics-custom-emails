using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Sdk;
using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing;
using Microsoft.Xrm.Sdk.Query;

namespace DXTools.CRM.Solutions.CustomEmails.IntegrationTests.Common.TextProcessing.Functions
{
    [TestClass]
    public class ToStringTest
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
        public void TestToStringFunction()
        {
            Entity contact = CreateContact("Test 1");
            Entity contact2 = CreateContact("Test 2");

            Entity payment = CreatePayment(contact);

            Entity paymentTask = CreatePaymentTask(payment, contact);
            Entity paymentTask2 = CreatePaymentTask(payment, contact, contact2);

            ISyntacticParser syntacticParser = new SyntacticParser(new LexicalParser());
            ITextProcessor textProcessor = new TextProcessor(this.CrmOrganisationService);

            //Date Time
            Assert.AreEqual(((DateTime)payment["createdon"]).ToString("dd/MM/yyyy"), textProcessor.Process("{ToString(createdon,\"dd/MM/yyyy\")}", payment.ToEntityReference()) );
            Assert.AreEqual(((DateTime)payment["createdon"]).ToString("dd/MM/yyyy"), textProcessor.Process("{ ToString(createdon , \"dd/MM/yyyy\") }", payment.ToEntityReference()));
            Assert.AreEqual(((DateTime)payment["createdon"]).ToString("dd/MM/yyyy"), textProcessor.Process("{ ToString ( createdon , \"dd/MM/yyyy\" ) }", payment.ToEntityReference()));

            //Test ActivityParty type
            string splitSeparator = ", ";
            Assert.AreEqual(contact["fullname"], textProcessor.Process("{ ToString ( customers ) }", paymentTask.ToEntityReference()));
            Assert.AreEqual(contact["fullname"], textProcessor.Process("{ ToString( customers, \"" + splitSeparator + "\") }", paymentTask.ToEntityReference()));
            Assert.AreEqual(contact["fullname"] + splitSeparator + contact2["fullname"], textProcessor.Process("{ToString(customers, \"" + splitSeparator + "\")}", paymentTask2.ToEntityReference()));

            //Test Money type
            Assert.AreEqual(((Money)payment["dxtools_amount"]).Value.ToString("0.00"), textProcessor.Process("{ToString(dxtools_amount,\"0.00\")}", payment.ToEntityReference()));
        }

        private Entity CreateContact(string firstName)
        {
            Entity newRecord = new Entity("contact");
            newRecord["firstname"] = firstName;
            newRecord["lastname"] = DateTime.Now.ToString();
            newRecord["emailaddress1"] = "test.emails@testing.com";
            newRecord["parentcustomerid"] = CreateParentOrganisation().ToEntityReference();

            newRecord.Id = CrmOrganisationService.Create(newRecord);

            newRecord = CrmOrganisationService.Retrieve("contact", newRecord.Id, new ColumnSet(true));

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
            newRecord["dxtools_paymentdirection"] = new OptionSetValue((int)PaymentDirections.IN);
            newRecord["dxtools_contactid"] = contact.ToEntityReference();

            newRecord.Id = CrmOrganisationService.Create(newRecord);

            newRecord = CrmOrganisationService.Retrieve(
                newRecord.LogicalName
                , newRecord.Id
                , new ColumnSet(true));

            return newRecord;
        }

        private Entity CreatePaymentTask(Entity payment, params Entity[] contacts)
        {
            Entity newRecord = new Entity("dxtools_paymenttask");
            newRecord["subject"] = "DXT-Test" + DateTime.Now.Ticks;
            newRecord["regardingobjectid"] = payment.ToEntityReference();

            if (contacts != null && contacts.Length > 0)
            {
                Entity customerParty;
                Entity[] partyContactList = new Entity[contacts.Length];
                for (int i = 0; i < contacts.Length; i++)
                {
                    customerParty = new Entity("activityparty");
                    customerParty["partyid"] = contacts[i].ToEntityReference();
                    partyContactList[i] = customerParty;
                }
                newRecord["customers"] = partyContactList;
            }
            
            newRecord.Id = CrmOrganisationService.Create(newRecord);

            return newRecord;
        }
    }
}
