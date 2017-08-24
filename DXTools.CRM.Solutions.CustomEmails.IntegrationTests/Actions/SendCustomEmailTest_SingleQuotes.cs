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
using DXTools.CRM.Solutions.CustomEmails.Common.Helpers;


namespace DXTools.CRM.Solutions.CustomEmails.IntegrationTests.Actions
{
    /// <summary>
    /// 
    /// </summary>
    [TestClass]
    public class SendCustomEmailTest_SingleQuotes : XrmIntegrationTest
    {
        #region Local Test Variables

        private const string templateName = "DXTools Quote Test";
        private Entity payment;
        private Entity template;
        private string rowEmailSubject;

        #endregion

        #region Setup

        protected override OrganizationRequest SetupTriggerRequest()
        {
            Entity template = TemplatesHelper.RetrieveTemplateByName(templateName, this.OrganizationService);
            
            if (!template.Contains("title"))
                Assert.Fail("Template doesn't contain a title");
            rowEmailSubject = template["title"].ToString();

            Entity contact = CreateContact();

            payment = CreatePayment(contact);
            EntityReference paymentReference = payment.ToEntityReference();

            WhoAmIResponse whoAmI = OrganizationService.Execute(new WhoAmIRequest()) as WhoAmIResponse;
            Entity whoAmIUser = new Entity("systemuser");
            whoAmIUser.Id = whoAmI.UserId;

            EntityCollection to = new EntityCollection();
            to.Entities.Add(contact);

            return new dxtools_SendCustomEmailRequest()
            {
                EmailTemplateName = template["title"].ToString()
                ,
                ToRecipient = to
                ,
                FromSender = new EntityReference("systemuser", whoAmI.UserId)
                ,
                Regarding = paymentReference
                ,
                RecordContext = paymentReference
                ,
                Attachments = string.Empty
                ,
                StaticParameters = "{\"currentvat\":\"20%\"}"
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
            Entity newRecord = new Entity("dxtools_payment");
            newRecord["dxtools_paymentreference"] = "DXT-" + DateTime.Now.Ticks;
            newRecord["dxtools_subject"] = "Integration Test - SendCustomEmailTest_WithTypeFormatters at " + DateTime.Now;
            newRecord["dxtools_amount"] = new Microsoft.Xrm.Sdk.Money(new Decimal(101.5));
            newRecord["dxtools_paymentdirection"] = new OptionSetValue(503530000);
            newRecord["dxtools_contactid"] = contact.ToEntityReference();

            Guid newPaymentRecordId = OrganizationService.Create(newRecord);

            Entity createdPayment = this.OrganizationService.Retrieve("dxtools_payment", newPaymentRecordId, new ColumnSet(true));

            return createdPayment;
        }

        #endregion

        #region Test

        [TestMethod]
        public void TestSendCustomEmail_SingleQuotes()
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

            //Retrieve sent email
            QueryExpression retrieveEmailQuery = new QueryExpression("email");
            retrieveEmailQuery.NoLock = true;
            retrieveEmailQuery.ColumnSet = new ColumnSet(true);
            retrieveEmailQuery.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, payment.Id);
            retrieveEmailQuery.Criteria.AddCondition("subject", ConditionOperator.BeginsWith, rowEmailSubject);
            EntityCollection emailCollection = this.OrganizationService.RetrieveMultiple(retrieveEmailQuery);
            Assert.IsTrue(emailCollection.Entities.Count == 1);
            
            Entity sentEmail = emailCollection.Entities[0];
            Assert.IsTrue(sentEmail.Contains("description"));

            string emailBody = sentEmail["description"].ToString();
            int createdOnStringIndex = emailBody.IndexOf("created on"); //date lenght = 10 characters (e.g. 11/11/2014)
            int dateIndex = createdOnStringIndex + 11;
            string createdOnDateValueString = emailBody.Substring(createdOnStringIndex + 11, 10);

            Assert.IsTrue(payment.Contains("createdon"));
            DateTime paymentCreatedOn = (DateTime)payment["createdon"];

            Assert.AreEqual(createdOnDateValueString, paymentCreatedOn.ToString("dd/MM/yyyy"));

            string stringAfterDateValue = emailBody.Substring(dateIndex + 10,1);

            if (stringAfterDateValue != ".")
                Assert.Fail("The caracther '.' was expected after the date value");
        }

        #endregion
    }
}
