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
    public class SendCustomEmailTest_WithStaticParameters : XrmIntegrationTest
    {
        #region Local Test Variables

        private Entity payment;
        private Entity template;
        private string rowEmailSubject;
        private string rowEmailBody;

        #endregion

        #region Setup

        protected override OrganizationRequest SetupTriggerRequest()
        {
            Entity template = CreateTemplate();

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

        private Entity CreateTemplate()
        {
            Entity template = new Entity("template");
            template["title"] = "DXTools - Integration Test - SendCustomEmailTest_WithTypeFormatters - " + DateTime.Now;
            template["templatetypecode"] = 8;//Global template
            template["ispersonal"] = false;
            template["languagecode"] = 1033;

            rowEmailSubject = "Integration Test " + DateTime.Now;
            rowEmailBody = "<p><strong> This is an integration test {CurrentVAT} </strong></p>";
            template["subject"] = TemplatesHelper.GetSubjectXML(rowEmailSubject);
            template["subjectpresentationxml"] = TemplatesHelper.GetSubjectPresentationXML(rowEmailSubject);
            template["body"] = TemplatesHelper.GetBodyXML(rowEmailBody);
            template["presentationxml"] = TemplatesHelper.GetBodyPresentationXML(rowEmailBody);

            template.Id = OrganizationService.Create(template);

            return template;

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

            newRecord.Id = OrganizationService.Create(newRecord);

            return newRecord;
        }

        #endregion

        #region Test

        [TestMethod]
        public void TestSendCustomEmail_WithStaticParameters()
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
            string expectedSubject = rowEmailSubject;
            string expectedBody = rowEmailBody.Replace("{CurrentVAT}", "20%");

            Assert.IsTrue(sentEmail["subject"].ToString().Contains(expectedSubject));
            Assert.IsTrue(sentEmail["description"].ToString().Contains(expectedBody));
           
        }

        #endregion
    }
}
