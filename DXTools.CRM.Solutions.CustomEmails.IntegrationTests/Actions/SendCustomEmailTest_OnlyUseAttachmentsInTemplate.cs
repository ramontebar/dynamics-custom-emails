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
    public class SendCustomEmailTest_OnlyUseAttachmentsInTemplate : XrmIntegrationTest
    {
        #region Local Test Variables

        private Entity payment;
        private string rowEmailSubject1 = "DXTools - Integration Test - Template 1 - SendCustomEmailTest_OnlyUseAttachmentsInTemplate - Title - ";
        private string rowEmailBody1 = "DXTools - Integration Test - Template 1 - SendCustomEmailTest_OnlyUseAttachmentsInTemplate - Body - ";
        private string rowEmailSubject2 = "DXTools - Integration Test - Template 2 - SendCustomEmailTest_OnlyUseAttachmentsInTemplate - Title - ";
        private string rowEmailBody2 = "DXTools - Integration Test - Template 2 - SendCustomEmailTest_OnlyUseAttachmentsInTemplate - Body - ";
        #endregion

        #region Setup

        protected override OrganizationRequest SetupTriggerRequest()
        {
            string attachmentName = string.Format("TestAttachment_{0}.txt", DateTime.Now.Ticks);

            Entity template = CreateTemplates(attachmentName);

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
                Attachments = attachmentName
                ,
                AllowDuplicateAttachments = true
                ,
                OnlyUseAttachmentsInTemplate = true
            };
        }

        private Entity CreateTemplates(string attachmentName)
        {
            string title1 = rowEmailSubject1 + DateTime.Now;
            string body1 = rowEmailBody1 + DateTime.Now;
            Entity template1 = new Entity("template");
            template1["title"] = title1;
            template1["subject"] = TemplatesHelper.GetSubjectXML(title1);
            template1["subjectpresentationxml"] = TemplatesHelper.GetSubjectPresentationXML(title1);
            template1["body"] = TemplatesHelper.GetBodyXML(body1);
            template1["presentationxml"] = TemplatesHelper.GetBodyPresentationXML(body1);
            template1["templatetypecode"] = 8;//Global template
            template1["ispersonal"] = false;
            template1["languagecode"] = 1033;
            template1.Id = OrganizationService.Create(template1);

            //Template 2
            string title2 = rowEmailSubject2 + DateTime.Now;
            string body2 = rowEmailBody2 + DateTime.Now;
            Entity template2 = new Entity("template");
            template2["title"] = title2;
            template2["subject"] = TemplatesHelper.GetSubjectXML(title2);
            template2["subjectpresentationxml"] = TemplatesHelper.GetSubjectPresentationXML(title2);
            template2["body"] = TemplatesHelper.GetBodyXML(body2);
            template2["presentationxml"] = TemplatesHelper.GetBodyPresentationXML(body2);
            template2["templatetypecode"] = 8;//Global template
            template2["ispersonal"] = false;
            template2["languagecode"] = 1033;
            template2.Id = OrganizationService.Create(template2);

            //Add attachment to template 1
            Entity activityMimeAttachment1 = new Entity("activitymimeattachment");
            activityMimeAttachment1["objectid"] = template1.ToEntityReference();
            activityMimeAttachment1["objecttypecode"] = "template";
            activityMimeAttachment1["filename"] = attachmentName;
            activityMimeAttachment1["body"] = Convert.ToBase64String(new UnicodeEncoding().GetBytes("Sample Annotation Text 1"));
            activityMimeAttachment1["mimetype"] = "text/plain";
            OrganizationService.Create(activityMimeAttachment1);

            //Add attachment to template 2
            Entity activityMimeAttachment2 = new Entity("activitymimeattachment");
            activityMimeAttachment2["objectid"] = template2.ToEntityReference();
            activityMimeAttachment2["objecttypecode"] = "template";
            activityMimeAttachment2["filename"] = attachmentName;
            activityMimeAttachment2["body"] = Convert.ToBase64String(new UnicodeEncoding().GetBytes("Sample Annotation Text 2"));
            activityMimeAttachment2["mimetype"] = "text/plain";
            OrganizationService.Create(activityMimeAttachment2);

            return template1;

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
            newRecord["dxtools_subject"] = "Integration Test - SendCustomEmailTest_OnlyUseAttachmentsInTemplate at " + DateTime.Now;
            newRecord["dxtools_amount"] = new Microsoft.Xrm.Sdk.Money(new Decimal(101.5));
            newRecord["dxtools_paymentdirection"] = new OptionSetValue(503530000);
            newRecord["dxtools_contactid"] = contact.ToEntityReference();

            newRecord.Id = OrganizationService.Create(newRecord);

            return newRecord;
        }

        #endregion

        #region Test

        [TestMethod]
        public void Test_SendCustomEmailTest_OnlyUseAttachmentsInTemplate()
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
            retrieveEmailQuery.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, payment.Id);
            retrieveEmailQuery.Criteria.AddCondition("subject", ConditionOperator.BeginsWith, rowEmailSubject1);
            EntityCollection emailCollection = this.OrganizationService.RetrieveMultiple(retrieveEmailQuery);

            Assert.IsTrue(emailCollection.Entities.Count == 1);

            Entity sentEmail = emailCollection.Entities[0];

            QueryExpression retrieveEmailAttachmentsQuery = new QueryExpression("activitymimeattachment");
            retrieveEmailAttachmentsQuery.NoLock = true;
            retrieveEmailAttachmentsQuery.ColumnSet = new ColumnSet(new string[] { "attachmentid" });
            retrieveEmailAttachmentsQuery.Criteria.AddCondition("objectid", ConditionOperator.Equal, sentEmail.Id);
            EntityCollection emailAttachmentCollection = this.OrganizationService.RetrieveMultiple(retrieveEmailAttachmentsQuery);

            Assert.IsTrue(emailAttachmentCollection.Entities.Count == 1);
        }

        #endregion
    }
}
