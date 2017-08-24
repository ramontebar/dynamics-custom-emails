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
    public class SendCustomEmailTest_NotAllowDuplicates : XrmIntegrationTest
    {
        #region Local Test Variables

        private Entity payment;

        private string attachmentName;

        private Entity template1;

        private Entity template2;

        #endregion

        #region Setup

        protected override OrganizationRequest SetupTriggerRequest()
        {
            attachmentName = string.Format("TestAttachment_{0}.txt", DateTime.Now.Ticks);

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
                AllowDuplicateAttachments = false
                ,
                OnlyUseAttachmentsInTemplate = false
            };
        }

        private Entity CreateTemplates(string attachmentName)
        {
            //Template 1
            string title1 = "DXTools - Integration Test - Template 1 - SendCustomEmailTest_NotAllowDuplicates - Title - " + DateTime.Now;
            string body1 = "DXTools - Integration Test - Template 1 - SendCustomEmailTest_NotAllowDuplicates - Body - " + DateTime.Now;
            template1 = new Entity("template");
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
            string title2 = "DXTools - Integration Test - Temaplate 2 - SendCustomEmailTest_NotAllowDuplicates - Title" + DateTime.Now;
            string body2 = "DXTools - Integration Test - Temaplate 2 - SendCustomEmailTest_NotAllowDuplicates - Body" + DateTime.Now;
            template2 = new Entity("template");
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
            newRecord["dxtools_subject"] = "Integration Test - SendCustomEmailTest_NotAllowDuplicates at " + DateTime.Now;
            newRecord["dxtools_amount"] = new Microsoft.Xrm.Sdk.Money(new Decimal(101.5));
            newRecord["dxtools_paymentdirection"] = new OptionSetValue(503530000);
            newRecord["dxtools_contactid"] = contact.ToEntityReference();

            newRecord.Id = OrganizationService.Create(newRecord);

            return newRecord;
        }

        #endregion

        #region Test

        [TestMethod]
        public void Test_SendCustomEmailTest_NotAllowDuplicates()
        {
            base.Test();
        }

        #endregion

        #region Verify

        protected override void Verify()
        {
            Assert.IsNotNull(Error);

            Assert.IsTrue(Error.Message.Contains(template1["title"].ToString()));

            Assert.IsTrue(Error.Message.Contains(template2["title"].ToString()));

            Assert.IsTrue(Error.Message.Contains(attachmentName));
        }

        #endregion
    }
}
