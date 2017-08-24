using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using DXTools.CRM.Solutions.CustomEmails.Common;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Query;
using System.Diagnostics;
using System.Text;

namespace DXTools.CRM.Solutions.CustomEmails.IntegrationTests.Common
{
    ///<summary>
    /// 
    ///Legend:
    /// t => template
    /// ama => activitymimeattachment
    /// a => attachment
    /// n => attachment.name
    /// b => attachment.body
    ///
    /// Test Cases: 
    ///(1) 
    ///t1 ama1 a1(n1,b1)
    ///t2 ama2 a1(n1,b1)
    ///(2)
    ///t1 ama1 a1(n1,b1)
    ///t2 ama2 a2(n1,b1)
    ///(3)
    ///t1 ama1 a1(n1,b1)
    ///t2 ama2 a2(n1,b2)
    ///(4)
    ///t1 ama1 a1(n1,b1)
    ///t2 ama2 a2(n2,b2)
    ///(5)
    ///t1 ama1 a1(n1,b1)
    ///t2 ama2 a2(n2,b1)
    ///
    /// Cases 1,2,3 are critical to avoid misunderstandings with duplicate attachments
    ///
    ///If AllowDuplicates = False
    ///We don't allow two or more attachments with the same filename.
    ///
    ///If AllowDuplicates = True
    ///We allow two or more attachments with the same filename
    ///
    ///OnlyAttachmentsInTemplate =>  Filter by those attachments in the current template
    /// </summary>
    [TestClass]
    public class CreateActivityMimeAttachmentsTest
    {
        protected OrganizationService CrmOrganisationService { get; set; }

        protected CreateCustomEmail CreateCustomEmailInstance { get; set; }

        [TestInitialize]
        public void InitialiseTests()
        {
            CrmConnection crmConnection = new CrmConnection("Xrm");
            
            CrmOrganisationService = new OrganizationService(crmConnection);

            CreateCustomEmailInstance = new CreateCustomEmail(CrmOrganisationService,null);
 
        }

        /// <summary>
        /// Test case (1): OnlyAttachmentsInTemplate = True
        /// Becuase given attachment is not within the given template, 
        /// although it is part of the existing template "DXTools - Sample T&C", 
        /// no activityMimeAttachment records are created
        /// </summary>
        [TestMethod]
        public void UseOnlyAttachmentsInCurrentTemplate()
        {
            Entity email = new Entity("email");
            email["subject"] = "DXTools - Integration Test - UseOnlyAttachmentsInCurrentTemplate - " + DateTime.Now;
            email.Id = this.CrmOrganisationService.Create(email);

            Entity template = new Entity("template");
            template["title"] = "DXTools - Integration Test - UseOnlyAttachmentsInCurrentTemplate - " + DateTime.Now;
            template["templatetypecode"] = 8;//Global template
            template["ispersonal"] = false;
            template["languagecode"] = 1033;
            template.Id = this.CrmOrganisationService.Create(template);

            //Dot Not Allow duplicates = false
            List<Entity> activityMimeAttachments= CreateCustomEmailInstance.CreateActivityMimeAttachments(template, email, new string[]{"DXTools Integration Test.docx"}, false, true);
            Assert.IsTrue(activityMimeAttachments.Count == 0);

            //Dot Not Allow duplicates = true
            activityMimeAttachments = CreateCustomEmailInstance.CreateActivityMimeAttachments(template, email, new string[] { "DXTools Integration Test.docx" }, true, true);
            Assert.IsTrue(activityMimeAttachments.Count == 0);
        }

        /// <summary>
        /// Test case (2),(3): OnlyAttachmentsInTemplate = True
        /// Two different files with the same name are created. 
        /// Because only attachments from the given template are allowed, 
        /// only one will be used
        /// </summary>
        [TestMethod]
        public void UseOnlyAttachmentsInCurrentTemplate2()
        {
            Entity email = new Entity("email");
            email["subject"] = "DXTools - Integration Test - ReuseSameAttachment - " + DateTime.Now;
            email.Id = this.CrmOrganisationService.Create(email);

            Entity template1 = new Entity("template");
            template1["title"] = "DXTools - Integration Test - Template 1 - UseOnlyAttachmentsInCurrentTemplate2 - " + DateTime.Now;
            template1["templatetypecode"] = 8;//Global template
            template1["ispersonal"] = false;
            template1["languagecode"] = 1033;
            template1.Id = this.CrmOrganisationService.Create(template1);

            Entity template2 = new Entity("template");
            template2["templatetypecode"] = 8;//Global template
            template2["ispersonal"] = false;
            template2["languagecode"] = 1033;
            template2["title"] = "DXTools - Integration Test - Temaplate 2 - UseOnlyAttachmentsInCurrentTemplate2 - " + DateTime.Now;
            template2.Id = this.CrmOrganisationService.Create(template2);

            string attachmentName = string.Format("TestAttachment_{0}.txt", DateTime.Now.Ticks);

            //Add attachment to template 1
            Entity activityMimeAttachment1 = new Entity("activitymimeattachment");
            activityMimeAttachment1["objectid"] = template1.ToEntityReference();
            activityMimeAttachment1["objecttypecode"] = "template";
            activityMimeAttachment1["filename"] = attachmentName;
            activityMimeAttachment1["body"] = Convert.ToBase64String(new UnicodeEncoding().GetBytes("Sample Annotation Text 1"));
            activityMimeAttachment1["mimetype"] = "text/plain";
            this.CrmOrganisationService.Create(activityMimeAttachment1);

            //Add attachment to template 2
            Entity activityMimeAttachment2 = new Entity("activitymimeattachment");
            activityMimeAttachment2["objectid"] = template2.ToEntityReference();
            activityMimeAttachment2["objecttypecode"] = "template";
            activityMimeAttachment2["filename"] = attachmentName;
            activityMimeAttachment2["body"] = Convert.ToBase64String(new UnicodeEncoding().GetBytes("Sample Annotation Text 2"));
            activityMimeAttachment2["mimetype"] = "text/plain";
            this.CrmOrganisationService.Create(activityMimeAttachment2);

            List<Entity> activityMimeAttachments = CreateCustomEmailInstance.CreateActivityMimeAttachments(template2, email, new string[] { attachmentName }, true, true);
            Assert.IsTrue(activityMimeAttachments.Count == 1);
        }

        /// <summary>
        /// Test case (1): OnlyAttachmentsInTemplate = False
        /// Because given attachment is now within the given template, 
        /// although it is part of tye existing template "DXTools - Sample T&C", 
        /// an activityMimeAttachment record is created
        /// </summary>
        [TestMethod]
        public void UseAnyAttachment()
        {
            Entity email = new Entity("email");
            email["subject"] = "DXTools - Integration Test - AttachmentInCurrentTemplate - " + DateTime.Now;
            email.Id = this.CrmOrganisationService.Create(email);

            Entity template = new Entity("template");
            template["title"] = "DXTools - Integration Test - AttachmentInCurrentTemplate - " + DateTime.Now;
            template["templatetypecode"] = 8;//Global template
            template["ispersonal"] = false;
            template["languagecode"] = 1033;
            template.Id = this.CrmOrganisationService.Create(template);

            //Dot Not Allow duplicates = false
            List<Entity> activityMimeAttachments = CreateCustomEmailInstance.CreateActivityMimeAttachments(template, email, new string[] { "DXTools Integration Test.docx" }, false, false);
            Assert.IsTrue(activityMimeAttachments.Count == 1);

            //Dot Not Allow duplicates = true (expected same result as above)
            activityMimeAttachments = CreateCustomEmailInstance.CreateActivityMimeAttachments(template, email, new string[] { "DXTools Integration Test.docx" }, true, false);
            Assert.IsTrue(activityMimeAttachments.Count == 1);
        }

        /// <summary>
        /// Test case (1): Only one physical attachment is reused
        /// 
        /// This test is the same as previous and you can see how attachment is reused
        /// </summary>
        public void UseAnyAttachmentReusingExisting()
        {
            Entity email = new Entity("email");
            email["subject"] = "DXTools - Integration Test - AttachmentInCurrentTemplate - " + DateTime.Now;
            email.Id = this.CrmOrganisationService.Create(email);

            Entity template = new Entity("template");
            template["title"] = "DXTools - Integration Test - AttachmentInCurrentTemplate - " + DateTime.Now;
            template["templatetypecode"] = 8;//Global template
            template["ispersonal"] = false;
            template["languagecode"] = 1033;
            template.Id = this.CrmOrganisationService.Create(template);

            Entity attachment = RetrieveAttachment("DXTools Integration Test.docx");

            Entity activityMimeAttachment = new Entity("activitymimeattachment");
            activityMimeAttachment["objectid"] = new EntityReference("template", template.Id);
            activityMimeAttachment["objecttypecode"] = "template";
            activityMimeAttachment["attachmentid"] = attachment.ToEntityReference();
            activityMimeAttachment.Id = this.CrmOrganisationService.Create(activityMimeAttachment);

            //Dot Not Allow duplicates = false
            List<Entity> activityMimeAttachments = CreateCustomEmailInstance.CreateActivityMimeAttachments(template, email, new string[] { "DXTools Integration Test.docx" }, false, false);
            Assert.IsTrue(activityMimeAttachments.Count == 1);

            //Dot Not Allow duplicates = true (expected same result as above)
            activityMimeAttachments = CreateCustomEmailInstance.CreateActivityMimeAttachments(template, email, new string[] { "DXTools Integration Test.docx" }, true, false);
            Assert.IsTrue(activityMimeAttachments.Count == 1);
        }

        /// <summary>
        /// Test case (1): Only one physical attachment is reused
        /// </summary>
        [TestMethod]
        public void ReuseSameAttachment()
        {
            Entity email = new Entity("email");
            email["subject"] = "DXTools - Integration Test - ReuseSameAttachment - " + DateTime.Now;
            email.Id = this.CrmOrganisationService.Create(email);

            Entity template1 = new Entity("template");
            template1["title"] = "DXTools - Integration Test - Template 1 - ReuseSameAttachment - " + DateTime.Now;
            template1["templatetypecode"] = 8;//Global template
            template1["ispersonal"] = false;
            template1["languagecode"] = 1033;
            template1.Id = this.CrmOrganisationService.Create(template1);

            Entity template2 = new Entity("template");
            template2["templatetypecode"] = 8;//Global template
            template2["ispersonal"] = false;
            template2["languagecode"] = 1033;
            template2["title"] = "DXTools - Integration Test - Temaplate 2 - ReuseSameAttachment - " + DateTime.Now;
            template2.Id = this.CrmOrganisationService.Create(template2);

            string attachmentName = string.Format("TestAttachment_{0}.txt", DateTime.Now.Ticks);

            //Add attachment to template 1
            Entity activityMimeAttachment1 = new Entity("activitymimeattachment");
            activityMimeAttachment1["objectid"] = template1.ToEntityReference();
            activityMimeAttachment1["objecttypecode"] = "template";
            activityMimeAttachment1["filename"] = attachmentName;
            activityMimeAttachment1["body"] = Convert.ToBase64String(new UnicodeEncoding().GetBytes("Sample Annotation Text"));
            activityMimeAttachment1["mimetype"] = "text/plain";
            this.CrmOrganisationService.Create(activityMimeAttachment1);

            //Add existing attachment to template 2
            //Dot Not Allow duplicates = false
            List<Entity> activityMimeAttachments = CreateCustomEmailInstance.CreateActivityMimeAttachments(template2, email, new string[] { attachmentName }, false, false);
            Assert.IsTrue(activityMimeAttachments.Count == 1);
            Assert.IsTrue(RetrieveAttachments(attachmentName).Entities.Count == 1);
            this.CrmOrganisationService.Delete("activitymimeattachment", activityMimeAttachments[0].Id); //Clean

            //Add existing attachment to template 2
            //Dot Not Allow duplicates = true  (expected same result as above)
            activityMimeAttachments = CreateCustomEmailInstance.CreateActivityMimeAttachments(template2, email, new string[] { attachmentName }, true, false);
            Assert.IsTrue(activityMimeAttachments.Count == 1);
            Assert.IsTrue(RetrieveAttachments(attachmentName).Entities.Count == 1);
            this.CrmOrganisationService.Delete("activitymimeattachment", activityMimeAttachments[0].Id); //Clean

            //Clean (ToDo)
        }

        /// <summary>
        /// Test case (2) and (3): Different physical attachments with same filename
        /// </summary>
        [TestMethod]
        public void DifferentAttachments()
        {
            Entity email = new Entity("email");
            email["subject"] = "DXTools - Integration Test - ReuseSameAttachment - " + DateTime.Now;
            email.Id = this.CrmOrganisationService.Create(email);

            Entity template1 = new Entity("template");
            template1["title"] = "DXTools - Integration Test - Template 1 - DifferentAttachments - " + DateTime.Now;
            template1["templatetypecode"] = 8;//Global template
            template1["ispersonal"] = false;
            template1["languagecode"] = 1033;
            template1.Id = this.CrmOrganisationService.Create(template1);

            Entity template2 = new Entity("template");
            template2["templatetypecode"] = 8;//Global template
            template2["ispersonal"] = false;
            template2["languagecode"] = 1033;
            template2["title"] = "DXTools - Integration Test - Temaplate 2 - DifferentAttachments - " + DateTime.Now;
            template2.Id = this.CrmOrganisationService.Create(template2);

            Entity template3 = new Entity("template");
            template3["templatetypecode"] = 8;//Global template
            template3["ispersonal"] = false;
            template3["languagecode"] = 1033;
            template3["title"] = "DXTools - Integration Test - Temaplate 3 - DifferentAttachments - " + DateTime.Now;
            template3.Id = this.CrmOrganisationService.Create(template3);

            string attachmentName = string.Format("TestAttachment_{0}.txt", DateTime.Now.Ticks);

            //Add attachment to template 1
            Entity activityMimeAttachment1 = new Entity("activitymimeattachment");
            activityMimeAttachment1["objectid"] = template1.ToEntityReference();
            activityMimeAttachment1["objecttypecode"] = "template";
            activityMimeAttachment1["filename"] = attachmentName;
            activityMimeAttachment1["body"] = Convert.ToBase64String(new UnicodeEncoding().GetBytes("Sample Annotation Text 1"));
            activityMimeAttachment1["mimetype"] = "text/plain";
            this.CrmOrganisationService.Create(activityMimeAttachment1);

            //Add attachment to template 2
            Entity activityMimeAttachment2 = new Entity("activitymimeattachment");
            activityMimeAttachment2["objectid"] = template2.ToEntityReference();
            activityMimeAttachment2["objecttypecode"] = "template";
            activityMimeAttachment2["filename"] = attachmentName;
            activityMimeAttachment2["body"] = Convert.ToBase64String(new UnicodeEncoding().GetBytes("Sample Annotation Text 2"));
            activityMimeAttachment2["mimetype"] = "text/plain";
            this.CrmOrganisationService.Create(activityMimeAttachment2);

            //Add attachment to template 3 (Allow duplicates)
            List<Entity> activityMimeAttachments = CreateCustomEmailInstance.CreateActivityMimeAttachments(template3, email, new string[] { attachmentName }, true, false);
            Assert.IsTrue(activityMimeAttachments.Count == 2);
            Assert.IsTrue(RetrieveAttachments(attachmentName).Entities.Count == 2);
            this.CrmOrganisationService.Delete("activitymimeattachment", activityMimeAttachments[0].Id); //Clean
            this.CrmOrganisationService.Delete("activitymimeattachment", activityMimeAttachments[1].Id);

            //Add attachment to template 3 (Do Not Allow duplicates)
            try
            {
                activityMimeAttachments = CreateCustomEmailInstance.CreateActivityMimeAttachments(template3, email, new string[] { attachmentName }, false, false);
            }
            catch (Exception exc)
            {
                Assert.IsTrue(!string.IsNullOrEmpty(exc.Message));
                Assert.IsTrue(exc.Message.Contains("Dupicate attachments are not allowed"));
                Assert.IsTrue(exc.Message.Contains(attachmentName));
                Assert.IsTrue(exc.Message.Contains(template1["title"].ToString()));
                Assert.IsTrue(exc.Message.Contains(template2["title"].ToString()));
            }
        }


        #region Private Helpers Methods

        private Entity RetrieveAttachment(string attachmentName)
        {
            EntityCollection recordCollection = RetrieveAttachments(attachmentName);
            if (recordCollection.Entities.Count > 0)
                return recordCollection.Entities[0];
            else
                throw new Exception("No attachments with name " + attachmentName);
        }

        private EntityCollection RetrieveAttachments(string attachmentName)
        {
            QueryExpression query = new QueryExpression("attachment");
            query.Criteria.AddCondition("filename", ConditionOperator.Equal, attachmentName);
            query.ColumnSet = new ColumnSet(true);
            query.NoLock = true;

            EntityCollection recordCollection = this.CrmOrganisationService.RetrieveMultiple(query);
            return recordCollection;
        }

        //[TestMethod]
        //public void GetTemplates()
        //{
        //    QueryExpression query = new QueryExpression("template");
        //    query.ColumnSet = new ColumnSet(true);
        //    query.NoLock = true;

        //    EntityCollection templates = this.CrmOrganisationService.RetrieveMultiple(query);
        //    foreach (Entity template in templates.Entities)
        //        Debug.WriteLine(string.Format("{0} - {1}",template["title"],template["templatetypecode"]));
        //}

        #endregion
    }
}
