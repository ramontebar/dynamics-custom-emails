using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using DXTools.CRM.Solutions.CustomEmails.Common.Helpers;

namespace DXTools.CRM.Solutions.CustomEmails.IntegrationTests.Common
{
    [TestClass]
    public class CreateTemplateProgramaticallyTest
    {
        protected OrganizationService CrmOrganisationService { get; set; }

        [TestInitialize]
        public void InitialiseTests()
        {
            CrmConnection crmConnection = new CrmConnection("Xrm");
            CrmOrganisationService = new OrganizationService(crmConnection);
        }

        [TestMethod]
        public void TestCreateTemplateProgramatically()
        {
            string templateName = "DXTools - Test Create Template Programatically " + DateTime.Now;
            string subject = "DXTools Template created Programatically " + DateTime.Now;
            string body = "<p><strong>Amazing</strong></p>";

            Entity newTemplate = BuildTemplate(templateName, subject, body); 

            //Validate Template has been created correctly
            Entity existingTemplateRecord = TemplatesHelper.RetrieveTemplateByName(templateName, this.CrmOrganisationService);
            Assert.IsNotNull(existingTemplateRecord);
            Assert.IsTrue(existingTemplateRecord["subject"].ToString().Contains(subject));
            Assert.IsTrue(existingTemplateRecord["body"].ToString().Contains(body));

            //Validate email based on template is correct
            Entity emailRecord = GetEmailInstantiateFromTemplate(newTemplate, GetCurrentUser().Id);
            Assert.IsTrue(emailRecord["subject"].ToString().Contains(subject));
            Assert.IsTrue(emailRecord["description"].ToString().Contains(body)); //Email body

            //Validate email can be created in CRM succesfully
            this.CrmOrganisationService.Create(emailRecord);
        }

        private Entity GetCurrentUser()
        {
            WhoAmIResponse whoAmI = this.CrmOrganisationService.Execute(new WhoAmIRequest()) as WhoAmIResponse;
            Entity whoAmIUser = new Entity("systemuser");
            whoAmIUser.Id = whoAmI.UserId;
            return whoAmIUser;
        }


        private Entity BuildTemplate(string templateTitle, string subject, string body)
        {
            Entity newTemplate = new Entity("template");
            newTemplate["title"] = templateTitle;
            newTemplate["templatetypecode"] = 8;//Global template
            newTemplate["ispersonal"] = false;
            newTemplate["languagecode"] = 1033;

            //IMPORTANT NOTE: 
            //Both fields are mandatory to set 'subject' properly
            newTemplate["subject"] = TemplatesHelper.GetSubjectXML(subject); ;
            newTemplate["subjectpresentationxml"] = TemplatesHelper.GetSubjectPresentationXML(subject);

            //IMPORTANT NOTE: 
            //Both fields are mandatory to set 'body' properly
            newTemplate["body"] = TemplatesHelper.GetBodyXML(body); ;
            newTemplate["presentationxml"] = TemplatesHelper.GetBodyPresentationXML(body);

            newTemplate.Id = this.CrmOrganisationService.Create(newTemplate);

            return newTemplate;
        }

        private Entity GetEmailInstantiateFromTemplate(Entity template, Guid userId)
        {
            InstantiateTemplateRequest instantiateTemplateRequest = new InstantiateTemplateRequest()
            {
                ObjectId = userId
                  ,
                ObjectType = "systemuser"
                  ,
                TemplateId = template.Id
            };

            InstantiateTemplateResponse instantiateTemplateResponse = this.CrmOrganisationService.Execute(instantiateTemplateRequest) as InstantiateTemplateResponse;
            return instantiateTemplateResponse.EntityCollection.Entities[0];
        }
    }
}
