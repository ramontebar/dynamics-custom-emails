using DXTools.CRM.Solutions.CustomEmails.Common;
using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common
{
    public class CreateCustomEmail
    {
        public IOrganizationService OrganizationService 
        { 
            get; 
            private set; 
        }

        public ITextProcessor CRMTextProcessor
        {
            get;
            private set;
        }

        public CreateCustomEmail(IOrganizationService organisationService, ITextProcessor crmTextProcessor)
        {
            this.OrganizationService = organisationService;
            this.CRMTextProcessor = crmTextProcessor;
        }

        /// <summary>
        /// Create an email record using the given dictionary to replace dynamic fields
        /// </summary>
        /// <param name="templateName">Name of existing Email template</param>
        /// <param name="userId">User GUID that will be used to Instantiate Template. See CRM SDK InstantiateTemplateRequest</param>
        /// <param name="dictionary">[Optional] Collection of dynamic fields within the template to be replaced</param>
        /// <param name="regarding">[Optional] Email regarding record</param>
        /// <param name="from">Email sender (activityparty record type)</param>
        /// <param name="toRecipient">List of recipients (activityparty record type)</param>
        /// <param name="attachments">List of attachment names (e.g. "DXTools - Payments.pdf")</param>
        /// <param name="allowDuplicateAttachments">
        /// If allowDuplicateAttachments = True, two or more attachments are allowed with the same filename
        /// </param>
        /// <param name="onlyUseAttachmentsInTemplate">
        /// If OnlyAttachmentsInTemplate = True, the attachments to be considered will be only the ones 
        /// in the given template. Otherwise, it will search across any attachment
        /// </param>
        /// <returns>New Email record</returns>
        public Entity CreateEmail(string templateName, Guid userId, Dictionary<string, string> dictionary, EntityReference regarding, EntityReference from, EntityReferenceCollection toRecipient, string[] attachments, bool allowDuplicateAttachments, bool onlyUseAttachmentsInTemplate)
        {
            return CreateEmail(templateName, userId, dictionary, null, regarding, from, toRecipient, attachments, allowDuplicateAttachments, onlyUseAttachmentsInTemplate);
        }

        /// <summary>
        /// Create an email record using the given dictionary to replace dynamic fields
        /// </summary>
        /// <param name="templateName">Name of existing Email template</param>
        /// <param name="userId">User GUID that will be used to Instantiate Template. See CRM SDK InstantiateTemplateRequest</param>
        /// <param name="context">Exisiting CRM record that sets the root context where dynamics
        /// fields values within the template will be taken from</param>
        /// <param name="regarding">[Optional] Email regarding record</param>
        /// <param name="from">Email sender (activityparty record type)</param>
        /// <param name="toRecipient">List of recipients (activityparty record type)</param>
        /// <param name="attachments">List of attachment names (e.g. "DXTools - Payments.pdf")</param>
        /// <param name="allowDuplicateAttachments">
        /// If allowDuplicateAttachments = True, two or more attachments are allowed with the same filename
        /// </param>
        /// <param name="onlyUseAttachmentsInTemplate">
        /// If onlyAttachmentsInTemplate = True, the attachments to be considered will be only the ones 
        /// in the given template. Otherwise, it will search across any attachment
        /// </param>
        /// <returns>New Email record</returns>
        public Entity CreateEmail(string templateName, Guid userId, EntityReference context, EntityReference regarding, EntityReference from, EntityReferenceCollection toRecipient, string[] attachments, bool allowDuplicateAttachments, bool onlyUseAttachmentsInTemplate)
        {
            return CreateEmail(templateName, userId, null, context, regarding, from, toRecipient, attachments, allowDuplicateAttachments, onlyUseAttachmentsInTemplate);
        }

        /// <summary>
        /// Create an email record using the given dictionary to replace dynamic fields
        /// </summary>
        /// <param name="templateName">Name of existing Email template</param>
        /// <param name="userId">User GUID that will be used to Instantiate Template. See CRM SDK InstantiateTemplateRequest</param>
        /// <param name="dictionary">[Optional] Collection of dynamic fields within the template to be replaced</param>
        /// <param name="context">[Optional] Exisiting CRM record that sets the root context where dynamics
        /// fields values within the template will be taken from</param>
        /// <param name="regarding">[Optional] Email regarding record</param>
        /// <param name="from">Email sender (activityparty record type)</param>
        /// <param name="toRecipient">List of recipients (activityparty record type)</param>
        /// <param name="attachments">List of attachment names (e.g. "DXTools - Payments.pdf")</param>
        /// <param name="allowDuplicateAttachments">
        /// If allowDuplicateAttachments = True, two or more attachments are allowed with the same filename
        /// </param>
        /// <param name="onlyAttachmentsInTemplate">
        /// If onlyUseAttachmentsInTemplate = True, the attachments to be considered will be only the ones 
        /// in the given template. Otherwise, it will search across any attachment
        /// </param>
        /// <returns>New Email record</returns>
        public Entity CreateEmail(string templateName, Guid userId, Dictionary<string, string> dictionary, EntityReference context, EntityReference regarding, EntityReference from, EntityReferenceCollection toRecipient, string[] attachments, bool allowDuplicateAttachments, bool onlyUseAttachmentsInTemplate)
        {
            Entity template = RetrieveEmailTemplate(templateName);

            Entity email = GetEmailInstantiateFromTemplate(template, userId);

            ProcessEmailText(email, dictionary, context);

            CreateEmail(email, regarding, from, toRecipient);

            CreateActivityMimeAttachments(template, email, attachments, allowDuplicateAttachments, onlyUseAttachmentsInTemplate);

            return email;
        }

        private Entity RetrieveEmailTemplate(string templateName)
        {
            QueryExpression query = new QueryExpression("template");
            query.NoLock = true;
            query.ColumnSet = new ColumnSet(new string[] { "body", "subject", "title", "templateid" });
            query.Criteria.AddCondition(new ConditionExpression("title", ConditionOperator.Equal, templateName));

            EntityCollection templatesCollection = this.OrganizationService.RetrieveMultiple(query);

            if (templatesCollection.Entities.Count == 0)
                throw new Exception("A template couldn't be found with name " + templateName);
            else if (templatesCollection.Entities.Count > 1)
                throw new Exception("There is more than one template with name " + templateName);
            else
            {
                TraceService.Trace("Email Template has been retrieved correctly");
                return templatesCollection.Entities[0];
            }

        }

        private Entity GetEmailInstantiateFromTemplate(Entity template, Guid userId)
        {
            TraceService.Trace("Email Template with ID {0} is going to be instanciated by user {1}", template.Id, userId);
            InstantiateTemplateRequest instantiateTemplateRequest = new InstantiateTemplateRequest()
            {
                ObjectId = userId
                  ,
                ObjectType = "systemuser"
                  ,
                TemplateId = template.Id
            };

            InstantiateTemplateResponse instantiateTemplateResponse = this.OrganizationService.Execute(instantiateTemplateRequest) as InstantiateTemplateResponse;
            Entity email = instantiateTemplateResponse.EntityCollection.Entities[0];
            TraceService.Trace("Email Template has been instanciated correctly. Email ID:{0}", email.Id);
            return email;
        }

        private void ProcessEmailText(Entity email, Dictionary<string, string> dictionary, EntityReference context)
        {
            string subject = this.CRMTextProcessor.Process(email["subject"] as string, context, dictionary);
            string description = this.CRMTextProcessor.Process(email["description"] as string, context, dictionary);
            TraceService.Trace("Email Subject has been processed correctly:'{0}'.", subject);
            TraceService.Trace("Email Description (body) has been processed correctly:'{0}'.", description);
            email["subject"] = subject;
            email["description"] = description;
        }

        private void CreateEmail(Entity email, EntityReference regarding, EntityReference from, EntityReferenceCollection toRecipient)
        {
            if (from != null)
            {
                Entity fromActivityParty = new Entity("activityparty");
                fromActivityParty["partyid"] = from;

                email["from"] = new Entity[] { fromActivityParty };
            }

            if (toRecipient != null)
            {
                Entity[] toActivityParties = new Entity[toRecipient.Count];
                for (int x = 0; x < toRecipient.Count; x++)
                {
                    toActivityParties[x] = new Entity("activityparty");
                    toActivityParties[x]["partyid"] = toRecipient[x];
                }
                email["to"] = toActivityParties;
            }
                
            if(regarding != null)
                email["regardingobjectid"] = regarding;

            TraceService.Trace("Email is going to be created");

            email.Id = this.OrganizationService.Create(email);

            TraceService.Trace("Email has been created correctly. ID:{0}",email.Id);
        }
        /// <summary>
        /// Create ActivityMimeAttachment records based on the list of attachment names given.
        /// </summary>
        /// <param name="template">Existing Email template record</param>
        /// <param name="builtEmail">Existing Email record</param>
        /// <param name="attachments">List of attachment names (e.g. "DXTools - Payments.pdf")</param>
        /// <param name="allowDuplicateAttachments">
        /// If allowDuplicateAttachments = True, two or more attachments are allowed with the same filename
        /// </param>
        /// <param name="onlyUseAttachmentsInTemplate">
        /// If OnlyAttachmentsInTemplate = True, the attachments to be considered will be only the ones 
        /// in the given template. Otherwise, it will search across any attachment
        /// </param>
        /// <returns>
        /// List of new ActivityMimeAttachment records, 
        /// which are references from the existing email record to the physical attachments.
        /// </returns>
        public List<Entity> CreateActivityMimeAttachments(Entity template, Entity builtEmail, string[] attachments, bool allowDuplicateAttachments, bool onlyUseAttachmentsInTemplate)
        {
            if (attachments != null && attachments.Length > 0)
            {               
                QueryExpression query = new QueryExpression("activitymimeattachment");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "attachmentid", "filename","objectid" });
                query.Criteria.FilterOperator = LogicalOperator.And;
                query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, "template");

                if (onlyUseAttachmentsInTemplate)
                    query.Criteria.AddCondition("objectid", ConditionOperator.Equal, template.Id);

                FilterExpression childFilter = query.Criteria.AddFilter(LogicalOperator.Or);
                for (int x = 0; x < attachments.Length; x++)
                {
                    childFilter.AddCondition(
                        new ConditionExpression("filename", ConditionOperator.Equal, attachments[x])
                    );
                }

                EntityCollection activitymimeattachmentCollection = this.OrganizationService.RetrieveMultiple(query);

                if(activitymimeattachmentCollection.Entities.Count == 0)
                    return new List<Entity>(0);

                IEnumerable<IGrouping<string, Entity>> activitymimeattachmentsGroupByFilename = activitymimeattachmentCollection.Entities.GroupBy(item => ((string)item["filename"]));
                IEnumerable<IGrouping<Guid, Entity>> activitymimeattachmentsGroupByFilenameAndAttachmentId = null;
                Entity activityMimeAttachment = null;
                List<Entity> activitymimeattachmentsResult = new List<Entity>();

                foreach (IGrouping<string, Entity> groupByFilename in activitymimeattachmentsGroupByFilename)
                {
                    activitymimeattachmentsGroupByFilenameAndAttachmentId = groupByFilename.GroupBy(item => ((EntityReference)item["attachmentid"]).Id);

                    if (allowDuplicateAttachments)
                    {
                        foreach (var activitymimeattachmentGroup in activitymimeattachmentsGroupByFilenameAndAttachmentId)
                        {
                            activityMimeAttachment = new Entity("activitymimeattachment");
                            activityMimeAttachment["objectid"] = new EntityReference("email", builtEmail.Id);
                            activityMimeAttachment["objecttypecode"] = "email";
                            activityMimeAttachment["attachmentid"] = new EntityReference("attachment", activitymimeattachmentGroup.Key);
                            activityMimeAttachment["subject"] = string.Format("Attachment linked to template {0}", template != null && template.Contains("title") ? template["title"] : string.Empty);
                            activityMimeAttachment.Id = this.OrganizationService.Create(activityMimeAttachment);
                            activitymimeattachmentsResult.Add(activityMimeAttachment);
                        }
                    }
                    else
                    {
                        if (activitymimeattachmentsGroupByFilenameAndAttachmentId.Count() > 1)
                            throw new Exception(string.Format("Dupicate attachments are not allowed. There are two o more templates ({0}) with different attachments and same file name ({1}).", GetTemplateNamesFromActivitymimeattachments(activitymimeattachmentsGroupByFilenameAndAttachmentId.SelectMany(i => i)), groupByFilename.Key));
                        else
                        {
                            activityMimeAttachment = new Entity("activitymimeattachment");
                            activityMimeAttachment["objectid"] = new EntityReference("email", builtEmail.Id);
                            activityMimeAttachment["objecttypecode"] = "email";
                            activityMimeAttachment["attachmentid"] = new EntityReference("attachment", activitymimeattachmentsGroupByFilenameAndAttachmentId.First().Key);
                            activityMimeAttachment["subject"] = string.Format("Attachment linked to template {0}", template != null && template.Contains("title") ? template["title"] : string.Empty);
                            activityMimeAttachment.Id = this.OrganizationService.Create(activityMimeAttachment);
                            activitymimeattachmentsResult.Add(activityMimeAttachment);
                        }
                    }                    
                }
                return activitymimeattachmentsResult;
            }
            else
                return new List<Entity>(0);
        }

        private string GetTemplateNamesFromActivitymimeattachments(IEnumerable<Entity> activitymimeattachments)
        {
            int numberOfTemplates = activitymimeattachments.Count();
            if (numberOfTemplates == 0)
                return string.Empty;

            QueryExpression query = new QueryExpression("template");
            query.NoLock = true;
            query.ColumnSet = new ColumnSet(new string[] { "title" });
            query.Criteria = new FilterExpression(LogicalOperator.Or);

            foreach (Entity activitymimeattachment in activitymimeattachments)
                query.Criteria.AddCondition("templateid", ConditionOperator.Equal, ((EntityReference)activitymimeattachment["objectid"]).Id);            

            EntityCollection templateCollection = this.OrganizationService.RetrieveMultiple(query);

            numberOfTemplates = templateCollection.Entities.Count;
               
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < numberOfTemplates; i++)
            {
                sb.Append(templateCollection.Entities[i]["title"]);
                if (i + 1 < numberOfTemplates)
                    sb.Append(", ");
            }
            return sb.ToString();
            
        }

    }
}
