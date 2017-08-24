// <copyright file="SendCustomEmailActivity.cs" company="">
// Copyright (c) 2014 All Rights Reserved
// </copyright>
// <author></author>
// <date>1/29/2014 8:38:29 AM</date>
// <summary>Implements the SendCustomEmailActivity Workflow Activity.</summary>
namespace DXTools.CRM.Solutions.CustomEmails.Workflows
{
    using System;
    using System.Activities;
    using System.ServiceModel;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Workflow;
    using System.Text;
    using System.Collections.Generic;
    using DXTools.CRM.Solutions.CustomEmails.Common;
    using Microsoft.Crm.Sdk.Messages;

    /// <summary>
    /// Send an existing email record. 
    /// This custom workflow activity can be used in conjuntion with "CreateCustomEmailActivity", so this first create a custom email,
    /// other steps may be done in the workflow and then finally the email can be sent using the current activity.
    /// </summary>
    public sealed class SendEmailActivity : BaseCodeActivity
    {
        [Input("Existing Email")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> Email { get; set; }

        /// <summary>
        /// Gets or sets whether to send the email, or to just record it as sent.
        /// 
        /// "true" if the email should be sent; otherwise, "false", just record it as sent
        /// </summary>
        [Input("Issue Send")]
        [Default("true")]
        public InArgument<bool> IssueSend { get; set; }

        [Input("Tracking Token")]
        [Default("")]
        public InArgument<String> TrackingToken { get; set; }

        [Output("Subject of the sent email")]
        [Default("")]
        public OutArgument<String> Subject { get; set; }


        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            if (this.WorkflowContext.InputParameters == null)
                throw new ArgumentNullException("Workflow context doesn't contain any Input Parameter");         

            EntityReference emailToBeSent = this.Email.Get<EntityReference>(executionContext);
            bool issueSend = this.IssueSend.Get<bool>(executionContext);
            string trackingToken = this.TrackingToken.Get<string>(executionContext);

            SendEmailRequest sendRequest = new SendEmailRequest()
            {
                EmailId = emailToBeSent.Id
                ,
                IssueSend = issueSend
                ,
                TrackingToken = trackingToken
            };

            SendEmailResponse sendEmailResponse = this.OrganizationService.Execute(sendRequest) as SendEmailResponse;

            this.Subject.Set(executionContext, sendEmailResponse.Subject);
        }
    }
}