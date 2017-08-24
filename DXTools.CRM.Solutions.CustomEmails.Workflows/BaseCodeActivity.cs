using DXTools.CRM.Solutions.CustomEmails.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Workflows
{
    public abstract class BaseCodeActivity : CodeActivity
    {
        [Input("Fail on Exception"), Default("true")]
        public InArgument<bool> FailOnException
        {
            get;
            set;
        }

        [Output("Exception Occured"), Default("false")]
        public OutArgument<bool> ExceptionOccured
        {
            get;
            set;
        }

        [Output("Exception Message"), Default("")]
        public OutArgument<string> ExceptionMessage
        {
            get;
            set;
        }

        protected IWorkflowContext WorkflowContext
        {
            get;
            set;
        }

        protected IOrganizationServiceFactory OrganizationServiceFactory
        {
            get;
            set;
        }

        protected IOrganizationService OrganizationService
        {
            get;
            set;
        }

        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext)
        {
            if (executionContext == null)
                throw new ArgumentNullException("Code Activity Context is null");

            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            DXTools.CRM.Solutions.CustomEmails.Common.TraceService.Initialise(tracingService);

            if (tracingService == null)
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            this.WorkflowContext = context;

            if (context == null)
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            OrganizationServiceFactory = serviceFactory;
            OrganizationService = service;

            tracingService.Trace("Entered custom activity, Correlation Id: {0}, Initiating User: {1}", context.CorrelationId, context.InitiatingUserId);

            try
            {
                TraceService.Trace("Entering ExecuteActivity {0}. Correlation Id: {1}", this.GetType().FullName, context.CorrelationId);
                this.ExecuteActivity(executionContext);
                TraceService.Trace("Ending ExecuteActivity {0}.  Correlation Id: {1}", this.GetType().FullName, context.CorrelationId);
            }
            catch (Exception e)
            {
                ExceptionOccured.Set(executionContext, true);
                ExceptionMessage.Set(executionContext, e.Message);

                if (FailOnException.Get<bool>(executionContext))
                {
                    throw new InvalidPluginExecutionException(e.Message, e);
                }
            }
        }

        protected abstract void ExecuteActivity(CodeActivityContext executionContext);
    }
}
