using DXTools.CRM.Solutions.CustomEmails.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Workflows
{
    public class GetProcessPropertiesActivity : BaseCodeActivity
    {
        protected override void ExecuteActivity(System.Activities.CodeActivityContext executionContext)
        {
            TraceService.Trace("Entering GetProcessPropertiesActivity");

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Worflow Category: " + WorkflowContext.WorkflowCategory);

            sb.AppendLine("The following Values were in the InputParameters:");
            foreach (KeyValuePair<string, object> pair in WorkflowContext.InputParameters)
            {
                sb.AppendLine(pair.Key);
            }

            sb.AppendLine("The following Values were in the OutputParameters:");
            foreach (KeyValuePair<string, object> pair in WorkflowContext.OutputParameters)
            {
                sb.AppendLine(pair.Key);
            }


            sb.AppendLine("End of GetProcessPropertiesActivity");

            throw new Exception(sb.ToString());
        }
    }
}
