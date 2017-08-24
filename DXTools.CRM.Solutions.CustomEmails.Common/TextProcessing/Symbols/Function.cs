using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Expressions
{
    public class Function : DynamicExpression
    {
        public Function(string name)
        {
            this.Value = string.IsNullOrEmpty(name) ? string.Empty : name.Trim();
            this.ChildExpressions = new List<DynamicExpression>();
        }

        public override object Resolve(SymbolContext context)
        {
            if (string.IsNullOrEmpty(this.Value))
                return string.Empty;

            string functionName = this.Value.Trim();

            object[] parametes = new object[ChildExpressions.Count];
            if(ChildExpressions.Count > 0)
            {
                for (int i = 0; i < parametes.Length; i++)
			    {
			        parametes[i] = ChildExpressions[i].Resolve(context);
			    }
            }

            Type functionType = Type.GetType("DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Functions." + functionName);
            IDynFunctionHandler functionHanlder =  Activator.CreateInstance(functionType) as IDynFunctionHandler;

            return functionHanlder.Invoke(parametes);
        }
    }
}
