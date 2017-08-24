using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Expressions
{
    public class FunctionParameter : DynamicExpression
    {
        public FunctionParameter(string name)
        {
            this.Value = name;
        }

        public override object Resolve(SymbolContext context)
        {
            return this.Value;
        }
    }
}
