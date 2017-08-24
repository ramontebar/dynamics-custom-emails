using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Expressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing
{
    public class DynamicExpression : Symbol
    {
        public List<DynamicExpression> ChildExpressions { get; protected set; }

        public DynamicExpression()
        {
            this.ChildExpressions = new List<DynamicExpression>();
        }

        public override object Resolve(SymbolContext context)
        {
            if (ChildExpressions.Count == 0)
                return string.Empty;
            else
                return ChildExpressions[0].Resolve(context);
        }
    }
}
