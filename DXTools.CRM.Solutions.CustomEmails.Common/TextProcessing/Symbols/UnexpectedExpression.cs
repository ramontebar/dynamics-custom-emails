using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Expressions
{
    public class UnexpectedExpression : Expression
    {
        public override object Resolve(ExpressionContext context)
        {
            return string.Empty;
        }
    }
}
