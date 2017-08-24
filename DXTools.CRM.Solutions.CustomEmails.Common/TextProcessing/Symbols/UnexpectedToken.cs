using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Expressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing
{
    public class UnexpectedToken : Expression
    {
        public UnexpectedToken(string token)
        {
            this.Value = token;
        }

        public override object Resolve(ExpressionContext context)
        {
            return string.Empty;
        }
    }
}
