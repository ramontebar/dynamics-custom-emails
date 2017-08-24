using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing
{
    public interface ILexicalParser
    {
        /// <summary>
        /// Retrieve tokens
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        //List<string> Parse(string text);

        void SetText(string text);

        string CurrentToken();

        string MoveToNextToken();

        string NextToken();
    }
}
