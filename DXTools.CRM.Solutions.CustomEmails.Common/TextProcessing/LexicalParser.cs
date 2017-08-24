using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing
{
    public class LexicalParser : ILexicalParser
    {
        private string text =string.Empty;
        private int index = 0;
        private string currentToken = null;

        public void SetText(string text)
        {
            this.text = text;
            index = 0;
            currentToken = GetToken(0);
        }

        public string NextToken()
        {
            if (currentToken == null)
                return null;
            else
                return GetToken(index + currentToken.Length);
        }

        public string CurrentToken()
        {
            return currentToken;
        }

        private string GetToken(int startIndex)
        {
            if (startIndex >= text.Length)
                return null;

            StringBuilder tokenBuilder = new StringBuilder();
            char currentCharacter;

            int tempIndex = startIndex;

            while (tempIndex < text.Length)
            {
                currentCharacter = text[tempIndex];

                if (currentCharacter == '{'
                   || currentCharacter == '}'
                   || currentCharacter == '.'
                   || currentCharacter == '('
                   || currentCharacter == ')'
                   || currentCharacter == ','
                   || currentCharacter == '"'
                   || currentCharacter == '\'')
                {
                    if (tokenBuilder.Length > 0)
                        return tokenBuilder.ToString();
                    else
                        return currentCharacter.ToString();
                }
                else
                {
                    tokenBuilder.Append(currentCharacter);
                }
                tempIndex++;
            }

            return tokenBuilder.ToString();
        }

        public string MoveToNextToken()
        {
            if (currentToken == null)
                return null;

            string oldToken = currentToken;
            currentToken = NextToken();
            index += oldToken.Length;

            return currentToken;
        }
    }
}
