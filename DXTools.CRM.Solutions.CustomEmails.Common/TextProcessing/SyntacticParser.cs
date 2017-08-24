using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Expressions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing
{
    public class SyntacticParser : ISyntacticParser
    {
        public ILexicalParser LexicalParser { get; private set; }

        public SyntacticParser(ILexicalParser lexicalParser)
        {
            this.LexicalParser = lexicalParser;
        }

        public Symbol Parse(string text)
        {
            this.LexicalParser.SetText(text);
            return GetDynamicExpression();
        }

        private DynamicExpression GetDynamicExpression()
        {
            DynamicExpression dynamicExpression = new DynamicExpression();

            string currentToken = this.LexicalParser.CurrentToken();          
            string nextToken = this.LexicalParser.NextToken();

            if (nextToken == "(")
            {
                dynamicExpression.ChildExpressions.Add(GetFunctionExpression());
            }
            else if (nextToken == ".")
            {
                dynamicExpression.ChildExpressions.Add(GetAttributeExpression());
            }
            else
            {
                dynamicExpression.ChildExpressions.Add(new AttributeExpression(currentToken));
            }

            return dynamicExpression;
        }

        private Function GetFunctionExpression()
        {
            Function function = new Function(this.LexicalParser.CurrentToken());
            this.LexicalParser.MoveToNextToken();
            string currentToken = this.LexicalParser.MoveToNextToken();

            while (currentToken != ")" && currentToken != null)
            {
                if (currentToken == "\"" || currentToken == "'")
                {
                    currentToken = this.LexicalParser.MoveToNextToken();

                    StringBuilder textParameterBuilder = new StringBuilder();
                    while (currentToken != "\"" && currentToken != "'" && currentToken != ")" && currentToken != null)
                    {
                        textParameterBuilder.Append(currentToken);
                        currentToken = this.LexicalParser.MoveToNextToken();
                    }

                    FunctionParameter functionParameter = new FunctionParameter(textParameterBuilder.ToString());
                    function.ChildExpressions.Add(functionParameter);
                }
                else
                {
                    string nextToken = this.LexicalParser.NextToken();
                    if (!string.IsNullOrWhiteSpace(currentToken) 
                        && (nextToken == "," || nextToken == "(" || nextToken == ")" || nextToken == ".") )
                    {
                        DynamicExpression expression = GetDynamicExpression();
                        function.ChildExpressions.Add(expression);
                    }
                }
                currentToken = this.LexicalParser.MoveToNextToken();
            }
            return function;

        }

        private AttributeExpression GetAttributeExpression()
        {
            AttributeExpression attribute = new AttributeExpression(this.LexicalParser.CurrentToken());

            string nextToken = this.LexicalParser.MoveToNextToken();
            if (nextToken == ".")
            {
                this.LexicalParser.MoveToNextToken();
                attribute.ChildExpressions.Add(GetAttributeExpression());
            }

            return attribute;
        }
    }
}
