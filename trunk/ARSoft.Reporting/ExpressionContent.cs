namespace ARSoft.Reporting
{
    public class ExpressionContent : ReportContent
    {
        private readonly ExpressionEvaluator expressionEvaluator;

        public ExpressionContent()
        {
            expressionEvaluator = new ExpressionEvaluator();
            expressionEvaluator.ModelVariableName = "model";
        }

        public string Expression
        {
            get
            {
                return expressionEvaluator.Expression;
            }

            set
            {
                expressionEvaluator.Expression = value;
            }
        }

        public override void Write(IReportWriter writer, object datasource)
        {
            expressionEvaluator.AddVariable("Context", writer.Context);
            expressionEvaluator.Compile(datasource.GetType());
            var value = expressionEvaluator.EvaluateExpression(datasource);
            var valueString = value == null ? string.Empty : value.ToString();
            writer.WriteTextElement(X, Y, valueString);
        }
    }
}