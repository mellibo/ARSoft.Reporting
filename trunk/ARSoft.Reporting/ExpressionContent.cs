namespace ARSoft.Reporting
{
    using Ciloci.Flee;

    public class ExpressionContent : ReportContent
    {
        public string Expression { get; set; }
        
        public string EvaluateExpression(object model)
        {
            var context = new ExpressionContext();
            context.Variables.DefineVariable("model", model.GetType());

            var compiled = context.CompileDynamic(Expression);
            context.Variables["model"] = model;
            var value = compiled.Evaluate();
            return value == null ? string.Empty : value.ToString();
        }

        public override void Write(ExcelWriter excelWriter, object datasource)
        {
            excelWriter.WriteTextElement(X, Y, this.EvaluateExpression(datasource));
        }
    }
}