namespace ARSoft.Reporting
{
    using Ciloci.Flee;

    public class ExpressionContent : ReportContent
    {
        public string Expression { get; set; }
        
        public override string GetText(object model)
        {
            var context = new ExpressionContext();
            context.Variables.DefineVariable("model", model.GetType());

            var compiled = context.CompileDynamic(Expression);
            context.Variables["model"] = model;
            var value = compiled.Evaluate();
            return value == null ? string.Empty : value.ToString();
        }
    }
}