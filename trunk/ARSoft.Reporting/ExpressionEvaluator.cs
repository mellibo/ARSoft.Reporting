namespace ARSoft.Reporting
{
    using System;

    using Ciloci.Flee;

    public class ExpressionEvaluator
    {
        private IDynamicExpression compiled;

        private ExpressionContext context;

        public string Expression { get; set; }

        public object EvaluateExpression(object model)
        {
            this.context.Variables[this.ModelVariableName] = model;
            return this.compiled.Evaluate();
        }

        public void Compile(Type modelType)
        {
            this.context = new ExpressionContext();
            this.context.Variables.DefineVariable(this.ModelVariableName, modelType);

            this.compiled = this.context.CompileDynamic(this.Expression);
        }

        public string ModelVariableName { get; set; }
    }
}