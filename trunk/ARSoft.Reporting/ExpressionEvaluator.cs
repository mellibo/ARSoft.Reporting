namespace ARSoft.Reporting
{
    using System;
    using System.Collections.Generic;
    using System.Dynamic;

    using Ciloci.Flee;

    public class ExpressionEvaluator
    {
        private readonly Dictionary<string, object> variables;

        private IDynamicExpression compiled;

        private ExpressionContext context;

        public ExpressionEvaluator()
        {
            this.variables = new Dictionary<string, object>();
        }

        public string Expression { get; set; }

        public object EvaluateExpression(object model)
        {
            this.context.Variables[this.ModelVariableName] = model;
            foreach (var variable in this.variables)
            {
                this.context.Variables[variable.Key] = variable.Value;
            }

            return this.compiled.Evaluate();
        }

        public void Compile(Type modelType)
        {
            this.context = new ExpressionContext();
            this.context.Variables.DefineVariable(this.ModelVariableName, modelType);
            foreach (var variable in this.variables)
            {
                this.context.Variables.DefineVariable(variable.Key, variable.Value == null ? typeof(object) : variable.Value.GetType());
            }

            this.compiled = this.context.CompileDynamic(this.Expression);
        }

        public string ModelVariableName { get; set; }

        public void AddVariable(string name, object value)
        {
            if (!this.variables.ContainsKey(name))
            {
                this.variables.Add(name, value);
            }
        }
    }
}