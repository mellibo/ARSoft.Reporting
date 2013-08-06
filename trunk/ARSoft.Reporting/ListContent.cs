namespace ARSoft.Reporting
{
    using System;
    using System.Collections;

    public class ListContent : ReportContent
    {
        public string DataSourceExpression { get; set; }

        private ReportContentContainer contents;

        public ListContent()
        {
            this.contents = new ReportContentContainer();
        }

        public override void Write(IReportWriter excelWriter, object datasource)
        {
            IEnumerable internalDatasource = null;
            if (!string.IsNullOrWhiteSpace(DataSourceExpression))
            {
                ExpressionEvaluator evaluator = new ExpressionEvaluator();
                evaluator.ModelVariableName = Guid.NewGuid().ToString();
                evaluator.Expression = DataSourceExpression;
                evaluator.Compile(datasource.GetType());
                internalDatasource = evaluator.EvaluateExpression(datasource) as IEnumerable;
            }
            else
            {
                internalDatasource = datasource as IEnumerable;
            }
            foreach (var item in internalDatasource)
            {
                foreach (var reportContent in contents.Contents)
                {
                    reportContent.Write(excelWriter, item);
                }
                excelWriter.NewRow();
            }
        }

        public ReportContentContainer Content
        {
            get
            {
                return this.contents;
            }
        }
    }
}