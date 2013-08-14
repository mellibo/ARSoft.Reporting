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
            var internalDatasource = this.GetInternalDatasource(datasource);

            foreach (var item in internalDatasource)
            {
                foreach (var reportContent in contents.Contents)
                {
                    reportContent.Write(excelWriter, item);
                }

                if (Direction == DirectionEnum.Vertical) excelWriter.CrLf();
            }
        }

        private IEnumerable GetInternalDatasource(object datasource)
        {
            IEnumerable internalDatasource;
            if (!string.IsNullOrWhiteSpace(this.DataSourceExpression))
            {
                var evaluator = new ExpressionEvaluator();
                evaluator.ModelVariableName = Guid.NewGuid().ToString();
                evaluator.Expression = this.DataSourceExpression;
                evaluator.Compile(datasource.GetType());
                internalDatasource = evaluator.EvaluateExpression(datasource) as IEnumerable;
            }
            else
            {
                internalDatasource = datasource as IEnumerable;
            }

            return internalDatasource;
        }

        public ReportContentContainer Content
        {
            get
            {
                return this.contents;
            }
        }

        public DirectionEnum Direction { get; set; }

    }
}