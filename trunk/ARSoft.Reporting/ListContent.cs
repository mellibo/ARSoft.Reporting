namespace ARSoft.Reporting
{
    using System;
    using System.Collections;
    using System.Collections.Generic;

    public class ListContent : ReportContent
    {
        public string DataSourceExpression { get; set; }

        private ReportContentContainer contents;

        private Dictionary<Type, string> itemTemplates;

        public ListContent()
        {
            this.itemTemplates = new Dictionary<Type, string>();
            this.contents = new ReportContentContainer();
        }

        public override void Write(IReportWriter writer, object datasource)
        {
            int listX;
            if (this.X.HasValue)
            {
                listX = this.X.Value - 1;
            }
            else
            {
                listX = writer.LastX;
            }

            writer.SetCurrentX(listX);

            if (this.Y.HasValue) writer.SetCurrentY(this.Y.Value);
            var internalDatasource = this.GetInternalDatasource(datasource);
            writer.Context.ItemNumber = 1;
            var itemTemplate = this.itemTemplates.ContainsKey(writer.GetType())
                                   ? this.itemTemplates[writer.GetType()]
                                   : null;
            foreach (var item in internalDatasource)
            {
                writer.StartRow(itemTemplate);
                foreach (var reportContent in contents.Contents)
                {
                    reportContent.Write(writer, item);
                }

                if (Direction == DirectionEnum.Vertical)
                {
                    writer.CrLf();
                    writer.SetCurrentX(listX);
                }

                writer.Context.ItemNumber++;
            }
        }

        private IEnumerable GetInternalDatasource(object datasource)
        {
            IEnumerable internalDatasource;
            if (!string.IsNullOrWhiteSpace(this.DataSourceExpression))
            {
                var evaluator = new ExpressionEvaluator();
                evaluator.ModelVariableName = "model";
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

        public Dictionary<Type, string> ItemTemplates
        {
            get
            {
                return this.itemTemplates;
            }
        }
    }
}