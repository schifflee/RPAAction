namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 单元格-聚焦
    /// </summary>
    public class Range_Active : ExcelAction
    {
        public Range_Active(string wbPath = null, string wsName = null, string range = null)
            : base(wbPath, wsName)
        {
            this.range = range;
            Run();
        }

        protected override void Action()
        {
            base.Action();
            wb.Activate();
            ws.Select();
            if (range != null && (!range.Equals("")))
            {
                app.Range[range].Select();
            }
        }
    }
}
