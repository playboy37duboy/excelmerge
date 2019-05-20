using System.ComponentModel;
namespace KeLi.ExcelMerge.App
{
    /// <summary>
    /// 标题
    /// </summary>
    public class TestFirst
    {
        /// <summary>
        /// 主数据建筑业态
        /// </summary>
        [Span(1)]
        [Description("主数据建筑业态")]
        public string MainBusinessType { get; set; }

        /// <summary>
        /// 总建筑面积(㎡)
        /// </summary>
        [Span(3)]
        [Description("总建筑面积(㎡)")]
        public string ToltalArea { get; set; }

        /// <summary>
        /// 可租赁面积(㎡)
        /// </summary>
        [Span(3)]
        [Description("可租赁面积(㎡)")]
        public string LeaseArea { get; set; }

        /// <summary>
        /// 电梯数
        /// </summary>
        [Span(2)]
        [Description("电梯数")]
        public string ElevatorNum { get; set; }

        /// <summary>
        /// 主数据建筑业态
        /// </summary>
        [Span(1)]
        [Description("主数据建筑业态1")]
        public string MainBusinessType1 { get; set; }

        /// <summary>
        /// 主数据建筑业态
        /// </summary>
        [Span(1)]
        [Description("主数据建筑业态2")]
        public string MainBusinessType2 { get; set; }


        /// <summary>
        /// 主数据建筑业态
        /// </summary>
        [Span(1)]
        [Description("主数据建筑业态3")]
        public string MainBusinessType3 { get; set; }

        /// <summary>
        /// 主数据建筑业态
        /// </summary>
        [Span(1)]
        [Description("主数据建筑业态4")]
        public string MainBusinessType4 { get; set; }

        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态5")]
        public string ElevatorNumFreight5 { get; set; }


        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态6")]
        public string ElevatorNumFreight6 { get; set; }


        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态7")]
        public string ElevatorNumFreight7 { get; set; }

        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态8")]
        public string ElevatorNumFreight8 { get; set; }


        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态9")]
        public string ElevatorNumFreight9 { get; set; }


        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态10")]
        public string ElevatorNumFreight10 { get; set; }

        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态11")]
        public string ElevatorNumFreight11 { get; set; }

        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态12")]
        public string ElevatorNumFreight12 { get; set; }
    }

    /// <summary>
    /// 数据
    /// </summary>
    public class TestSecond
    {
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="strs"></param>
        public TestSecond(params string[] strs)
        {
            MainBusinessType = strs[0];
            ToltalAreaTotal = strs[1];
            ToltalAreaEarth = strs[2];
            ToltalAreaUnder = strs[3];
            LeaseAreaTotal = strs[4];
            LeaseAreaEarth = strs[5];
            LeaseAreaUnder = strs[6];
            ElevatorNumPassenger = strs[7];
            ElevatorNumFreight = strs[8];
            ElevatorNumFreight1 = strs[9];
            ElevatorNumFreight2 = strs[10];
            ElevatorNumFreight3 = strs[11];
            ElevatorNumFreight4 = strs[12];
            ElevatorNumFreight5 = strs[13];
            ElevatorNumFreight6 = strs[14];
            ElevatorNumFreight7 = strs[15];
            ElevatorNumFreight8 = strs[16];
            ElevatorNumFreight9 = strs[17];
            ElevatorNumFreight10 = strs[18];
            ElevatorNumFreight11 = strs[19];
            ElevatorNumFreight12 = strs[20];
        }

        /// <summary>
        /// 主数据建筑业态.主数据建筑业态
        /// </summary>
        [Description("主数据建筑业态")]
        public string MainBusinessType { get; set; }

        /// <summary>
        /// 总建筑面积(㎡).主数据建筑业态
        /// </summary>
        [Description("总面积")]
        public string ToltalAreaTotal { get; set; }

        /// <summary>
        /// 总建筑面积(㎡).地上
        /// </summary>
        [Description("地上")]
        public string ToltalAreaEarth { get; set; }

        /// <summary>
        /// 总建筑面积(㎡).地下
        /// </summary>
        [Description("地下")]
        public string ToltalAreaUnder { get; set; }

        /// <summary>
        /// 可租赁面积(㎡).总面积
        /// </summary>
        [Description("总面积")]
        public string LeaseAreaTotal { get; set; }

        /// <summary>
        /// 可租赁面积(㎡).地上
        /// </summary>
        [Description("地上")]
        public string LeaseAreaEarth { get; set; }

        /// <summary>
        /// 可租赁面积(㎡).地下
        /// </summary>
        [Description("地下")]
        public string LeaseAreaUnder { get; set; }

        /// <summary>
        /// 电梯数.客梯
        /// </summary>
        [Description("客梯")]
        [Reference("MainBusinessType")]
        public string ElevatorNumPassenger { get; set; }

        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("货梯")]
        public string ElevatorNumFreight { get; set; }

        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态1")]
        public string ElevatorNumFreight1 { get; set; }


        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态2")]
        public string ElevatorNumFreight2 { get; set; }


        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态3")]
        public string ElevatorNumFreight3 { get; set; }

        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态4")]
        public string ElevatorNumFreight4 { get; set; }

        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态5")]
        public string ElevatorNumFreight5 { get; set; }


        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态6")]
        public string ElevatorNumFreight6 { get; set; }


        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态7")]
        public string ElevatorNumFreight7 { get; set; }

        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态8")]
        public string ElevatorNumFreight8 { get; set; }


        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态9")]
        public string ElevatorNumFreight9 { get; set; }


        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态10")]
        public string ElevatorNumFreight10 { get; set; }

        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态11")]
        public string ElevatorNumFreight11 { get; set; }

        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("主数据建筑业态12")]
        public string ElevatorNumFreight12 { get; set; }

    }
}
