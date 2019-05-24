using System;
using System.ComponentModel;

namespace KeLi.ExcelMerge.App.Models
{
    /// <summary>
    /// 标题
    /// </summary>
    public class TestFirst
    {
        /// <summary>
        /// 主数据建筑业态
        /// </summary>
        [Description("主数据建筑业态")]
        public string MainBusinessType { get; set; }

        /// <summary>
        /// 总建筑面积
        /// </summary>
        [Span(3)]
        [Description("总建筑面积")]
        public string ToltalArea { get; set; }

        /// <summary>
        /// 可租赁面积(㎡)
        /// </summary>
        [Span(3)]
        [Description("可租赁面积")]
        public string LeaseArea { get; set; }

        /// <summary>
        /// 电梯数
        /// </summary>
        [Span(2)]
        [Description("电梯数")]
        public string ElevatorNum { get; set; }

        /// <summary>
        /// 房间名称
        /// </summary>
        [Description("房间名称")]
        public string RoomName { get; set; }

        /// <summary>
        /// 计容建筑面积
        /// </summary>
        [Description("计容建筑面积")]
        public string JrArea { get; set; }

        /// <summary>
        /// 计容系数
        /// </summary>
        [Description("计容系数")]
        public string JrFactor { get; set; }

        /// <summary>
        /// 是否装修
        /// </summary>
        [Description("是否装修")]
        public string IsDecorate { get; set; }

        /// <summary>
        /// 上下水条件
        /// </summary>
        [Description("上下水条件")]
        public string WaterCondition { get; set; }
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
        public TestSecond(params object[] strs)
        {
            MainBusinessType = strs[0].ToString();
            ToltalAreaTotal = Convert.ToDouble(strs[1]);
            ToltalAreaEarth = Convert.ToDouble(strs[2]);
            ToltalAreaUnder = Convert.ToDouble(strs[3]);
            LeaseAreaTotal = Convert.ToDouble(strs[4]);
            LeaseAreaEarth = Convert.ToDouble(strs[5]);
            LeaseAreaUnder = Convert.ToDouble(strs[6]);
            ElevatorNumPassenger = Convert.ToInt32(strs[7]);
            ElevatorNumFreight = Convert.ToInt32(strs[8]);
            RoomName = strs[9].ToString();
            JrArea = Convert.ToDouble(strs[10]);
            JrFactor = Convert.ToDouble(strs[11]);
            IsDecorate = Convert.ToBoolean(strs[12]);
            WaterCondition = strs[13].ToString();
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
        public double ToltalAreaTotal { get; set; }

        /// <summary>
        /// 总建筑面积.地上
        /// </summary>
        [Description("地上")]
        public double ToltalAreaEarth { get; set; }

        /// <summary>
        /// 总建筑面积.地下
        /// </summary>
        [Description("地下")]
        public double ToltalAreaUnder { get; set; }

        /// <summary>
        /// 可租赁面积.总面积
        /// </summary>
        [Description("总面积")]
        public double LeaseAreaTotal { get; set; }

        /// <summary>
        /// 可租赁面积.地上
        /// </summary>
        [Description("地上")]
        public double LeaseAreaEarth { get; set; }

        /// <summary>
        /// 可租赁面积.地下
        /// </summary>
        [Description("地下")]
        public double LeaseAreaUnder { get; set; }

        /// <summary>
        /// 电梯数.客梯
        /// </summary>
        [Description("客梯")]
        [Reference("MainBusinessType")]
        public int ElevatorNumPassenger { get; set; }

        /// <summary>
        /// 电梯数.货梯
        /// </summary>
        [Description("货梯")]
        public int ElevatorNumFreight { get; set; }

        /// <summary>
        /// 房间名称
        /// </summary>
        [Description("房间名称")]
        public string RoomName { get; set; }

        /// <summary>
        /// 计容建筑面积
        /// </summary>
        [Description("计容建筑面积")]
        public double JrArea { get; set; }

        /// <summary>
        /// 计容系数
        /// </summary>
        [Description("计容系数")]
        public double JrFactor { get; set; }

        /// <summary>
        ///是否装修
        /// </summary>
        [Description("是否装修")]
        public bool IsDecorate { get; set; }

        /// <summary>
        /// 上下水条件
        /// </summary>
        [Description("上下水条件")]
        public string WaterCondition { get; set; }
    }
}
