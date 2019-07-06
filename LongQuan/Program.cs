using DBEN.DBI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LongQuan
{
    class Program
    {
        static void Main(string[] args)
        {

            var dt = ExcelHelper.ImportExceltoDt(@"F:\LiXingZhou\龙泉\群发短信处理程序\LongQuan\LongQuan\112014年李行周.xls", 0, "1");


            for (int i=0;i<dt.Rows.Count;i++)
            {
                Console.WriteLine("\r\n/******"+dt.Rows[i]["期数"]+dt.Rows[i]["姓名"]+(i+1) + "电话:"+dt.Rows[i]["电话"] + "*********/\r\n");

                Console.WriteLine($"【北京龙泉寺周末心义工】“七年缘聚 初心未变”{dt.Rows[i]["姓名"]} 师兄您好，您初次结识周末心义工是在2014年{dt.Rows[i]["期数"]}活动。2016年北京龙泉寺心义工已陪我们走过第七个年头，周末心义工以此殊胜因缘发心举办“北京龙泉寺周末心义工七周年祈福庆典法会”，特意为您准备了专场吉祥普佛，并与众师兄一起见证我们走过的那段美好时光同时憧憬丰富多彩的未来，在修学佛法的道路上同心同愿同行。\r\n活动日期：2017年1月8日\r\n报名链接：http://t.cn/RIIsIUi \r\n您也可回复【姓名+性别+电话+是否有龙泉寺皈依证】直接报名。 ");
            }

            Console.ReadLine();
        }
    }
}
