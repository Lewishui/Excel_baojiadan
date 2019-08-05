using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SDZdb
{
    public class clsuserinfo
    {
        public string Order_id { get; set; }
        public string name { get; set; }
        public string password { get; set; }
        public string Btype { get; set; }
        public string denglushijian { get; set; }
        public string Createdate { get; set; }
        public string AdminIS { get; set; }
        public string jigoudaima { get; set; }
        public string userTime { get; set; }
        public string mibao { get; set; }

    }
    public class clCard_info
    {
        public string Order_id { get; set; }//=FItemID
        public string daima_gonghao { get; set; }
        public string mingcheng { get; set; }
        public string quanming { get; set; }
        public string xingbie { get; set; }
        public string minzu { get; set; }
        public string chushengriqi { get; set; }
        public string zhengjianleixing { get; set; }
        public string zhengjianhaoma { get; set; }
        public string jiatingzhuzhi { get; set; }
        public string zhengjianyouxiao { get; set; }
        public string jiguan { get; set; }
        public string shenheren { get; set; }
        public string fujian { get; set; }
        public string tupian { get; set; }
        public string CardType { get; set; }

        //PIC
        public string FTypeID { get; set; }
        public string FItemID { get; set; }
        //public string FFileName { get; set; }//=tupian
        public string FData { get; set; }
        public string FVersion { get; set; }
        public string FSaveMode { get; set; }
        public string FPage { get; set; }
        public string FEntryID { get; set; }


        public byte[] imagebytes { get; set; }

        //
        public string zhengjianyouxiaoStart { get; set; }
    }
    public class clt_Item_info
    {
        public string Order_id { get; set; }//=FItemID
        public string FItemID { get; set; }
        public string FItemClassID { get; set; }
        public string FExternID { get; set; }
        public string FNumber { get; set; }
        public string FParentID { get; set; }
        public string FLevel { get; set; }
        public string FDetail { get; set; }
        public string FName { get; set; }
        public string FUnUsed { get; set; }
        public string FBrNo { get; set; }
        public string FFullNumber { get; set; }
        public string FDiff { get; set; }
        public string FDeleted { get; set; }
        public string FShortNumber { get; set; }
        public string FFullName { get; set; }
        public string UUID { get; set; }
        public string FGRCommonID { get; set; }
        public string FSystemType { get; set; }
        public string FUseSign { get; set; }
        public string FChkUserID { get; set; }
        public string FAccessory { get; set; }
        public string FGrControl { get; set; }
        public DateTime FModifyTime { get; set; }
        public string FHavePicture { get; set; }
    }

    public class cls_order_info
    {
        public string Order_id { get; set; }//=FItemID
        public string PotNo { get; set; }
        public string DDate { get; set; }
        public string AlCnt { get; set; }
        public string Lsp { get; set; }
        public string Djzsp { get; set; }
        public string Djwd { get; set; }
        public string Fzb { get; set; }
        public string FeCnt { get; set; }
        public string SiCnt { get; set; }
        public string AlOCnt { get; set; }
        public string CaFCnt { get; set; }
        public string MgCnt { get; set; }
        public string LDYJ { get; set; }
        public string MLsp { get; set; }
        public string LPW { get; set; }


    }
    public class cls_sixzhuanjiagebiao_info
    {
        public string Order_id { get; set; }//=FItemID
        public string touxing_B { get; set; }
        public string zuzuangbizhong_C { get; set; }
        public string shapianbizhong_D { get; set; }
        public string dianpiandanjia_E { get; set; }
        public string handianpianqianjia_F { get; set; }
        public string handianpiandunjia_G { get; set; }//含垫片吨价
        public string guigexinghao_H { get; set; }
        public string guige_I { get; set; }
        public string bizhong_J { get; set; }
        public string ganjia_K { get; set; }
        public string dunjia_L { get; set; }
        public string yuanmei_M { get; set; }
        public string gongxu6_N { get; set; }
        public string yuanmei_O { get; set; }
        public string gongxu5_P { get; set; }
        public string shujin_Q { get; set; }
        public string yunfei_R { get; set; }
        public string gongxu4_S { get; set; }
        public string gongxu3_T { get; set; }
        public string gongxu2_U { get; set; }
        public string chengpinsi_V { get; set; }
        public string shunhao_W { get; set; }
        public string panyuan2_X { get; set; }
        public string gongxu1_Y { get; set; }
        public string panyuan1_Z { get; set; }
        public string shunhao_AA { get; set; }
        public string panyuan_AB { get; set; }


    }
    public class cls_GPS_info
    {
        //对应TD站点名称	站点名称	地市	区域	厂家	入场时间	现场工程师	联系电话	站点ID	站点经度	站点维度	站点地址

        public string Order_id { get; set; }//=FItemID
        public string duiying { get; set; }
        public string zhandianmingcheng { get; set; }
        public string dishi { get; set; }
        public string quyu { get; set; }
        public string changjia { get; set; }
        public string ruchangshijian { get; set; }
        public string xianchanggongchengsi { get; set; }
        public string lianxidianhua { get; set; }
        public string zhandianID { get; set; }
        public string zhandianjingdu { get; set; }
        public string zhandianweidu { get; set; }
        public string zhandiandizhi { get; set; }

    }
    public class cls_gaizaoqianjinggao_info
    {
        public string Order_id { get; set; }//=FItemID
        public string xuliehao { get; set; }
        public string wangyuanleixing { get; set; }
        public string guzangyuan { get; set; }
        public string chanshengshijian { get; set; }
        public string yuanyin { get; set; }
        public string yuanyinmiaoshu { get; set; }
        public string gaojingbianhao { get; set; }
        public string gaojingmingcheng { get; set; }


    }
    public class cls_gaizaoHOUjinggao_info
    {
        public string Order_id { get; set; }//=FItemID
        public string xuliehao { get; set; }
        public string wangyuanleixing { get; set; }
        public string guzangyuan { get; set; }
        public string chanshengshijian { get; set; }
        public string yuanyin { get; set; }
        public string yuanyinmiaoshu { get; set; }
        public string gaojingbianhao { get; set; }
        public string gaojingmingcheng { get; set; }


    }
    public class cls_zongqingdan_zhibiao_info
    {
        public string Order_id { get; set; }//=FItemID
        public string jizhanmingcheng { get; set; }
        public string shijian { get; set; }
        public string LTE_C { get; set; }
        public string LTE_D { get; set; }
        public string LTE_E { get; set; }
        public string yuanyinmiaoshu { get; set; }
        public string gaojingbianhao { get; set; }
        public string gaojingmingcheng { get; set; }


    }
    public class cls_xiangmujihuazongbiao_info
    {
        public string Order_id { get; set; }//=FItemID
        public string xuhao_A { get; set; }
        public string tiaomaneirong_B { get; set; }
        public string tuhao_C { get; set; }
        public string mingcheng_D { get; set; }
        public string caizhi_E { get; set; }
        public string shuliang_F { get; set; }
        public string danwei_G { get; set; }
        public string taoshu_H { get; set; }
        public string xiangmujiaoqi_I { get; set; }
        public string zongshuliang_J { get; set; }
        public string wuliuzhouqi_K { get; set; }
        public string zhuangpeizhouqi_L { get; set; }
        public string lingjianchengpinzhouqi_M { get; set; }
        public string shifouxuyao_N { get; set; }
        public string bianmianchulizhouqi_O { get; set; }
        public string lingjianbanchengpinzhouqi_P { get; set; }
        public string beizhu_Q { get; set; }
        public string genchuineirong_R { get; set; }
        public string genchuijiedian_S { get; set; }
        public string xiatushijian_T { get; set; }
        public string xiaruriqi_U { get; set; }
        public string xiangmubiaohao_V { get; set; }
        public string tuhao1_W { get; set; }

    }

    public class cls_Sheet0home_info
    {
        public string Order_id { get; set; }//=FItemID
        public string guanjianci_A { get; set; }
        public string sousuorenqi_B { get; set; }
        public string zaixianshangpinshu_C { get; set; }
        public string zhifuzhuanhuanlv_D { get; set; }
        public string dianjilv_E { get; set; }
        public string shangchengdianjizhanbi_F { get; set; }
        public string sousuoci_G { get; set; }
        public string anlanqi_H { get; set; }
        public string beizhu1 { get; set; }


    }
    public class cls_kucun_info
    {
        //单据类型	仓库	款号	款名	颜色	尺码	库存数量	吊牌价	性别	季节	类别	年份

        public string Order_id { get; set; }//=FItemID
        public string danjuleixing { get; set; }//=FItemID
        public string cangku { get; set; }//=FItemID
        public string kuanhao { get; set; }//=FItemID
        public string kuanming { get; set; }//=FItemID
        public string yanse { get; set; }//=FItemID
        public string cima { get; set; }//=FItemID
        public string kucunshuliang { get; set; }//=FItemID
        public string diaopaijia { get; set; }//=FItemID
        public string xingbie { get; set; }//=FItemID
        public string jijie { get; set; }//=FItemID
        public string leibie { get; set; }//=FItemID
        public string nianfen { get; set; }//=FItemID

        public string beizhu1 { get; set; }//=FItemID
        public string beizhu2 { get; set; }//=FItemID
        public string beizhu3 { get; set; }//=FItemID


    }
    public class cls_xiaoshou_info
    {
        //单据类型	店铺名称	款号	款名	颜色	尺码	数量	吊牌价	性别	季节	类别	年份
        public string Order_id { get; set; }//=FItemID
        public string danjuleixing { get; set; }//=FItemID
        public string dianpumingcheng { get; set; }//=FItemID
        public string kuanhao { get; set; }//=FItemID
        public string kuanming { get; set; }//=FItemID
        public string yanse { get; set; }//=FItemID
        public string cima { get; set; }//=FItemID
        public string shuliang { get; set; }//=FItemID
        public string diaopaijia { get; set; }//=FItemID
        public string xingbie { get; set; }//=FItemID
        public string jijie { get; set; }//=FItemID
        public string leibie { get; set; }//=FItemID
        public string nianfen { get; set; }//=FItemID


        public string beizhu1 { get; set; }//=FItemID
        public string beizhu2 { get; set; }//=FItemID
        public string beizhu3 { get; set; }//=FItemID
    }
    public class cls_diaobo_info
    {
        //款号	性别		销				销 汇总	存				存 汇总					调出				调入
        //款号	性别	尺码	公主岭家具城店	公主岭一店	公主岭专卖店	四平铁东专卖店		公主岭家具城店	公主岭一店	公主岭专卖店	四平铁东专卖店		可用				公主岭家具城店	公主岭一店	公主岭专卖店	四平铁东专卖店	

        public string Order_id { get; set; }//=FItemID
        public string kuanhao { get; set; }//=FItemID
        public string xingbie { get; set; }//=FItemID
        public string cima { get; set; }//=FItemID
        public string fendianming { get; set; }//=FItemID
        public string fendianming_xiaoshou { get; set; }//销售数量
        public string fendian_huizong { get; set; }//=销售汇总
        public string kucunming { get; set; }//=FItemID
        public string kucun_shengyu { get; set; }//=库存剩余
        public string kucun_huizong { get; set; }//=FItemID
        public string diaopaijia { get; set; }//=FItemID

        public string keyong_kucun { get; set; }//=可用剩余
        public string diaochudianpuming { get; set; }//=FItemID
        public string diaoru { get; set; }//=FItemID
        public string beizhu1 { get; set; }//=FItemID
        public string beizhu2 { get; set; }//=FItemID
        public string beizhu3 { get; set; }//=FItemID

        //
        public string kuanming { get; set; }//=FItemID
        public string yanse { get; set; }//=FItemID
 
        public string jijie { get; set; }//=FItemID
        public string leibie { get; set; }//=FItemID
        public string nianfen { get; set; }//=FItemID
        public string huizong { get; set; }//=FItemID
    }
}
