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
        public string guigexinghao_H{ get; set; }
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
}
