using Microsoft.Office.Interop.Word;
//using Microsoft.Office.Interop.Graph;
using text.doors.Common;
using text.doors.dal;
using text.doors.Model.DataBase;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Young.Core.Common;
using text.doors.Default;
using System.Linq;
using System.Drawing;
using System.Data;
using System.Drawing.Drawing2D;
using System.IO;

namespace text.doors.Detection
{
    public partial class ExportReport : Form
    {
        private string _tempCode = "";
        private static Young.Core.Logger.ILog Logger = Young.Core.Logger.LoggerManager.Current();

        Formula formula = new Formula();


        public ExportReport(string code)
        {
            InitializeComponent();
            this._tempCode = code;
            cm_Report.SelectedIndex = 0;
        }


        private void btn_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            if (cm_Report.SelectedIndex == 0)
            {
                MessageBox.Show("请选择模板！", "请选择模板！", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                return;
            }

            Eexport(cm_Report.SelectedItem.ToString());
        }


        private void Eexport(string fileName)
        {
            try
            {
                string strResult = string.Empty;
                string strPath = System.Windows.Forms.Application.StartupPath + "\\template";
                string strFile = string.Format(@"{0}\{1}", strPath, fileName);

                FolderBrowserDialog path = new FolderBrowserDialog();
                path.ShowDialog();

                label3.Visible = true;
                if (string.IsNullOrWhiteSpace(path.SelectedPath))
                {
                    return;
                }
                btn_ok.Enabled = false;
                cm_Report.Enabled = false;
                btn_close.Enabled = false;

                string[] name = fileName.Split('.');

                string _name = name[0] + "_" + _tempCode + "." + name[1];

                var saveExcelUrl = path.SelectedPath + "\\" + _name;

                Model_dt_Settings settings = new DAL_dt_Settings().GetInfoByCode(_tempCode);

                if (settings == null)
                {
                    MessageBox.Show("未查询到相关编号!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                    return;
                }

                var dc = new Dictionary<string, string>();
                if (fileName == "门窗检验报告.doc")
                {
                    dc = GetDWDetectionReport(settings);
                }
                else if (fileName == "实验室记录.doc")
                {
                    dc = GetDetectionReport(settings, saveExcelUrl);
                }

                WordUtility wu = new WordUtility(strFile, saveExcelUrl);
                if (wu.GenerateWordByBookmarks(dc))
                {
                    label3.Visible = false;

                    MessageBox.Show("导出成功", "导出成功", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                    this.Hide();
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
                MessageBox.Show("数据出现问题，导出失败!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                this.Close();
            }
        }


        /// <summary>
        /// 获取门窗检测报告文档
        /// </summary>
        /// <param name="?"></param>
        /// <returns></returns>
        private Dictionary<string, string> GetDWDetectionReport(Model_dt_Settings settings)
        {
            Dictionary<string, string> dc = new Dictionary<string, string>();
            dc.Add("检测条件第0樘型号规格", settings.GuiGeXingHao);
            dc.Add("检测条件第0樘大气压力", settings.DaQiYaLi);
            dc.Add("检测条件第0樘委托单位", settings.WeiTuoDanWei);
            dc.Add("检测条件第0樘委托单位重复1", settings.WeiTuoDanWei);
            dc.Add("检测条件第0樘工程名称", settings.GongChengMingCheng);
            dc.Add("检测条件第0樘开启方式", settings.KaiQiFangShi);
            dc.Add("检测条件第0樘开启缝长", settings.KaiQiFengChang);
            dc.Add("检测条件第0樘当前温度", settings.DangQianWenDu);
            dc.Add("检测条件第0樘总面积", settings.ZongMianJi);
            dc.Add("检测条件第0樘最大玻璃", settings.ZuiDaBoLi);
            dc.Add("检测条件第0樘来样方式", settings.CaiYangFangShi);
            dc.Add("检测条件第0樘来样方式重复1", settings.CaiYangFangShi);
            dc.Add("检测条件第0樘样品名称", settings.YangPinMingCheng);
            dc.Add("检测条件第0樘样品名称重复1", settings.YangPinMingCheng);
            dc.Add("检测条件第0樘框扇密封", settings.KuangShanMiFang);
            dc.Add("检测条件第0樘检验数量", settings.GuiGeShuLiang);
            dc.Add("检测条件第0樘检验日期", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验日期重复1", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验日期重复2", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验编号", settings.dt_Code);
            dc.Add("检测条件第0樘检验编号重复1", settings.dt_Code);
            dc.Add("检测条件第0樘检验编号重复2", settings.dt_Code);
            dc.Add("检测条件第0樘检验编号重复3", settings.dt_Code);
            dc.Add("检测条件第0樘检验项目", settings.JianYanXiangMu);
            dc.Add("检测条件第0樘正压气密等级设计值", settings.ZhengYaQiMiDengJiSheJiZhi);
            dc.Add("检测条件第0樘负压气密等级设计值", settings.FuYaQiMiDengJiSheJiZhi);
         
            if (settings.dt_qm_Info.Count > 0)
            {
                var z_qm_level = formula.Get_Z_AirTightLevel(settings.dt_qm_Info);
                dc.Add("检测条件第0樘综合气密正压等级", z_qm_level.ToString());
                var f_qm_level = formula.Get_F_AirTightLevel(settings.dt_qm_Info);
                dc.Add("检测条件第0樘综合气密负压等级", f_qm_level.ToString());
            }
            else
            {
                dc.Add("检测条件第0樘综合气密正压等级", "--");
                dc.Add("检测条件第0樘综合气密负压等级", "--");
            }

            if (settings.dt_sm_Info.Count > 0)
            {
                var sm_level = formula.GetWaterTightLevel(settings.dt_sm_Info);
                var YL = formula.GetWaterTightPressure(settings.dt_sm_Info);

                dc.Add("检测条件第0樘水密等级", sm_level.ToString());
                //  dc.Add("检测条件第0樘水密等级设计值", sm_level.ToString());
                dc.Add("检测条件第0樘水密等级设计值", settings.ShuiMiDengJiSheJiZhi.ToString());
                dc.Add("检测条件第0樘水密保持风压", YL.ToString());
            }
            else
            {
                dc.Add("检测条件第0樘水密等级", "--");
                dc.Add("检测条件第0樘水密等级设计值", "--");
                dc.Add("检测条件第0樘水密保持风压", "--");
            }

       
            double zFc = 0, fFc = 0, zMj = 0, fMj = 0;
            if (settings.dt_qm_Info != null && settings.dt_qm_Info.Count > 0)
            {
                zFc = Math.Round(settings.dt_qm_Info.Sum(t => double.Parse(t.qm_Z_FC)) / settings.dt_qm_Info.Count, 2);
                fFc = Math.Round(settings.dt_qm_Info.Sum(t => double.Parse(t.qm_F_FC)) / settings.dt_qm_Info.Count, 2);
                zMj = Math.Round(settings.dt_qm_Info.Sum(t => double.Parse(t.qm_Z_MJ)) / settings.dt_qm_Info.Count, 2);
                fMj = Math.Round(settings.dt_qm_Info.Sum(t => double.Parse(t.qm_F_MJ)) / settings.dt_qm_Info.Count, 2);
            }

            dc.Add("检测条件第0樘正缝长渗透量", zFc.ToString());
            dc.Add("检测条件第0樘负缝长渗透量", fFc.ToString());
            dc.Add("检测条件第0樘正面积渗透量", zMj.ToString());
            dc.Add("检测条件第0樘负面积渗透量", fMj.ToString());
            dc.Add("检测条件第0樘玻璃品种", settings.BoLiPinZhong);
            dc.Add("检测条件第0樘玻璃密封", settings.BoLiMiFeng);
            dc.Add("检测条件第0樘生产单位", settings.ShengChanDanWei);
            dc.Add("检测条件第0樘负责人", settings.WeiTuoRen);
            dc.Add("检测条件第0樘镶嵌方式", settings.XiangQianFangShi);
            return dc;
        }


        #region 获取检测报告文档
        /// <summary>
        /// 获取检测报告文档
        /// </summary>
        /// <param name="?"></param>
        /// <returns></returns>
        private Dictionary<string, string> GetDetectionReport(Model_dt_Settings settings, string file)
        {
            Dictionary<string, string> dc = new Dictionary<string, string>();

            #region 基础
            dc.Add("检测条件第0樘杆件长度", settings.GanJianChangDu);
            dc.Add("实验室气压", settings.DaQiYaLi);
            dc.Add("实验室温度", settings.DangQianWenDu);
            dc.Add("集流管经", (DefaultBase._D * 1000).ToString());
            dc.Add("检测条件第0樘五金件状况", settings.WuJinJianZhuangKuang);
            dc.Add("检测条件第0樘型号规格", settings.GuiGeXingHao);
            dc.Add("检测条件第0樘大气压力", settings.DaQiYaLi);
            dc.Add("检测条件第0樘委托单位", settings.WeiTuoDanWei);
            dc.Add("检测条件第0樘委托单位重复1", settings.WeiTuoDanWei);
            dc.Add("检测条件第0樘委托单位重复2", settings.WeiTuoDanWei);
            dc.Add("检测条件第0樘委托单位重复3", settings.WeiTuoDanWei);
            dc.Add("检测条件第0樘工程名称", settings.GongChengMingCheng);
            dc.Add("检测条件第0樘工程地点", settings.GongChengDiDian);
            dc.Add("检测条件第0樘开启缝长", settings.KaiQiFengChang);
            dc.Add("检测条件第0樘开启缝长重复1", settings.KaiQiFengChang);
            dc.Add("检测条件第0樘当前温度", settings.DangQianWenDu);
            dc.Add("检测条件第0樘总面积", settings.ZongMianJi);
            dc.Add("检测条件第0樘总面积重复2", settings.ZongMianJi);
            dc.Add("检测条件第0樘最大玻璃", settings.ZuiDaBoLi);
            dc.Add("检测条件第0樘来样方式", settings.CaiYangFangShi);
            dc.Add("检测条件第0樘样品名称", settings.YangPinMingCheng);
            dc.Add("检测条件第0樘样品名称重复1", settings.YangPinMingCheng);
            dc.Add("检测条件第0樘样品名称重复2", settings.YangPinMingCheng);
            dc.Add("检测条件第0樘样品名称重复3", settings.YangPinMingCheng);
            dc.Add("检测条件第0樘框扇密封", settings.KuangShanMiFang);
            dc.Add("检测条件第0樘检验数量", settings.GuiGeShuLiang);
            dc.Add("检测条件第0樘检验编号", settings.dt_Code);
            dc.Add("检测条件第0樘检验编号重复1", settings.dt_Code);
            dc.Add("检测条件第0樘检验编号重复2", settings.dt_Code);
            dc.Add("检测条件第0樘检验编号重复3", settings.dt_Code);
            dc.Add("检测条件第0樘检验编号重复4", settings.dt_Code);
            dc.Add("检测条件第0樘检验编号重复5", settings.dt_Code);
            dc.Add("检测条件第0樘检验日期重复1", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验日期重复2", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验日期重复3", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验日期重复4", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验日期重复5", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验日期重复6", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验日期重复7", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验项目", settings.JianYanXiangMu);
            dc.Add("检测条件第0樘正压气密等级设计值", settings.ZhengYaQiMiDengJiSheJiZhi);
            dc.Add("检测条件第0樘负压气密等级设计值", settings.FuYaQiMiDengJiSheJiZhi);
            dc.Add("检测条件第0樘水密等级设计值", settings.ShuiMiDengJiSheJiZhi);
            dc.Add("检测条件第0樘玻璃厚度", settings.BoLiHouDu);
            dc.Add("检测条件第0樘玻璃品种", settings.BoLiPinZhong);
            dc.Add("检测条件第0樘玻璃密封", settings.BoLiMiFeng);
            dc.Add("检测条件第0樘抗风压等级设计值", settings.KangFengYaDengJiSheJiZhi);
            dc.Add("检测条件第0樘镶嵌方式", settings.XiangQianFangShi);


            dc.Add("检测条件第0樘单扇单锁点", settings.DanShanDanSuoDian);

            #endregion

            if (settings.dt_qm_Info.Count > 0)
            {
                #region 气密
                //检测条件第0樘综合气密等级
                var z_qm_level = formula.Get_Z_AirTightLevel(settings.dt_qm_Info);
                dc.Add("检测条件第0樘正压气密等级", z_qm_level.ToString());
                var f_qm_level = formula.Get_F_AirTightLevel(settings.dt_qm_Info);
                dc.Add("检测条件第0樘负压气密等级", f_qm_level.ToString());

                if (settings.dt_qm_Info != null && settings.dt_qm_Info.Count > 0)
                {
                    for (int i = 0; i < settings.dt_qm_Info.Count; i++)
                    {
                        if (i == 0)
                        {
                            dc.Add("气密检测第1樘总的渗透正升压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_z_zd100).ToString("#0.00"));
                            dc.Add("气密检测第1樘总的渗透正降压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_j_z_zd100).ToString("#0.00"));
                            dc.Add("气密检测第1樘总的渗透负升压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_f_zd100).ToString("#0.00"));
                            dc.Add("气密检测第1樘总的渗透负降压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_j_f_zd100).ToString("#0.00"));
                            dc.Add("气密检测第1樘附加渗透负升压150帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_f_fj150).ToString("#0.00"));
                            dc.Add("气密检测第1樘总的渗透负升压150帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_f_zd150).ToString("#0.00"));
                            dc.Add("气密检测第1樘总的渗透正升压150帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_z_zd150).ToString("#0.00"));
                            dc.Add("气密检测第1樘附加渗透正升压150帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_z_fj150).ToString("#0.00"));
                            dc.Add("气密检测第1樘附加渗透正升压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_z_fj100).ToString("#0.00"));
                            dc.Add("气密检测第1樘附加渗透正降压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_j_z_fj100).ToString("#0.00"));
                            dc.Add("气密检测第1樘附加渗透负升压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_f_fj100).ToString("#0.00"));
                            dc.Add("气密检测第1樘附加渗透负降压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_j_f_fj100).ToString("#0.00"));
                            dc.Add("流量第一樘升100附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_z_fj100)).ToString("#0.00"));
                            dc.Add("流量第一樘升150附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_z_fj150)).ToString("#0.00"));
                            dc.Add("流量第一樘负升150附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_f_fj150)).ToString("#0.00"));
                            dc.Add("流量第一樘负升100附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_f_fj100)).ToString("#0.00"));
                            dc.Add("流量第一樘负升100总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_f_zd100)).ToString("#0.00"));
                            dc.Add("流量第一樘升100总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_z_zd100)).ToString("#0.00"));
                            dc.Add("流量第一樘升150总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_z_zd150)).ToString("#0.00"));
                            dc.Add("流量第一樘负升150总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_f_zd150)).ToString("#0.00"));
                            dc.Add("流量第一樘降100总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_j_z_zd100)).ToString("#0.00"));
                            dc.Add("流量第一樘降100附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_j_z_fj100)).ToString("#0.00"));
                            dc.Add("流量第一樘负降100总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_j_f_zd100)).ToString("#0.00"));
                            dc.Add("流量第一樘负降100附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_j_f_fj100)).ToString("#0.00"));
                        }
                        if (i == 1)
                        {
                            dc.Add("气密检测第2樘总的渗透正升压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_z_zd100).ToString("#0.00"));
                            dc.Add("气密检测第2樘总的渗透正升压150帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_z_zd150).ToString("#0.00"));
                            dc.Add("气密检测第2樘总的渗透负升压150帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_f_zd150).ToString("#0.00"));
                            dc.Add("气密检测第2樘附加渗透负升压150帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_f_fj150).ToString("#0.00"));
                            dc.Add("气密检测第2樘附加渗透正升压150帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_z_fj150).ToString("#0.00"));
                            dc.Add("气密检测第2樘总的渗透正降压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_j_z_zd100).ToString("#0.00"));
                            dc.Add("气密检测第2樘总的渗透负升压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_f_zd100).ToString("#0.00"));
                            dc.Add("气密检测第2樘总的渗透负降压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_j_f_zd100).ToString("#0.00"));
                            dc.Add("气密检测第2樘附加渗透正升压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_z_fj100).ToString("#0.00"));
                            dc.Add("气密检测第2樘附加渗透正降压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_j_z_fj100).ToString("#0.00"));
                            dc.Add("气密检测第2樘附加渗透负升压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_f_fj100).ToString("#0.00"));
                            dc.Add("气密检测第2樘附加渗透负降压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_j_f_fj100).ToString("#0.00"));

                            //第二樘
                            dc.Add("流量第二樘升100附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_z_fj100)).ToString("#0.00"));
                            dc.Add("流量第二樘升150附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_z_fj150)).ToString("#0.00"));
                            dc.Add("流量第二樘负升150附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_f_fj150)).ToString("#0.00"));
                            dc.Add("流量第二樘负升100附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_f_fj100)).ToString("#0.00"));
                            dc.Add("流量第二樘负升100总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_f_zd100)).ToString("#0.00"));
                            dc.Add("流量第二樘升100总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_z_zd100)).ToString("#0.00"));
                            dc.Add("流量第二樘升150总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_z_zd150)).ToString("#0.00"));
                            dc.Add("流量第二樘负升150总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_f_zd150)).ToString("#0.00"));
                            dc.Add("流量第二樘降100总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_j_z_zd100)).ToString("#0.00"));
                            dc.Add("流量第二樘降100附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_j_z_fj100)).ToString("#0.00"));
                            dc.Add("流量第二樘负降100总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_j_f_zd100)).ToString("#0.00"));
                            dc.Add("流量第二樘负降100附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_j_f_fj100)).ToString("#0.00"));
                        }
                        if (i == 2)
                        {
                            dc.Add("气密检测第3樘总的渗透正升压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_z_zd100).ToString("#0.00"));
                            dc.Add("气密检测第3樘总的渗透正升压150帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_z_zd150).ToString("#0.00"));
                            dc.Add("气密检测第3樘总的渗透负升压150帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_f_zd150).ToString("#0.00"));
                            dc.Add("气密检测第3樘附加渗透负升压150帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_f_fj150).ToString("#0.00"));
                            dc.Add("气密检测第3樘附加渗透正升压150帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_z_fj150).ToString("#0.00"));
                            dc.Add("气密检测第3樘总的渗透正降压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_j_z_zd100).ToString("#0.00"));
                            dc.Add("气密检测第3樘总的渗透负升压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_f_zd100).ToString("#0.00"));
                            dc.Add("气密检测第3樘总的渗透负降压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_j_f_zd100).ToString("#0.00"));
                            dc.Add("气密检测第3樘附加渗透正升压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_z_fj100).ToString("#0.00"));
                            dc.Add("气密检测第3樘附加渗透正降压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_j_z_fj100).ToString("#0.00"));
                            dc.Add("气密检测第3樘附加渗透负升压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_s_f_fj100).ToString("#0.00"));
                            dc.Add("气密检测第3樘附加渗透负降压100帕时风速", double.Parse(settings.dt_qm_Info[i].qm_j_f_fj100).ToString("#0.00"));
                            //流量
                            dc.Add("流量第三樘负升100总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_f_zd100)).ToString("#0.00"));
                            dc.Add("流量第三樘升100总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_z_zd100)).ToString("#0.00"));
                            dc.Add("流量第三樘负升100附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_f_fj100)).ToString("#0.00"));
                            dc.Add("流量第三樘升100附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_z_fj100)).ToString("#0.00"));
                            dc.Add("流量第三樘升150总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_z_zd150)).ToString("#0.00"));
                            dc.Add("流量第三樘负升150总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_f_zd150)).ToString("#0.00"));
                            dc.Add("流量第三樘升150附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_z_fj150)).ToString("#0.00"));
                            dc.Add("流量第三樘负升150附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_s_f_fj150)).ToString("#0.00"));
                            dc.Add("流量第三樘负降100总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_j_f_zd100)).ToString("#0.00"));
                            dc.Add("流量第三樘降100总的", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_j_z_zd100)).ToString("#0.00"));
                            dc.Add("流量第三樘降100附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_j_z_fj100)).ToString("#0.00"));
                            dc.Add("流量第三樘负降100附加", formula.MathFlow(double.Parse(settings.dt_qm_Info[i].qm_j_f_fj100)).ToString("#0.00"));
                        }
                    }
                }
                #endregion
            }
            else
            {
                dc.Add("检测条件第0樘综合气密等级", "--");
            }
            if (settings.dt_sm_Info.Count > 0)
            {
                #region 水密
                var sm_level = formula.GetWaterTightLevel(settings.dt_sm_Info);
                dc.Add("检测条件第0樘水密等级", sm_level.ToString());

                for (int i = 0; i < settings.dt_sm_Info.Count; i++)
                {
                    if (i == 0)
                        dc.Add("检测条件第0樘水密检测方法", settings.dt_sm_Info[i].Method);


                    string[] arr = settings.dt_sm_Info[i].sm_PaDesc.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    var one = "";
                    var two = "";
                    if (arr.Length == 0)
                        continue;

                    else if (arr.Length == 1)
                        one = arr[0];

                    else if (arr.Length == 2)
                    {
                        one = arr[0];
                        two = arr[1];
                    }

                    if (two.Contains("▲") || two.Contains("●"))
                    {
                        if (i == 0)
                        {
                            //if (settings.dt_sm_Info[i].sm_Pa == "0")
                            //{
                            //    dc.Add("水密检测第1樘压力0帕状态", one);
                            //    dc.Add("水密检测第1樘压力0帕部位", two);
                            //}
                            if (settings.dt_sm_Info[i].sm_Pa == "0")
                            {
                                dc.Add("水密检测第1樘压力100帕状态", one);
                                dc.Add("水密检测第1樘压力100帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "100")
                            {
                                dc.Add("水密检测第1樘压力150帕状态", one);
                                dc.Add("水密检测第1樘压力150帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "150")
                            {
                                dc.Add("水密检测第1樘压力200帕状态", one);
                                dc.Add("水密检测第1樘压力200帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "200")
                            {
                                dc.Add("水密检测第1樘压力250帕状态", one);
                                dc.Add("水密检测第1樘压力250帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "250")
                            {
                                dc.Add("水密检测第1樘压力300帕状态", one);
                                dc.Add("水密检测第1樘压力300帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "300")
                            {
                                dc.Add("水密检测第1樘压力350帕状态", one);
                                dc.Add("水密检测第1樘压力350帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "350")
                            {
                                dc.Add("水密检测第1樘压力400帕状态", one);
                                dc.Add("水密检测第1樘压力400帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "400")
                            {
                                dc.Add("水密检测第1樘压力500帕状态", one);
                                dc.Add("水密检测第1樘压力500帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "500")
                            {
                                dc.Add("水密检测第1樘压力600帕状态", one);
                                dc.Add("水密检测第1樘压力600帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "600")
                            {
                                dc.Add("水密检测第1樘压力700帕状态", one);
                                dc.Add("水密检测第1樘压力700帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "700")
                            {
                                dc.Add("水密检测第1樘压力700帕状态", one);
                                dc.Add("水密检测第1樘压力700帕部位", two);
                            }
                            dc.Add("水密检测第1樘水密实验备注", settings.dt_sm_Info[i].sm_Remark);

                        }
                        if (i == 1)
                        {
                            //if (settings.dt_sm_Info[i].sm_Pa == "0")
                            //{
                            //    dc.Add("水密检测第2樘压力0帕状态", one);
                            //    dc.Add("水密检测第2樘压力0帕部位", two);
                            //}
                            if (settings.dt_sm_Info[i].sm_Pa == "0")
                            {
                                dc.Add("水密检测第2樘压力100帕状态", one);
                                dc.Add("水密检测第2樘压力100帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "100")
                            {
                                dc.Add("水密检测第2樘压力150帕状态", one);
                                dc.Add("水密检测第2樘压力150帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "150")
                            {
                                dc.Add("水密检测第2樘压力200帕状态", one);
                                dc.Add("水密检测第2樘压力200帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "200")
                            {
                                dc.Add("水密检测第2樘压力250帕状态", one);
                                dc.Add("水密检测第2樘压力250帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "250")
                            {
                                dc.Add("水密检测第2樘压力300帕状态", one);
                                dc.Add("水密检测第2樘压力300帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "300")
                            {
                                dc.Add("水密检测第2樘压力350帕状态", one);
                                dc.Add("水密检测第2樘压力350帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "350")
                            {
                                dc.Add("水密检测第2樘压力400帕状态", one);
                                dc.Add("水密检测第2樘压力400帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "400")
                            {
                                dc.Add("水密检测第2樘压力500帕状态", one);
                                dc.Add("水密检测第2樘压力500帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "500")
                            {
                                dc.Add("水密检测第2樘压力600帕状态", one);
                                dc.Add("水密检测第2樘压力600帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "600")
                            {
                                dc.Add("水密检测第2樘压力700帕状态", one);
                                dc.Add("水密检测第2樘压力700帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "700")
                            {
                                dc.Add("水密检测第2樘压力700帕状态", one);
                                dc.Add("水密检测第2樘压力700帕部位", two);
                            }
                            dc.Add("水密检测第2樘水密实验备注", settings.dt_sm_Info[i].sm_Remark);
                        }
                        if (i == 2)
                        {
                            //if (settings.dt_sm_Info[i].sm_Pa == "0")
                            //{
                            //    dc.Add("水密检测第3樘压力0帕状态", one);
                            //    dc.Add("水密检测第3樘压力0帕部位", two);
                            //}
                            if (settings.dt_sm_Info[i].sm_Pa == "0")
                            {
                                dc.Add("水密检测第3樘压力100帕状态", one);
                                dc.Add("水密检测第3樘压力100帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "100")
                            {
                                dc.Add("水密检测第3樘压力150帕状态", one);
                                dc.Add("水密检测第3樘压力150帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "150")
                            {
                                dc.Add("水密检测第3樘压力200帕状态", one);
                                dc.Add("水密检测第3樘压力200帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "200")
                            {
                                dc.Add("水密检测第3樘压力250帕状态", one);
                                dc.Add("水密检测第3樘压力250帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "250")
                            {
                                dc.Add("水密检测第3樘压力300帕状态", one);
                                dc.Add("水密检测第3樘压力300帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "300")
                            {
                                dc.Add("水密检测第3樘压力350帕状态", one);
                                dc.Add("水密检测第3樘压力350帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "350")
                            {
                                dc.Add("水密检测第3樘压力400帕状态", one);
                                dc.Add("水密检测第3樘压力400帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "400")
                            {
                                dc.Add("水密检测第3樘压力500帕状态", one);
                                dc.Add("水密检测第3樘压力500帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "500")
                            {
                                dc.Add("水密检测第3樘压力600帕状态", one);
                                dc.Add("水密检测第3樘压力600帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "600")
                            {
                                dc.Add("水密检测第3樘压力700帕状态", one);
                                dc.Add("水密检测第3樘压力700帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "700")
                            {
                                dc.Add("水密检测第2樘压力700帕状态", one);
                                dc.Add("水密检测第2樘压力700帕部位", two);
                            }
                            dc.Add("水密检测第3樘水密实验备注", settings.dt_sm_Info[i].sm_Remark);
                        }
                    }
                    else
                    {
                        if (i == 0)
                        {
                            if (settings.dt_sm_Info[i].sm_Pa == "0")
                            {
                                dc.Add("水密检测第1樘压力0帕状态", one);
                                dc.Add("水密检测第1樘压力0帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "100")
                            {
                                dc.Add("水密检测第1樘压力100帕状态", one);
                                dc.Add("水密检测第1樘压力100帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "150")
                            {
                                dc.Add("水密检测第1樘压力150帕状态", one);
                                dc.Add("水密检测第1樘压力150帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "200")
                            {
                                dc.Add("水密检测第1樘压力200帕状态", one);
                                dc.Add("水密检测第1樘压力200帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "250")
                            {
                                dc.Add("水密检测第1樘压力250帕状态", one);
                                dc.Add("水密检测第1樘压力250帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "300")
                            {
                                dc.Add("水密检测第1樘压力300帕状态", one);
                                dc.Add("水密检测第1樘压力300帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "350")
                            {
                                dc.Add("水密检测第1樘压力350帕状态", one);
                                dc.Add("水密检测第1樘压力350帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "400")
                            {
                                dc.Add("水密检测第1樘压力400帕状态", one);
                                dc.Add("水密检测第1樘压力400帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "500")
                            {
                                dc.Add("水密检测第1樘压力500帕状态", one);
                                dc.Add("水密检测第1樘压力500帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "600")
                            {
                                dc.Add("水密检测第1樘压力600帕状态", one);
                                dc.Add("水密检测第1樘压力600帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "700")
                            {
                                dc.Add("水密检测第1樘压力700帕状态", one);
                                dc.Add("水密检测第1樘压力700帕部位", two);
                            }
                            dc.Add("水密检测第1樘水密实验备注", settings.dt_sm_Info[i].sm_Remark);

                        }
                        if (i == 1)
                        {
                            if (settings.dt_sm_Info[i].sm_Pa == "0")
                            {
                                dc.Add("水密检测第2樘压力0帕状态", one);
                                dc.Add("水密检测第2樘压力0帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "100")
                            {
                                dc.Add("水密检测第2樘压力100帕状态", one);
                                dc.Add("水密检测第2樘压力100帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "150")
                            {
                                dc.Add("水密检测第2樘压力150帕状态", one);
                                dc.Add("水密检测第2樘压力150帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "200")
                            {
                                dc.Add("水密检测第2樘压力200帕状态", one);
                                dc.Add("水密检测第2樘压力200帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "250")
                            {
                                dc.Add("水密检测第2樘压力250帕状态", one);
                                dc.Add("水密检测第2樘压力250帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "300")
                            {
                                dc.Add("水密检测第2樘压力300帕状态", one);
                                dc.Add("水密检测第2樘压力300帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "350")
                            {
                                dc.Add("水密检测第2樘压力350帕状态", one);
                                dc.Add("水密检测第2樘压力350帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "400")
                            {
                                dc.Add("水密检测第2樘压力400帕状态", one);
                                dc.Add("水密检测第2樘压力400帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "500")
                            {
                                dc.Add("水密检测第2樘压力500帕状态", "36");
                                dc.Add("水密检测第2樘压力500帕部位", "36");
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "600")
                            {
                                dc.Add("水密检测第2樘压力600帕状态", one);
                                dc.Add("水密检测第2樘压力600帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "700")
                            {
                                dc.Add("水密检测第2樘压力700帕状态", one);
                                dc.Add("水密检测第2樘压力700帕部位", two);
                            }
                            dc.Add("水密检测第2樘水密实验备注", settings.dt_sm_Info[i].sm_Remark);
                        }
                        if (i == 2)
                        {
                            if (settings.dt_sm_Info[i].sm_Pa == "0")
                            {
                                dc.Add("水密检测第3樘压力0帕状态", one);
                                dc.Add("水密检测第3樘压力0帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "100")
                            {
                                dc.Add("水密检测第3樘压力100帕状态", one);
                                dc.Add("水密检测第3樘压力100帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "150")
                            {
                                dc.Add("水密检测第3樘压力150帕状态", one);
                                dc.Add("水密检测第3樘压力150帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "200")
                            {
                                dc.Add("水密检测第3樘压力200帕状态", one);
                                dc.Add("水密检测第3樘压力200帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "250")
                            {
                                dc.Add("水密检测第3樘压力250帕状态", one);
                                dc.Add("水密检测第3樘压力250帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "300")
                            {
                                dc.Add("水密检测第3樘压力300帕状态", one);
                                dc.Add("水密检测第3樘压力300帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "350")
                            {
                                dc.Add("水密检测第3樘压力350帕状态", one);
                                dc.Add("水密检测第3樘压力350帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "400")
                            {
                                dc.Add("水密检测第3樘压力400帕状态", one);
                                dc.Add("水密检测第3樘压力400帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "500")
                            {
                                dc.Add("水密检测第3樘压力500帕状态", "36");
                                dc.Add("水密检测第3樘压力500帕部位", "36");
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "600")
                            {
                                dc.Add("水密检测第3樘压力600帕状态", one);
                                dc.Add("水密检测第3樘压力600帕部位", two);
                            }
                            if (settings.dt_sm_Info[i].sm_Pa == "700")
                            {
                                dc.Add("水密检测第3樘压力700帕状态", one);
                                dc.Add("水密检测第3樘压力700帕部位", two);
                            }
                            dc.Add("水密检测第3樘水密实验备注", settings.dt_sm_Info[i].sm_Remark);
                        }
                    }
                }
                #endregion
            }
            else
            {
                dc.Add("检测条件第0樘水密等级", "--");
            }

            #region  缝长计算
            double zFc = 0, fFc = 0, zMj = 0, fMj = 0;
            if (settings.dt_qm_Info != null && settings.dt_qm_Info.Count > 0)
            {
                zFc = Math.Round(settings.dt_qm_Info.Sum(t => double.Parse(t.qm_Z_FC)) / settings.dt_qm_Info.Count, 2);
                fFc = Math.Round(settings.dt_qm_Info.Sum(t => double.Parse(t.qm_F_FC)) / settings.dt_qm_Info.Count, 2);
                zMj = Math.Round(settings.dt_qm_Info.Sum(t => double.Parse(t.qm_Z_MJ)) / settings.dt_qm_Info.Count, 2);
                fMj = Math.Round(settings.dt_qm_Info.Sum(t => double.Parse(t.qm_F_MJ)) / settings.dt_qm_Info.Count, 2);
            }

            dc.Add("检测条件第0樘正缝长渗透量", zFc.ToString());
            dc.Add("检测条件第0樘负缝长渗透量", fFc.ToString());
            dc.Add("检测条件第0樘正面积渗透量", zMj.ToString());
            dc.Add("检测条件第0樘负面积渗透量", fMj.ToString());
            #endregion


            return dc;
        }
        #endregion

    }
}
