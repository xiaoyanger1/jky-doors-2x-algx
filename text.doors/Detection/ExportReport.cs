﻿using Microsoft.Office.Interop.Word;
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
using System.Collections;

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

                btn_ok.Enabled = false;
                cm_Report.Enabled = false;
                btn_close.Enabled = false;


                string[] name = fileName.Split('.');

                string _name = name[0] + "_" + _tempCode + "." + name[1];

                var saveExcelUrl = path.SelectedPath + "\\" + _name;

                Model_dt_Settings settings = new DAL_dt_Settings().Getdt_SettingsResByCode(_tempCode);

                if (settings == null)
                {
                    MessageBox.Show("未查询到相关编号");
                    this.Close();
                    return;
                }

                var dc = new Dictionary<string, string>();
                if (fileName == "门窗检验报告.doc")
                {
                    dc = GetDWDetectionReport(settings);
                }
                else if (fileName == "试验室记录.doc")
                {
                    dc = GetDetectionReport(settings);
                }
                else if (fileName == "现场报告（1樘）.doc")
                {
                    dc = GetTong1(settings);
                }
                else if (fileName == "现场报告（2樘）.doc")
                {
                    dc = GetTong2(settings);
                }
                else if (fileName == "现场报告（3樘）.doc")
                {
                    dc = GetTong3(settings);
                }

                WordUtility wu = new WordUtility(strFile, saveExcelUrl);
                if (wu.GenerateWordByBookmarks(dc))
                {
                    MessageBox.Show("导出成功", "导出成功",
                             MessageBoxButtons.OK,
                             MessageBoxIcon.None,
                             MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.ServiceNotification
                            );
                    this.Hide();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("数据出现问题，导出失败！");
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
            if (settings.JianYanXiangMu == "气密性能检测" || settings.JianYanXiangMu == "气密性能及水密性能检测")
            {
                var qm_zy_level = Get_QMLevel(settings, 1);
                dc.Add("检测条件第0樘综合气密等级", qm_zy_level.ToString());

                var qm_fy_level = Get_QMLevel(settings, 2);
                dc.Add("检测条件气密性能负压属国标等级", qm_fy_level.ToString());

            }
            else { dc.Add("检测条件第0樘综合气密等级", "--"); }



            if (settings.JianYanXiangMu == "水密性能检测" || settings.JianYanXiangMu == "气密性能及水密性能检测")
            {
                var sm_level = Get_SMLevel(settings);
                var YL = Get_SMYL(settings);

                dc.Add("检测条件第0樘水密等级", sm_level.ToString());
                dc.Add("检测条件第0樘水密等级设计值", sm_level.ToString());
                dc.Add("检测条件第0樘水密保持风压", YL.ToString());
            }
            else
            {
                dc.Add("检测条件第0樘水密等级", "--");
                dc.Add("检测条件第0樘水密等级设计值", "--");
                dc.Add("检测条件第0樘水密保持风压", "--");
            }
            dc.Add("检测条件第0樘正缝长渗透量", qm_z_FC.ToString());
            dc.Add("检测条件第0樘负缝长渗透量", qm_f_FC.ToString());
            dc.Add("检测条件第0樘正面积渗透量", qm_z_MJ.ToString());
            dc.Add("检测条件第0樘负面积渗透量", qm_f_MJ.ToString());
            dc.Add("检测条件第0樘玻璃品种", settings.BoLiPinZhong);
            dc.Add("检测条件第0樘玻璃密封", settings.BoLiMiFeng);
            dc.Add("检测条件第0樘生产单位", settings.ShengChanDanWei);
            dc.Add("检测条件第0樘负压气密等级设计值", settings.FuYaQiMiDengJiSheJiZhi);
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
        private Dictionary<string, string> GetDetectionReport(Model_dt_Settings settings)
        {
            Dictionary<string, string> dc = new Dictionary<string, string>();


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
            dc.Add("检测条件第0樘检验日期重复1", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验日期重复2", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验日期重复3", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验日期重复4", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验日期重复5", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验项目", settings.JianYanXiangMu);
            dc.Add("检测条件第0樘正压气密等级设计值", settings.ZhengYaQiMiDengJiSheJiZhi);
            dc.Add("检测条件第0樘水密等级设计值", settings.ShuiMiDengJiSheJiZhi);
            dc.Add("检测条件第0樘玻璃厚度", settings.BoLiHouDu);
            dc.Add("检测条件第0樘玻璃品种", settings.BoLiPinZhong);
            dc.Add("检测条件第0樘玻璃密封", settings.BoLiMiFeng);
            dc.Add("检测条件第0樘负压气密等级设计值", settings.FuYaQiMiDengJiSheJiZhi);
            dc.Add("检测条件第0樘镶嵌方式", settings.XiangQianFangShi);

            if (settings.JianYanXiangMu == "气密性能检测" || settings.JianYanXiangMu == "气密性能及水密性能检测")
            {
                #region 气密
                var qm_zy_level = Get_QMLevel(settings, 1);
                dc.Add("检测条件第0樘综合气密等级", qm_zy_level.ToString());

                var qm_fy_level = Get_QMLevel(settings, 2);
                dc.Add("检测条件气密性能负压属国标等级", qm_fy_level.ToString());

                if (settings.dt_qm_Info != null && settings.dt_qm_Info.Count > 0)
                {
                    Formula formula = new Formula();
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
            if (settings.JianYanXiangMu == "水密性能检测" || settings.JianYanXiangMu == "气密性能及水密性能检测")
            {
                #region 水密
                var sm_level = Get_SMLevel(settings);
                dc.Add("检测条件第0樘水密等级", sm_level.ToString());

                if (settings.dt_sm_Info != null && settings.dt_sm_Info.Count > 0)
                {
                    for (int i = 0; i < settings.dt_sm_Info.Count; i++)
                    {
                        string[] arr = settings.dt_sm_Info[i].sm_PaDesc.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        var one = "";
                        var two = "";
                        if (arr.Length == 0)
                        {
                            continue;
                        }
                        else if (arr.Length == 1)
                        {
                            one = arr[0];
                        }
                        else if (arr.Length == 2) { one = arr[0]; two = arr[1]; }
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
                                dc.Add("水密检测第1樘压力500帕状态", "36");
                                dc.Add("水密检测第1樘压力500帕部位", "36");
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
            else { dc.Add("检测条件第0樘水密等级", "--"); }
            dc.Add("检测条件第0樘正缝长渗透量", qm_z_FC.ToString());
            dc.Add("检测条件第0樘负缝长渗透量", qm_f_FC.ToString());
            dc.Add("检测条件第0樘正面积渗透量", qm_z_MJ.ToString());
            dc.Add("检测条件第0樘负面积渗透量", qm_f_MJ.ToString());

            dc.Add("检测条件第0樘水密检测方法", "--法");
            return dc;
        }

        #endregion

        #region 分樘号

        private Dictionary<string, string> GetTong1(Model_dt_Settings settings)
        {
            Dictionary<string, string> dc = new Dictionary<string, string>();
            dc.Add("检测条件第0樘委托人", settings.WeiTuoRen);
            dc.Add("检测条件第0樘委托单位", settings.WeiTuoDanWei);
            dc.Add("检测条件第0樘委托编号", settings.WeiTuoBianHao);
            dc.Add("检测条件第0樘工程名称", settings.GongChengMingCheng);
            dc.Add("检测条件第0樘开启方式", settings.KaiQiFangShi);
            dc.Add("检测条件第0樘样品名称", settings.YangPinMingCheng);
            dc.Add("检测条件第0樘委托日期", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验编号", settings.dt_Code);
            dc.Add("检测条件第0樘玻璃品种", settings.BoLiPinZhong);
            dc.Add("检测条件第0樘玻璃密封", settings.BoLiMiFeng);
            dc.Add("检测条件第0樘生产单位", settings.ShengChanDanWei);
            dc.Add("检测条件第0樘镶嵌方式", settings.XiangQianFangShi);
            dc.Add("检测条件第0樘框扇密封", settings.KuangShanMiFang);
            dc.Add("检测条件第0樘检验日期", DateTime.Now.ToString("yyyy-MM-dd"));
            dc.Add("检测条件第0樘检验日期1", DateTime.Now.ToString("yyyy-MM-dd"));

            dc.Add("检测条件第0樘型号规格", settings.GuiGeXingHao);
            //试件位置
            if (settings.JianYanXiangMu == "水密性能检测" || settings.JianYanXiangMu == "气密性能及水密性能检测")
            {
                dc.Add("检测条件第0樘第一樘水密性能", settings.dt_sm_Info?.Find(t => t.info_DangH == "第1樘")?.sm_Pa ?? "0");
            }
            else
            {
                dc.Add("检测条件第0樘第一樘水密性能", "--");
            }
            if (settings.JianYanXiangMu == "气密性能检测" || settings.JianYanXiangMu == "气密性能及水密性能检测")
            {
                dc.Add("检测条件第0樘正缝长渗透量", settings.dt_qm_Info?.Find(t => t.info_DangH == "第1樘")?.qm_Z_FC ?? "0");
                dc.Add("检测条件第0樘正面积渗透量", settings.dt_qm_Info?.Find(t => t.info_DangH == "第1樘")?.qm_Z_MJ ?? "0");
            }
            else
            {
                dc.Add("检测条件第0樘正缝长渗透量", "--");
                dc.Add("检测条件第0樘正面积渗透量", "--");
            }
            return dc;
        }


        private Dictionary<string, string> GetTong2(Model_dt_Settings settings)
        {
            Dictionary<string, string> dc = new Dictionary<string, string>();
            dc.Add("检测条件第0樘委托人", settings.WeiTuoRen);
            dc.Add("检测条件第0樘委托单位", settings.WeiTuoDanWei);
            dc.Add("检测条件第0樘委托编号", settings.WeiTuoBianHao);
            dc.Add("检测条件第0樘工程名称", settings.GongChengMingCheng);
            dc.Add("检测条件第0樘开启方式", settings.KaiQiFangShi);
            dc.Add("检测条件第0樘样品名称", settings.YangPinMingCheng);
            dc.Add("检测条件第0樘委托日期", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验编号", settings.dt_Code);
            dc.Add("检测条件第0樘玻璃品种", settings.BoLiPinZhong);
            dc.Add("检测条件第0樘玻璃密封", settings.BoLiMiFeng);
            dc.Add("检测条件第0樘生产单位", settings.ShengChanDanWei);
            dc.Add("检测条件第0樘镶嵌方式", settings.XiangQianFangShi);
            dc.Add("检测条件第0樘框扇密封", settings.KuangShanMiFang);
            dc.Add("检测条件第0樘检验日期", DateTime.Now.ToString("yyyy-MM-dd"));
            dc.Add("检测条件第0樘检验日期1", DateTime.Now.ToString("yyyy-MM-dd"));

            //dc.Add("检测条件第0樘资料编号", settings);
            dc.Add("检测条件第0樘型号规格", settings.GuiGeXingHao);

            if (settings.JianYanXiangMu == "水密性能检测" || settings.JianYanXiangMu == "气密性能及水密性能检测")
            {
                dc.Add("检测条件第0樘第一樘水密性能", settings.dt_sm_Info?.Find(t => t.info_DangH == "第1樘")?.sm_Pa ?? "0");
                dc.Add("检测条件第0樘第二樘水密性能", settings.dt_sm_Info?.Find(t => t.info_DangH == "第2樘")?.sm_Pa ?? "0");
            }
            else
            {
                dc.Add("检测条件第0樘第一樘水密性能", "--");
                dc.Add("检测条件第0樘第二樘水密性能", "--");
            }
            if (settings.JianYanXiangMu == "气密性能检测" || settings.JianYanXiangMu == "气密性能及水密性能检测")
            {
                dc.Add("检测条件第1樘正缝长渗透量", settings.dt_qm_Info?.Find(t => t.info_DangH == "第1樘")?.qm_Z_FC ?? "0");
                dc.Add("检测条件第1樘正面积渗透量", settings.dt_qm_Info?.Find(t => t.info_DangH == "第1樘")?.qm_Z_MJ ?? "0");
                dc.Add("检测条件第2樘正缝长渗透量", settings.dt_qm_Info?.Find(t => t.info_DangH == "第2樘")?.qm_Z_FC ?? "0");
                dc.Add("检测条件第2樘正面积渗透量", settings.dt_qm_Info?.Find(t => t.info_DangH == "第2樘")?.qm_Z_MJ ?? "0");
            }
            else
            {
                dc.Add("检测条件第1樘正缝长渗透量", "--");
                dc.Add("检测条件第1樘正面积渗透量", "--");
                dc.Add("检测条件第2樘正缝长渗透量", "--");
                dc.Add("检测条件第2樘正面积渗透量", "--");
            }
            return dc;
        }


        private Dictionary<string, string> GetTong3(Model_dt_Settings settings)
        {
            Dictionary<string, string> dc = new Dictionary<string, string>();
            dc.Add("检测条件第0樘委托人", settings.WeiTuoRen);
            dc.Add("检测条件第0樘委托单位", settings.WeiTuoDanWei);
            dc.Add("检测条件第0樘委托编号", settings.WeiTuoBianHao);
            dc.Add("检测条件第0樘工程名称", settings.GongChengMingCheng);
            dc.Add("检测条件第0樘开启方式", settings.KaiQiFangShi);
            dc.Add("检测条件第0樘样品名称", settings.YangPinMingCheng);
            dc.Add("检测条件第0樘委托日期", settings.JianYanRiQi);
            dc.Add("检测条件第0樘检验编号", settings.dt_Code);
            dc.Add("检测条件第0樘玻璃品种", settings.BoLiPinZhong);
            dc.Add("检测条件第0樘玻璃密封", settings.BoLiMiFeng);
            dc.Add("检测条件第0樘生产单位", settings.ShengChanDanWei);
            dc.Add("检测条件第0樘镶嵌方式", settings.XiangQianFangShi);
            dc.Add("检测条件第0樘框扇密封", settings.KuangShanMiFang);
            dc.Add("检测条件第0樘检验日期", DateTime.Now.ToString("yyyy-MM-dd"));
            dc.Add("检测条件第0樘委托日期1", DateTime.Now.ToString("yyyy-MM-dd"));
            dc.Add("检测条件第0樘型号规格", settings.GuiGeXingHao);
            //dc.Add("检测条件第0樘资料编号", settings);

            if (settings.JianYanXiangMu == "水密性能检测" || settings.JianYanXiangMu == "气密性能及水密性能检测")
            {
                dc.Add("检测条件第0樘第一樘水密性能", settings.dt_sm_Info?.Find(t => t.info_DangH == "第3樘")?.sm_Pa ?? "0");
                var sm_level = Get_SMLevel(settings);
                dc.Add("检测条件第0樘第一樘水密性能分级", sm_level.ToString());
            }
            else
            {
                dc.Add("检测条件第0樘第一樘水密性能", "--");
                dc.Add("检测条件第0樘第一樘水密性能分级", "--");
            }
            if (settings.JianYanXiangMu == "气密性能检测" || settings.JianYanXiangMu == "气密性能及水密性能检测")
            {
                dc.Add("检测条件第0樘正缝长渗透量", settings.dt_qm_Info?.Find(t => t.info_DangH == "第3樘")?.qm_Z_FC ?? "0");
                dc.Add("检测条件第0樘正面积渗透量", settings.dt_qm_Info?.Find(t => t.info_DangH == "第3樘")?.qm_Z_MJ ?? "0");
                var qm_zy_level = Get_QMLevel(settings, 1);
                dc.Add("检测条件第0樘综合气密等级", qm_zy_level.ToString());
            }
            else
            {
                dc.Add("检测条件第0樘正缝长渗透量", "--");
                dc.Add("检测条件第0樘正面积渗透量", "--");
                dc.Add("检测条件第0樘综合气密等级", "--");
            }
            return dc;
        }
        #endregion



        /// <summary>
        /// 获取水密压力
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private int Get_SMYL(Model_dt_Settings settings)
        {

            int qmValue = 0;
            if (settings != null && settings.dt_qm_Info.Count > 0)
            {
                if (settings.dt_sm_Info.Count == 3)
                {
                    List<int> list = new List<int>() { int.Parse(settings.dt_sm_Info[0].sm_Pa.ToString()), int.Parse(settings.dt_sm_Info[1].sm_Pa.ToString()), int.Parse(settings.dt_sm_Info[2].sm_Pa.ToString()) };
                    list.Sort();

                    int min = list[0], intermediate = list[1], max = list[2];

                    int minlevel = GetQMLevelList.Find(t => t.value == min).level,
                        intermediatelevel = GetQMLevelList.Find(t => t.value == intermediate).level,
                        maxlevel = GetQMLevelList.Find(t => t.value == max).level;

                    if ((maxlevel - intermediatelevel) > 2)
                    {
                        max = GetQMLevelList.Find(t => t.level == (intermediatelevel + 2)).value;
                    }

                    qmValue = (min + intermediate + max) / 3;
                }
                else
                {
                    for (int i = 0; i < settings.dt_sm_Info.Count; i++)
                    {
                        if (string.IsNullOrWhiteSpace(settings.dt_sm_Info[0].sm_Pa))
                        {
                            qmValue = 0;
                            break;
                        }
                        qmValue += int.Parse(settings.dt_sm_Info[0].sm_Pa.ToString());
                    }
                    qmValue = qmValue / settings.dt_sm_Info.Count;
                }
            }
            return qmValue;
        }


        /// <summary>
        /// 气密属性
        /// </summary>
        public double qm_z_FC = 0, qm_f_FC = 0, qm_z_MJ = 0, qm_f_MJ = 0;
        /// <summary>
        /// 水密属性
        /// </summary>
        public int sm_value = 999;

        #region  计算


        /// <summary>
        /// 获取水密等级
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private int Get_SMLevel(Model_dt_Settings settings)
        {

            int qmValue = 0;
            if (settings != null && settings.dt_qm_Info.Count > 0)
            {
                if (settings.dt_sm_Info.Count == 3)
                {
                    List<int> list = new List<int>() { int.Parse(settings.dt_sm_Info[0].sm_Pa.ToString()), int.Parse(settings.dt_sm_Info[1].sm_Pa.ToString()), int.Parse(settings.dt_sm_Info[2].sm_Pa.ToString()) };
                    list.Sort();

                    int min = list[0], intermediate = list[1], max = list[2];

                    int minlevel = GetQMLevelList.Find(t => t.value == min).level,
                        intermediatelevel = GetQMLevelList.Find(t => t.value == intermediate).level,
                        maxlevel = GetQMLevelList.Find(t => t.value == max).level;

                    if ((maxlevel - intermediatelevel) > 2)
                    {
                        max = GetQMLevelList.Find(t => t.level == (intermediatelevel + 2)).value;
                    }

                    qmValue = (min + intermediate + max) / 3;
                }
                else
                {
                    for (int i = 0; i < settings.dt_sm_Info.Count; i++)
                    {
                        if (string.IsNullOrWhiteSpace(settings.dt_sm_Info[0].sm_Pa))
                        {
                            qmValue = 0;
                            break;
                        }
                        qmValue += int.Parse(settings.dt_sm_Info[0].sm_Pa.ToString());
                    }
                    qmValue = qmValue / settings.dt_sm_Info.Count;
                }
            }
            return GetSMLevel(qmValue);
        }

        /// <summary>
        /// 获取不标准的等级
        /// 范式 气密正负缝长平均值等级 与 气密正负压缝长平均值等级 最大的最次
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private int Get_QMLevel(Model_dt_Settings settings, int type)
        {
            qm_z_FC = 0; qm_f_FC = 0; qm_z_MJ = 0; qm_f_MJ = 0;
            for (int i = 0; i < settings.dt_qm_Info.Count; i++)
            {
                qm_z_FC += double.Parse(settings.dt_qm_Info[i].qm_Z_FC.ToString());
                qm_f_FC += double.Parse(settings.dt_qm_Info[i].qm_F_FC.ToString());

                qm_z_MJ += double.Parse(settings.dt_qm_Info[i].qm_Z_MJ.ToString());
                qm_f_MJ += double.Parse(settings.dt_qm_Info[i].qm_F_MJ.ToString());
            }
            qm_z_FC = Math.Round(qm_z_FC / settings.dt_qm_Info.Count, 2);
            qm_f_FC = Math.Round(qm_f_FC / settings.dt_qm_Info.Count, 2);

            qm_z_MJ = Math.Round(qm_z_MJ / settings.dt_qm_Info.Count, 2);
            qm_f_MJ = Math.Round(qm_f_MJ / settings.dt_qm_Info.Count, 2);

            return GetQM_MaxLevel(qm_z_FC, qm_f_FC, qm_z_MJ, qm_f_MJ, type);
        }

        /// <summary>
        /// 获取气密最大等级
        /// </summary>
        /// <param name="fc"></param>
        /// <param name="mj"></param>
        /// <returns></returns>
        public int GetQM_MaxLevel(double qm_z_FC, double qm_f_FC, double qm_z_MJ, double qm_f_MJ, int type)
        {
            int level_z_FJ = 0, level_f_FJ = 0, level_z_MJ = 0, level_f_MJ = 0;
            level_z_FJ = GetFCLevel(qm_z_FC);
            level_f_FJ = GetFCLevel(qm_f_FC);
            level_z_MJ = GetMJLevel(qm_z_MJ);
            level_f_MJ = GetMJLevel(qm_f_MJ);
            int[] arr = null;
            if (type == 1)
            {
                arr = new int[] { level_z_FJ, level_z_MJ };
            }
            else if (type == 2)
            {
                arr = new int[] { level_f_FJ, level_f_MJ };
            }
            ArrayList list = new ArrayList(arr);
            list.Sort();
            return Convert.ToInt32(list[0]);

        }

        /// <summary>
        /// 获取缝长分级
        /// </summary>
        /// <returns></returns>
        public int GetFCLevel(double value)
        {
            int res = 0;
            if (4 >= value && value > 3.5)
            {
                res = 1;
            }
            else if (3.5 >= value && value > 3.0)
            {
                res = 2;
            }
            else if (3.0 >= value && value > 2.5)
            {
                res = 3;
            }
            else if (2.5 >= value && value > 2.0)
            {
                res = 4;
            }
            else if (2.0 >= value && value > 1.5)
            {
                res = 5;
            }
            else if (1.5 >= value && value > 1.0)
            {
                res = 6;
            }
            else if (1.0 >= value && value > 0.5)
            {
                res = 7;
            }
            else if (value <= 0.5)
            {
                res = 8;
            }
            return res;
        }

        /// <summary>
        /// 获取面积分级
        /// </summary>
        /// <returns></returns>
        public int GetMJLevel(double value)
        {
            int res = 0;
            if (12 >= value && value > 10.5)
            {
                res = 1;
            }
            else if (10.5 >= value && value > 9.0)
            {
                res = 2;
            }
            else if (9.0 >= value && value > 7.5)
            {
                res = 3;
            }
            else if (7.5 >= value && value > 6.0)
            {
                res = 4;
            }
            else if (6.0 >= value && value > 4.5)
            {
                res = 5;
            }
            else if (4.5 >= value && value > 3.0)
            {
                res = 6;
            }
            else if (3.0 >= value && value > 1.5)
            {
                res = 7;
            }
            else if (value <= 1.5)
            {
                res = 8;
            }
            return res;
        }

        /// <summary>
        /// 获取水密分级
        /// </summary>
        /// <returns></returns>
        public int GetSMLevel(int value)
        {
            int res = 0;
            if (value >= 100 && value < 150)
            {
                res = 1;
            }
            else if (value >= 150 && value < 250)
            {
                res = 2;
            }
            else if (value >= 250 && value < 350)
            {
                res = 3;
            }
            else if (value >= 300 && value < 500)
            {
                res = 4;
            }
            else if (value >= 500 && value < 700)
            {
                res = 5;
            }
            else if (value >= 700)
            {
                res = 6;
            }
            return res;
        }

        public static List<DataDict> GetQMLevelList
        {
            get
            {
                List<DataDict> dictList = new List<DataDict>();
                dictList.Add(new DataDict() { value = 0, level = 1 });
                dictList.Add(new DataDict() { value = 0, level = 1 });
                dictList.Add(new DataDict() { value = 100, level = 2 });
                dictList.Add(new DataDict() { value = 150, level = 3 });
                dictList.Add(new DataDict() { value = 200, level = 4 });
                dictList.Add(new DataDict() { value = 250, level = 5 });
                dictList.Add(new DataDict() { value = 300, level = 6 });
                dictList.Add(new DataDict() { value = 350, level = 7 });
                dictList.Add(new DataDict() { value = 400, level = 8 });
                dictList.Add(new DataDict() { value = 500, level = 9 });
                dictList.Add(new DataDict() { value = 600, level = 10 });
                dictList.Add(new DataDict() { value = 700, level = 11 });
                return dictList;
            }
        }
        #endregion
    }



    /// <summary>
    /// 气密等级字典
    /// </summary>
    public class DataDict
    {
        public int level { get; set; }
        public int value { get; set; }
    }
}
