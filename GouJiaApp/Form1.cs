using Aliyun.OSS;
using GouJiaApp.Utils;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace GouJiaApp
{
    public partial class Form1 : Form
    {
        static string accessKeyId = "LTAIJr4Fi6XdibZ0";
        static string accessKeySecret = "sjvR3TxUl3wmd9t3teWI638ryLQMc9";
        static string endPoint = "oss-cn-hangzhou.aliyuncs.com";
        static string bucketName = "idiplugin";

        List<string> fucaiKeys;
        List<string> zhucaiKeys;
        List<string> ruanzhuangKeys;
        List<string> zigouKeys;
        //List<string> chuguiKeys;
        //List<string> tatamiKeys;
        //List<string> yimaoguiKeys;

        List<string> fucaiOrder;
        List<string> zhucaiOrder;
        List<string> ruanzhuangOrder;
        List<string> zigouOrder;
        List<string> daidingOrder;
        //List<string> tatamiOrder;
        //List<string> yimaoguiOrder;

        List<string> hilightKeys;

        Dictionary<string, double> Q20kg_JBL;

        double sum0015D, sum0019D, sum0023D,sum0026D, sum0028D, sum0030D, sum0032D, sum0034D, sum0036D, sum0039D, sum0041D, sum0043D, sum0045D, sum0047D, sum0049D,
            sum0053D, sum0058D = 0;
        private MD5 md5;
        public Form1()
        {
            InitializeComponent();

            md5 = new MD5CryptoServiceProvider();

            fucaiKeys = new List<string>();
            //string[] temp = "地漏、瓷砖、乳胶漆、中央空调、地暖、给排水管、强弱电线管、木工材料、木工板材、石膏板".Split('、');
            StringBuilder keysSB = new StringBuilder();
            keysSB.Append("大金、格力、VRV、");
            keysSB.Append("智能系统、");
            keysSB.Append("德国Vaillant威能、");
            keysSB.Append("地漏、");
            keysSB.Append("嘉宝莉、多乐士、乳胶漆、");
            keysSB.Append("墙砖铺贴、地砖铺贴、东鹏");
            string[] temp = keysSB.ToString().Split('、');
            fucaiKeys.AddRange(temp);

            zhucaiKeys = new List<string>();
            keysSB.Clear();
            keysSB.Append("丰华、箭牌、卫浴五金专供、汉斯格雅、杜拉维特、科勒、卡丽、欧路莎、");
            keysSB.Append("摩恩、");
            keysSB.Append("名门、");
            keysSB.Append("方太、热水器、小厨宝、净水器、垃圾处理器、蒸箱、烤箱、微波炉、冰吧、");
            keysSB.Append("名族、集成吊顶安装、好易点、晾衣机、LED吸顶灯、");
            keysSB.Append("西蒙、欧普、调光变压器、");
            keysSB.Append("康E-梦天、背景造型、硬包背景、实木嵌条、金迪、垭口套、户内门专供、背景专供、");
            keysSB.Append("玄关柜、移门衣柜、开门衣柜、展示柜、储藏柜、书柜、电视柜、柜五金、");
            keysSB.Append("成品移门、");
            keysSB.Append("大自然、汉界、安信、地板、");
            keysSB.Append("索图、洗衣柜、洗衣柜专供、");
            keysSB.Append("钰尚壁纸、雅琪诺壁纸、壁纸、");
            keysSB.Append("戈兰迪、华纳、");
            keysSB.Append("橱柜、榻榻米、衣帽间");

            temp = keysSB.ToString().Split('、');
            zhucaiKeys.AddRange(temp);

            ruanzhuangKeys = new List<string>();
            keysSB.Clear();
            keysSB.Append("太阳姊妹、");
            keysSB.Append("床品四件套、");
            keysSB.Append("美登斯、主布、遮光布、世典、");
            keysSB.Append("灯美、吊灯、吸顶灯、艺术王朝、台灯、落地灯、");
            keysSB.Append("床垫、");
            keysSB.Append("冰箱、洗衣机、电视机、分体柜机、分体挂机、");
            keysSB.Append("地毯、");
            keysSB.Append("香阁丽莱、第法、新波普、新宏基、松可、高帆、顾家、丽晶、玛润奇、融峰、富渥得、凯沃、诺丁山、巴比特、木-帅可、木-松可、科鲁德、易式形态、凯沃、威格纳");

            temp = keysSB.ToString().Split('、');
            ruanzhuangKeys.AddRange(temp);

            zigouKeys = new List<string>();
            keysSB.Clear();
            keysSB.Append("百得、");
            keysSB.Append("德宝利莱、门槛石、窗台板、门套角柱、[深色石材]、浅色大理石、");
            keysSB.Append("淋浴房、");
            keysSB.Append("腻子找平、柔性防水、导线、石膏板、多层板、杉木集成板、胶粉、胶浆、石膏线、银桥、");
            keysSB.Append("网络线、电话线、T5灯管");

            temp = keysSB.ToString().Split('、');
            zigouKeys.AddRange(temp);

            //chuguiKeys = new List<string>();
            //chuguiKeys.Add("橱柜");

            //tatamiKeys = new List<string>();
            //tatamiKeys.Add("榻榻米");

            //yimaoguiKeys = new List<string>();
            //yimaoguiKeys.Add("衣帽间");

            fucaiOrder = new List<string>();
            temp = "地漏、乳胶漆、墙面漆、底漆、基础漆、瓷片、抛釉砖、防滑砖、仿古砖、瓷砖、中央空调、提升泵、外机、内机、地暖、壁挂炉".Split('、');
            fucaiOrder.AddRange(temp);

            zhucaiOrder = new List<string>();
            temp = ("台面、水晶之恋、香格里拉、立木星、春意俏、小荷初露、福寿绵绵、橱柜、调味篮、防滑垫、刀叉盘、碗篮、锅篮、帮滑轨、水槽、龙头、台盆、台下盆、马桶、坐便器、座便器、小便、蹲便、拖把池、马桶喷枪、花洒、淋浴凳、"
                + "置物架、卫浴四件套、厕纸架、浴巾架、毛巾架、单杆巾架、毛巾环、五金配件、三角阀、双用角阀、地排下水管、入墙式墙排下水管、墙排下水管、翻盖式脸盆下水、软管、波纹管、洗衣机柜、洗衣机柜龙头、浴室柜、镜柜、镜框、"
                + "防雾膜、晾衣机、单开门、双移门、单移门、厨房移门、书房移门、单开移门、单开玻璃门、单边套、半边套、踢脚线、哑口套、免漆、双扇移门、门连套、整门、背景墙、暗门、墙背景木饰面、隔断、隔断木饰面、隐门、门锁、把手、碰吸、门吸、闭门器、合页、"
                + "挖手、轨道、吊轮、吊轨、三节轨、地板、实木地板、实木复合地板、玄关、玄关柜、展示柜、开门柜、开门衣柜、移门柜、移门衣柜、衣帽间、衣柜、书柜、电视柜、"
                + "阳台柜、阳台吊柜、阳台地柜、酒柜、嵌条、储藏柜、储物柜、非标、榻榻米、隔板、射灯、筒灯、阅读灯、调光变压器、吸顶灯、指示灯、支架、LED、t5、灯带、小夜灯、感应灯、开关、插座、空白盖板、"
                + "吊顶、集成吊顶、浴霸、换气扇、凉霸、集成照明、油烟机、灶具、消毒柜、热水器、微波炉、小厨宝、净水器、垃圾处理器、壁纸、硬包").Split('、');
            zhucaiOrder.AddRange(temp);

            ruanzhuangOrder = new List<string>();
            temp = ("三人沙发、三人位沙发、三位沙发、3人沙发、双人沙发、双人位沙发、二人沙发、双位沙发、2人沙发、单人沙发、单人位沙发、1人沙发、单人左扶手、单人右扶手、无扶手单人沙发、转角沙发、L沙发、L沙发（左躺）、L沙发（右躺）、沙发、家具、休闲椅、单椅、脚踏、茶几、边几、角几、圆几、大方几、小方几、椭圆几、"
                + "吧台、电视柜、五斗柜、酒柜、餐桌、圆桌、餐台、餐椅、餐边柜、备餐柜、床、排骨架、床尾凳、床头柜、高低铺、衣柜、书桌、书台、写字台、书椅、扶手椅、椅子、"
                + "凳、榻、书架、书柜、玄关桌、玄关柜、边柜、鞋柜、柜子、梳妆凳、妆椅、妆台、梳妆镜、妆镜、屏风、储物柜、装饰柜、升降椅、吊灯、吸顶灯、壁灯、台灯、灯具、落地灯、窗帘、布拉帘、布百叶、地毯、挂画、装饰画、床品、"
                + "床垫、冰箱、电冰箱、洗衣机、电视、烤箱、分体挂机、分体柜机").Split('、');
            ruanzhuangOrder.AddRange(temp);

            zigouOrder = new List<string>();
            temp = ("止逆阀、淋浴房、浴室柜台面、台面、门套柱脚、脚座、布朗布、大理石、门槛石、新埃及米黄、挡水条、T5灯管、t5  14W 支架、不锈钢条、电线、网线、双绞线、电缆、电话线、BV2.5、水管、线管、石膏板、多层板、石材板、集成板、防水浆料、腻子、砂浆、瓷砖胶、界面剂、石膏").Split('、');
            zigouOrder.AddRange(temp);

            daidingOrder = new List<string>();
            daidingOrder.Add("橱柜");

            hilightKeys = new List<string>();
            temp = ("墙面漆、底漆、踢脚线、把手、顶线-封板").Split('、');
            hilightKeys.AddRange(temp);

            
        }

        private void idiTextBox_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fd = new OpenFileDialog())
            {
                fd.Filter = "Excel文件|*.xlsx;*.xls";
                DialogResult res = fd.ShowDialog();

                if (res == DialogResult.OK)
                {
                    idiTextBox.Text = fd.FileName;
                }
            }
        }

        private void pathTextBox_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                DialogResult res = fbd.ShowDialog();
                if (res == DialogResult.OK)
                {
                    pathTextBox.Text = fbd.SelectedPath;
                }
            }
        }

        private void confirmButton_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(idiTextBox.Text))
            {
                MessageBox.Show("请选择IDI导出明细表！");
                return;
            }

            if (String.IsNullOrEmpty(pathTextBox.Text))
            {
                MessageBox.Show("请选择保存路径！");
                return;
            }

            if (String.IsNullOrEmpty(nameTextBox.Text))
            {
                MessageBox.Show("请填写保存文件名！");
                return;
            }

            OssClient client = new OssClient(endPoint, accessKeyId, accessKeySecret);
            try
            {
                if (client.DoesObjectExist(bucketName, "物料单/" + nameTextBox.Text + ".xlsx"))
                {
                    MessageBox.Show("服务器上已存在：" + nameTextBox.Text + ".xls。请重新命名！");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            Q20kg_JBL = new Dictionary<string, double>();

            generate();

            try
            {
                if (!client.DoesBucketExist(bucketName))
                {
                    MessageBox.Show("bucket idiplugin 不存在！");
                    return;
                }
                string savePath = pathTextBox.Text + @"\" + nameTextBox.Text + ".xlsx";
                client.PutObject(bucketName, "物料单/" + nameTextBox.Text + ".xlsx", savePath);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            MessageBox.Show("完成");
        }

        private void generate()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("请确认Excel已正确安装！");
                return;
            }

            string idiFilePath = idiTextBox.Text;
            if (IsOpened(idiFilePath))
            {
                MessageBox.Show(idiFilePath + "正在被使用中，请先关闭！");
                return;
            }

            Workbook idiWB = xlApp.Workbooks.Open(idiFilePath);
            Worksheet idiWS = null;
            try
            {
                idiWS = idiWB.Sheets["3.预算文件"];
            }
            catch
            {
                MessageBox.Show("所选的idi明细表不包含《3.预算文件》表单，请确认明细表是否正确！");
                return;
            }


            Workbook wb = xlApp.Workbooks.Add();
            Worksheet ws6 = wb.Sheets[1];
            //Worksheet ws8 = wb.Sheets.Add();
            //Worksheet ws7 = wb.Sheets.Add();
            //Worksheet ws6 = wb.Sheets.Add();
            Worksheet ws5 = wb.Sheets.Add();
            Worksheet ws4 = wb.Sheets.Add();
            Worksheet ws3 = wb.Sheets.Add();
            Worksheet ws2 = wb.Sheets.Add();
            Worksheet ws1 = wb.Sheets.Add();

            ws1.Name = "构家物料及物流方案";
            ws2.Name = "辅材包";
            ws3.Name = "主材包";
            ws4.Name = "软装包";
            ws5.Name = "城运商自购";
            //ws6.Name = "橱柜柜体";
            //ws7.Name = "榻榻米";
            //ws8.Name = "衣帽柜";
            ws6.Name = "待定";


            string savePath = pathTextBox.Text + @"\" + nameTextBox.Text + ".xlsx";

            string errorId = "";
            try
            {
                #region 分包

                initWorkSheet(ws2, "辅材包物料清单");
                initWorkSheet(ws3, "主材包物料清单");
                initWorkSheet(ws4, "软装包物料清单");
                initWorkSheet(ws5, "城运商自购物料清单");
                initWorkSheet(ws6, "橱柜柜体物料清单");
                //initWorkSheet(ws7, "榻榻米物料清单");
                //initWorkSheet(ws8, "衣帽柜物料清单");
                //initWorkSheet(ws9, "待定物料清单");

                //找出所有需要筛选的编号
                List<string> criterias = new List<string>();
                int usedRows = idiWS.UsedRange.Rows.Count;
                progressBar.Maximum = usedRows;
                for (int i = 7; i < usedRows; i++)
                {
                    string id = Convert.ToString(idiWS.Cells[3][i].Value);
                    if (!String.IsNullOrEmpty(id) && Regex.IsMatch(id, @"^[a-zA-Z0-9-]+$") && id.Length == 14 && slashCount(id) == 2 && !criterias.Contains(id))
                    {
                        criterias.Add(id);
                    }
                    ProgressBarController.setValue(progressBar, i);
                }
                progressBar.Value = 0;

                //遍历每个编号，统计数据
                progressBar.Maximum = criterias.Count;
                Range usedRange = idiWS.UsedRange;
                int curRow2 = 4;
                int curRow3 = 4;
                int curRow4 = 4;
                int curRow5 = 4;
                int curRow6 = 4;
                //int curRow7 = 3;
                //int curRow8 = 3;
                //int curRow9 = 3;

                List<Material> materials2 = new List<Material>();
                List<Material> materials3 = new List<Material>();
                List<Material> materials4 = new List<Material>();
                List<Material> materials5 = new List<Material>();
                List<Material> materials6 = new List<Material>();
                List<Material> materials7 = new List<Material>();
                List<Material> materials8 = new List<Material>();
                List<Material> materials9 = new List<Material>();

                foreach (string criteria in criterias)
                {
                    errorId = criteria;
                    #region 统计总工程量
                    //找求总工程量
                    usedRange.AutoFilter(3,
                        criteria,
                        Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd,
                        Type.Missing, true);
                    Range filteredRange = usedRange.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible, Type.Missing);
                    double sum = 0;
                    StringBuilder roomsSB = new StringBuilder();
                    string idiName = null;
                    string idiUnit = null;
                    foreach (Range c in filteredRange.Rows)
                    {
                        //排除表头
                        if (c.Row >= 5)
                        {
                            //筛选工程名称
                            string newName = c.Columns[4].Value;
                            if (newName != null)
                            {
                                if (idiName == null)
                                    idiName = newName;
                                else if (!newName.Equals(idiName) && newName.Contains("[") && newName.Contains("]"))
                                    idiName = newName;
                            }

                            if (idiUnit == null)
                                idiUnit = c.Columns[5].Value;

                            sum += Convert.ToDouble(c.Columns[6].Value);

                            string room = null;
                            if (c.Columns[2].Value2 != null)
                            {
                                room = Convert.ToString(c.Columns[2].Value2);
                            }

                            if (!String.IsNullOrEmpty(room) && !"0".Equals(room) && !roomsSB.ToString().Contains(room))
                            {
                                if (roomsSB.Length == 0)
                                    roomsSB.Append(room);
                                else
                                    roomsSB.Append("\n").Append(room);
                            }
                        }
                    }

                    if (criteria.Equals("WMZ19-014-0026"))
                        sum0026D = sum;
                    else if (criteria.Equals("WMZ19-014-0028"))
                        sum0028D = sum;
                    else if (criteria.Equals("WMZ19-014-0030"))
                        sum0030D = sum;
                    else if (criteria.Equals("WMZ19-014-0032"))
                        sum0032D = sum;
                    else if (criteria.Equals("WMZ19-014-0034"))
                        sum0034D = sum;
                    else if (criteria.Equals("WMZ19-014-0036"))
                        sum0036D = sum;
                    else if (criteria.Equals("WMZ19-014-0039"))
                        sum0039D = sum;
                    else if (criteria.Equals("WMZ19-014-0041"))
                        sum0041D = sum;
                    else if (criteria.Equals("WMZ19-014-0043"))
                        sum0043D = sum;
                    else if (criteria.Equals("WMZ19-014-0045"))
                        sum0045D = sum;
                    else if (criteria.Equals("WMZ19-014-0047"))
                        sum0047D = sum;
                    else if (criteria.Equals("WMZ19-014-0049"))
                        sum0049D = sum;
                    else if (criteria.Equals("WMZ19-014-0053"))
                        sum0053D = sum;
                    else if (criteria.Equals("WMZ19-014-0058"))
                        sum0058D = sum;
                   
                    #endregion

                    int wsNum = 0;
                    //通过web service 获得详细信息
                    try
                    {
                        HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create("http://ges.goujiawang.com/matter/getMatterByCodeV2?code=" + criteria);
                        //HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create("http://rges.goujiawang.com/matter/getMatterByCode?code=" + criteria);
                        request.Timeout = 10000;
                        request.ContentType = "application/json; charset=utf-8";

                        WebResponse response = request.GetResponse();

                        Stream streamWeb = response.GetResponseStream();
                        StringBuilder jsonSB = new StringBuilder("");

                        using (StreamReader reader = new StreamReader(streamWeb))
                        {
                            while (!reader.EndOfStream)
                            {
                                jsonSB.Append(reader.ReadLine());
                            }
                        }

                        //if (!String.IsNullOrEmpty(idiName))
                        //{
                        //    if (!idiName.Contains("台面") && !idiName.Contains("不锈钢水槽") && !idiName.Contains("摩恩") && idiName.Contains("橱柜"))
                        //        wsNum = 6;
                        //    else if (idiName.Contains("榻榻米"))
                        //        wsNum = 7;
                        //    else if (idiName.Contains("衣帽间"))
                        //        wsNum = 8;
                        //}

                        JObject jObj = JObject.Parse(jsonSB.ToString());
                        if (!JTokenIsNullOrEmpty(jObj["result"]))
                        {
                            //string matName = (string)jObj["result"]["name"];
                            //wsNum = getPackageWorkSheetNum(matName);

                            if (wsNum == 0)
                            {
                                if (JTokenIsNullOrEmpty(jObj["result"]["useNatureText"]))
                                    wsNum = 6;
                                else
                                    wsNum = getPackageWorkSheetNumByUseNature((string)jObj["result"]["useNatureText"]);
                            }

                            switch (wsNum)
                            {
                                case 2:
                                    addRowToList(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials2);
                                    break;
                                case 3:
                                    addRowToList(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials3);
                                    break;
                                case 4:
                                    addRowToList(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials4);
                                    break;
                                case 5:
                                    addRowToList(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials5);
                                    break;
                                case 6:
                                    addRowToList(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials6);
                                    break;
                                //case 7:
                                //    //insertRow(ws6, curRow6, criteria, roomsSB.ToString(), sum, idiUnit, jObj);
                                //    addRowToList(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials7);
                                //    break;
                                //case 8:
                                //    //insertRow(ws6, curRow6, criteria, roomsSB.ToString(), sum, idiUnit, jObj);
                                //    addRowToList(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials8);
                                //    break;
                                //case 9:
                                //    //insertRow(ws6, curRow6, criteria, roomsSB.ToString(), sum, idiUnit, jObj);
                                //    addRowToList(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials9);
                                //    break;
                            }
                        }
                        //数据库中没有查询到的编号
                        else
                        {
                            switch (wsNum)
                            {
                                case 2:
                                    addRowToListNoMatch(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials2);
                                    break;
                                case 3:
                                    addRowToListNoMatch(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials3);
                                    break;
                                case 4:
                                    addRowToListNoMatch(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials4);
                                    break;
                                case 5:
                                    addRowToListNoMatch(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials5);
                                    break;
                                case 6:
                                    addRowToListNoMatch(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials6);
                                    break;
                                case 7:
                                    addRowToListNoMatch(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials7);
                                    break;
                                case 8:
                                    addRowToListNoMatch(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials8);
                                    break;
                                case 9:
                                    addRowToListNoMatch(criteria, roomsSB.ToString(), sum, idiUnit, jObj, ref materials9);
                                    break;
                            }
                        }

                    }
                    catch (WebException e)
                    {
                        if (e.Status == WebExceptionStatus.Timeout)
                        {
                            MessageBox.Show("处理： " + criteria + "时出错！");
                            MessageBox.Show("链接数据库超时！");
                        }
                        else
                        {
                            MessageBox.Show("处理： " + criteria + "时出错！");
                            MessageBox.Show(e.Message + "\n" + e.StackTrace);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("处理： " + criteria + "时出错！");
                        throw;
                    }

                    ProgressBarController.setValue(progressBar, curRow2 + curRow3 + curRow4 + curRow5 + curRow6 - 19);
                    switch (wsNum)
                    {
                        case 2:
                            curRow2++;
                            break;
                        case 3:
                            curRow3++;
                            break;
                        case 4:
                            curRow4++;
                            break;
                        case 5:
                            curRow5++;
                            break;
                        case 6:
                            curRow6++;
                            break;
                        //case 7:
                        //    curRow7++;
                        //    break;
                        //case 8:
                        //    curRow8++;
                        //    break;
                        //case 9:
                        //    curRow9++;
                        //    break;
                    }
                }

                //开始写入excel
                progressBar.Value = 0;
                int process = 0;
                writeListToExcel(materials2, fucaiOrder, ws2, ref process, true);
                writeListToExcel(materials3, zhucaiOrder, ws3, ref process, true);
                writeListToExcel(materials4, ruanzhuangOrder, ws4, ref process, true);
                writeListToExcel(materials5, zigouOrder, ws5, ref process, true);
                writeListToExcel(materials6, daidingOrder, ws6, ref process, true);
                //writeListToExcel(materials7, tatamiOrder, ws7, ref process, true);
                //writeListToExcel(materials8, yimaoguiOrder, ws8, ref process, true);
                //writeListToExcel(materials9, new List<string>(), ws9, ref process, false);

                #endregion

                #region 构家物料及物流方案
                ws1.Cells.Font.Name = "汉仪旗黑-70S";
                ws1.Cells.WrapText = true;
                ws1.Cells.Font.Size = 11;

                Borders borders = ws1.Range[ws1.Cells[1, 1], ws1.Cells[38, 6]].Borders;
                borders.LineStyle = XlLineStyle.xlContinuous;
                borders.Weight = XlBorderWeight.xlThin;

                ws1.Range[ws1.Cells[1, 1], ws1.Cells[1, 6]].Merge();
                Range cell = ws1.Cells[1, 1];
                cell.Value = "构家整体家装项目信息";
                cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                cell.Font.Bold = true;
                cell.RowHeight = 31.5;
                cell.Font.Size = 24;

                ws1.Range[ws1.Cells[12, 5], ws1.Cells[37, 6]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws1.Range[ws1.Cells[12, 5], ws1.Cells[37, 6]].NumberFormat = "0.00";
                //System.Reflection.Assembly CurrAssembly = System.Reflection.Assembly.LoadFrom(System.Windows.Forms.Application.ExecutablePath);
                //System.IO.Stream stream = CurrAssembly.GetManifestResourceStream("GouJiaApp.Resources.logo_sm.png");
                //string temp = System.IO.Path.GetTempFileName();
                //System.Drawing.Image.FromStream(stream).Save(temp);
                //ws1.Shapes.AddPicture(temp, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 30, 1, -1, -1);
                ws1.Range[ws1.Cells[2, 1], ws1.Cells[2, 6]].Merge();
                cell = ws1.Cells[2, 1];
                cell.Value = "4S店项目销售订单编号：";
                cell.Font.Bold = true;
                cell.RowHeight = 20.25;
                cell.Font.Size = 16;


                ws1.Range[ws1.Cells[3, 1], ws1.Cells[3, 6]].Merge();
                cell = ws1.Cells[3, 1];
                cell.Value = "4S店信息：";
                cell.Font.Bold = true;
                cell.RowHeight = 20.25;
                cell.Font.Size = 16;


                ws1.Range[ws1.Cells[4, 1], ws1.Cells[4, 3]].Merge();
                cell = ws1.Cells[4, 1];
                cell.Value = "产品包名称：";
                cell.Font.Bold = true;
                cell.RowHeight = 20.25;
                cell.Font.Size = 16;
                ws1.Range[ws1.Cells[4, 4], ws1.Cells[4, 6]].Merge();
                cell = ws1.Cells[4, 4];
                cell.Value = "项目名称：";
                cell.Font.Bold = true;
                cell.Font.Size = 16;

                ws1.Range[ws1.Cells[5, 1], ws1.Cells[5, 3]].Merge();
                cell = ws1.Cells[5, 1];
                cell.Value = "开工时间：";
                cell.Font.Bold = true;
                cell.RowHeight = 20.25;
                cell.Font.Size = 16;
                ws1.Range[ws1.Cells[5, 4], ws1.Cells[5, 6]].Merge();
                cell = ws1.Cells[5, 4];
                cell.Value = "竣工时间：";
                cell.Font.Bold = true;
                cell.Font.Size = 16;

                ws1.Range[ws1.Cells[6, 1], ws1.Cells[6, 6]].Merge();
                cell = ws1.Cells[6, 1];
                cell.Value = "收货地址：";
                cell.Font.Bold = true;
                cell.RowHeight = 20.25;
                cell.Font.Size = 16;

                ws1.Range[ws1.Cells[7, 1], ws1.Cells[7, 3]].Merge();
                cell = ws1.Cells[7, 1];
                cell.Value = "收货人：";
                cell.Font.Bold = true;
                cell.RowHeight = 20.25;
                cell.Font.Size = 16;
                ws1.Range[ws1.Cells[7, 4], ws1.Cells[7, 6]].Merge();
                cell = ws1.Cells[7, 4];
                cell.Value = "收货人联系电话：";
                cell.Font.Bold = true;
                cell.Font.Size = 16;

                ws1.Range[ws1.Cells[8, 1], ws1.Cells[8, 3]].Merge();
                cell = ws1.Cells[8, 1];
                cell.Value = "物料对接人：";
                cell.Font.Bold = true;
                cell.RowHeight = 20.25;
                cell.Font.Size = 16;
                ws1.Range[ws1.Cells[8, 4], ws1.Cells[8, 6]].Merge();
                cell = ws1.Cells[8, 4];
                cell.Value = "物料对接人联系电话：";
                cell.Font.Bold = true;
                cell.Font.Size = 16;

                ws1.Range[ws1.Cells[9, 1], ws1.Cells[9, 3]].Merge();
                cell = ws1.Cells[9, 1];
                cell.Value = "项目经理：";
                cell.Font.Bold = true;
                cell.RowHeight = 20.25;
                cell.Font.Size = 16;
                ws1.Range[ws1.Cells[9, 4], ws1.Cells[9, 6]].Merge();
                cell = ws1.Cells[9, 4];
                cell.Value = "项目经理联系电话：";
                cell.Font.Bold = true;
                cell.Font.Size = 16;

                ws1.Range[ws1.Cells[10, 1], ws1.Cells[10, 3]].Merge();
                cell = ws1.Cells[10, 1];
                cell.Value = "业主：";
                cell.Font.Bold = true;
                cell.RowHeight = 20.25;
                cell.Font.Size = 16;
                ws1.Range[ws1.Cells[10, 4], ws1.Cells[10, 6]].Merge();
                cell = ws1.Cells[10, 4];
                cell.Value = "业主联系电话：";
                cell.Font.Bold = true;
                cell.Font.Size = 16;

                cell = ws1.Rows[11];
                cell.Font.Bold = true;
                cell.Font.Size = 14;
                cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                cell = ws1.Cells[11, 1];
                cell.Value = "序号";
                cell.ColumnWidth = 6.13;

                ws1.Cells[11, 2].ColumnWidth = 12.5;
                ws1.Cells[11, 3].ColumnWidth = 27.25;

                ws1.Range[ws1.Cells[11, 2], ws1.Cells[11, 3]].Merge();
                cell = ws1.Cells[11, 2];
                cell.Value = "物料说明";

                cell = ws1.Cells[11, 4];
                cell.Value = "详细参见";
                cell.ColumnWidth = 22.54;

                cell = ws1.Cells[11, 5];
                cell.Value = "总金额";
                cell.ColumnWidth = 15.54;

                cell = ws1.Cells[11, 6];
                cell.Value = "物流费";
                cell.ColumnWidth = 11.21;

                //辅材包
                ws1.Range[ws1.Cells[12, 1], ws1.Cells[15, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws1.Range[ws1.Cells[12, 3], ws1.Cells[15, 3]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                ws1.Range[ws1.Cells[12, 1], ws1.Cells[15, 1]].Merge();
                ws1.Cells[12, 1] = 1;

                ws1.Range[ws1.Cells[12, 2], ws1.Cells[15, 2]].Merge();
                ws1.Cells[12, 2] = "辅材包";

                ws1.Cells[12, 3] = "地漏";
                ws1.Cells[13, 3] = "中央空调、地暖、空气能热水器";
                ws1.Cells[14, 3] = "乳胶漆";
                ws1.Cells[15, 3] = "瓷砖";

                ws1.Range[ws1.Cells[12, 4], ws1.Cells[15, 4]].Merge();
                ws1.Cells[12, 4] = "辅材包";

                ws1.Range[ws1.Cells[12, 5], ws1.Cells[15, 5]].Merge();
                ws1.Cells[12, 5] = string.Format("=SUM(辅材包!S{0}:辅材包!S{1})",4,ws2.UsedRange.Rows.Count);

                //ws1.Range[ws1.Cells[7, 5], ws1.Cells[10, 5]].Merge();
                //ws1.Range[ws1.Cells[7, 6], ws1.Cells[10, 6]].Merge();
                //ws1.Range[ws1.Cells[7, 7], ws1.Cells[10, 7]].Merge();
                //ws1.Range[ws1.Cells[7, 8], ws1.Cells[10, 8]].Merge();
                //ws1.Range[ws1.Cells[7, 9], ws1.Cells[10, 9]].Merge();
                //ws1.Range[ws1.Cells[7, 10], ws1.Cells[10, 10]].Merge();

                //ws1.Range[ws1.Cells[7, 11], ws1.Cells[28, 11]].Merge();
                //cell = ws1.Cells[7, 11];
                //cell.VerticalAlignment = XlVAlign.xlVAlignCenter;
                //cell.Value = "50%订金进入供应链账户后第二个工作日";

                //ws1.Range[ws1.Cells[7, 12], ws1.Cells[10, 12]].Merge();
                //cell = ws1.Cells[7, 12];
                //cell.Value = "订金进入供应链账户后12天";

                //主材包
                ws1.Range[ws1.Cells[16, 1], ws1.Cells[28, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws1.Range[ws1.Cells[16, 3], ws1.Cells[28, 3]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                ws1.Range[ws1.Cells[16, 1], ws1.Cells[28, 1]].Merge();
                ws1.Cells[16, 1] = 2;

                ws1.Range[ws1.Cells[16, 2], ws1.Cells[28, 2]].Merge();
                ws1.Cells[16, 2] = "主材包";

                ws1.Cells[16, 3] = "卫浴五金";
                ws1.Cells[17, 3] = "坐便器及配件";
                ws1.Cells[18, 3] = "厨房水槽及龙头";
                ws1.Cells[19, 3] = "门五金";
                ws1.Cells[20, 3] = "厨房电器";
                ws1.Cells[21, 3] = "集成吊顶、LED灯、浴霸、凉霸";
                ws1.Cells[22, 3] = "照明灯具、开关面板";
                ws1.Cells[23, 3] = "木门、门套、踢脚线、背景墙";
                ws1.Cells[24, 3] = "柜体收纳";
                ws1.Cells[25, 3] = "壁纸、可调色乳胶漆、背景墙硬包";
                ws1.Cells[26, 3] = "地板";
                ws1.Cells[27, 3] = "浴室柜、阳台洗衣柜";
                ws1.Cells[28, 3] = "整体橱柜";

                ws1.Range[ws1.Cells[16, 4], ws1.Cells[28, 4]].Merge();
                ws1.Cells[16, 4] = "主材包";

                ws1.Range[ws1.Cells[16, 5], ws1.Cells[28, 5]].Merge();
                ws1.Cells[16, 5] = string.Format("=SUM(主材包!S{0}:主材包!S{1})", 4, ws3.UsedRange.Rows.Count);
                //ws1.Range[ws1.Cells[11, 5], ws1.Cells[23, 5]].Merge();
                //ws1.Range[ws1.Cells[11, 6], ws1.Cells[23, 6]].Merge();
                //ws1.Range[ws1.Cells[11, 7], ws1.Cells[23, 7]].Merge();
                //ws1.Range[ws1.Cells[11, 8], ws1.Cells[23, 8]].Merge();
                //ws1.Range[ws1.Cells[11, 9], ws1.Cells[23, 9]].Merge();
                //ws1.Range[ws1.Cells[11, 10], ws1.Cells[23, 10]].Merge();

                //ws1.Range[ws1.Cells[11, 12], ws1.Cells[23, 12]].Merge();
                //cell = ws1.Cells[11, 12];
                //cell.Value = "主材包除木作外在订金进入供应链后44天内到货\n木作需二次签字确认后44天内到货";


                //软装包
                ws1.Range[ws1.Cells[29, 1], ws1.Cells[33, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws1.Range[ws1.Cells[29, 3], ws1.Cells[33, 3]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                ws1.Range[ws1.Cells[29, 1], ws1.Cells[33, 1]].Merge();
                ws1.Cells[29, 1] = 3;

                ws1.Range[ws1.Cells[29, 2], ws1.Cells[33, 2]].Merge();
                ws1.Cells[29, 2] = "软装包";

                ws1.Cells[29, 3] = "成品窗帘";
                ws1.Cells[30, 3] = "艺术灯具";
                ws1.Cells[31, 3] = "家用电器";
                ws1.Cells[32, 3] = "地毯、挂画、床垫";
                ws1.Cells[33, 3] = "移动家具";

                ws1.Range[ws1.Cells[29, 4], ws1.Cells[33, 4]].Merge();
                ws1.Cells[29, 4] = "软装包";

                ws1.Range[ws1.Cells[29, 5], ws1.Cells[33, 5]].Merge();
                ws1.Cells[29, 5]= string.Format("=SUM(软装包!S{0}:软装包!S{1})", 4, ws4.UsedRange.Rows.Count);
                //ws1.Range[ws1.Cells[24, 5], ws1.Cells[28, 5]].Merge();
                //ws1.Range[ws1.Cells[24, 6], ws1.Cells[28, 6]].Merge();
                //ws1.Range[ws1.Cells[24, 7], ws1.Cells[28, 7]].Merge();
                //ws1.Range[ws1.Cells[24, 8], ws1.Cells[28, 8]].Merge();
                //ws1.Range[ws1.Cells[24, 9], ws1.Cells[28, 9]].Merge();
                //ws1.Range[ws1.Cells[24, 10], ws1.Cells[28, 10]].Merge();
                //ws1.Range[ws1.Cells[24, 12], ws1.Cells[28, 12]].Merge();
                //cell = ws1.Cells[24, 12];
                //cell.Value = "订金进入供应链账户后50天内到货";

                //城运商自购
                ws1.Range[ws1.Cells[34, 1], ws1.Cells[37, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws1.Range[ws1.Cells[34, 3], ws1.Cells[37, 3]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                ws1.Range[ws1.Cells[34, 1], ws1.Cells[37, 1]].Merge();
                ws1.Cells[34, 1] = 4;

                ws1.Range[ws1.Cells[34, 2], ws1.Cells[37, 2]].Merge();
                ws1.Cells[34, 2] = "城运商自购";

                ws1.Cells[34, 3] = "石材";
                ws1.Cells[35, 3] = "玻璃制品";
                ws1.Cells[36, 3] = "零星材料";
                ws1.Cells[37, 3] = "特殊定制品";

                ws1.Range[ws1.Cells[34, 4], ws1.Cells[37, 4]].Merge();
                ws1.Cells[34, 4] = "城运商自购";

                ws1.Range[ws1.Cells[34, 5], ws1.Cells[37, 5]].Merge();
                ws1.Cells[34, 5] = string.Format("=SUM(城运商自购!S{0}:城运商自购!S{1})", 4, ws5.UsedRange.Rows.Count);

                //物流费
                ws1.Range[ws1.Cells[12, 6], ws1.Cells[33, 6]].Merge();
                ws1.Range[ws1.Cells[34, 6], ws1.Cells[37, 6]].Merge();

                //合计
                ws1.Rows[38].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws1.Rows[38].RowHeight = 28;
                ws1.Range[ws1.Cells[38, 2], ws1.Cells[38, 4]].Merge();
                ws1.Cells[38, 2] = "合计";
                ws1.Range[ws1.Cells[38, 5], ws1.Cells[38, 6]].Merge();
                ws1.Cells[38, 5].NumberFormat = "0.00";
                ws1.Cells[38, 5] = "=E12+E16+E29+F12";
                ////软装包
                //ws1.Range[ws1.Cells[29, 1], ws1.Cells[29, 13]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                //ws1.Cells[29, 1] = 4;
                //ws1.Cells[29, 2] = "阳台包";
                //ws1.Cells[29, 4] = "阳台包";

                ////合计
                //ws1.Range[ws1.Cells[30, 1], ws1.Cells[30, 13]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                //ws1.Range[ws1.Cells[30, 1], ws1.Cells[30, 3]].Merge();
                //ws1.Cells[30, 1] = "合计";

                //ws1.Cells[30, 7] = "=SUM(G7:G29)";
                //ws1.Cells[30, 8] = "=SUM(H7:H29)";
                //ws1.Cells[30, 9] = "=SUM(I7:I29)";
                //ws1.Cells[30, 10] = "=SUM(J7:J29)";


                ////临保材料
                //ws1.Range[ws1.Cells[31, 1], ws1.Cells[33, 12]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                //ws1.Range[ws1.Cells[31, 3], ws1.Cells[33, 3]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                //ws1.Range[ws1.Cells[31, 1], ws1.Cells[33, 1]].Merge();
                //ws1.Cells[31, 1] = 5;

                //ws1.Range[ws1.Cells[31, 2], ws1.Cells[33, 2]].Merge();
                //ws1.Cells[31, 2] = "临保材料";

                //ws1.Cells[31, 3] = "成品保护";
                //ws1.Cells[32, 3] = "形象用品";
                //ws1.Cells[33, 3] = "周转材料";

                //ws1.Range[ws1.Cells[31, 4], ws1.Cells[33, 12]].Merge();
                //ws1.Cells[31, 4] = "临设保护包作为标品（城运商提前备仓）不作为产品包物料，详见【构家】 物料清单（临保材料）";

                ////4s店自购材料
                //ws1.Range[ws1.Cells[34, 1], ws1.Cells[37, 12]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                //ws1.Range[ws1.Cells[34, 3], ws1.Cells[37, 3]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                //ws1.Range[ws1.Cells[34, 1], ws1.Cells[37, 1]].Merge();
                //ws1.Cells[34, 1] = 6;

                //ws1.Range[ws1.Cells[34, 2], ws1.Cells[37, 2]].Merge();
                //ws1.Cells[34, 2] = "4S店自购物料";

                //ws1.Cells[34, 3] = "石材";
                //ws1.Cells[35, 3] = "玻璃制品";
                //ws1.Cells[36, 3] = "零星材料";
                //ws1.Cells[37, 3] = "特殊定制品";

                //ws1.Range[ws1.Cells[34, 4], ws1.Cells[37, 12]].Merge();
                //ws1.Cells[34, 4] = "详见【构家】 物料清单（自购物料）";

                ////说明
                //ws1.Range[ws1.Cells[38, 1], ws1.Cells[38, 12]].Merge();
                //ws1.Cells[38, 1].RowHeight = 72;

                //ws1.Cells[38, 1] = "说明：\n　　1、辅材包下定时，以整包的一半货款下订金；\n　　2、供货周期以城运商支付订金后第二天起算，若延期打款，则发货时间、到货时间依据打款时间顺延；\n　　3、软装包发货前付清尾款；";

                ////代入合计
                //ws1.Cells[7, 7] = string.Format("='辅材包'!L{0}", materials2.Count + 3);
                //ws1.Cells[11, 7] = string.Format("='主材包'!L{0}", materials3.Count + 3);
                //ws1.Cells[24, 7] = string.Format("='软装包'!L{0}", materials4.Count + 3);

                #endregion

                
                ProgressBarController.setValue(progressBar, 0);
            }
            catch (Exception ex)
            {
                MessageBox.Show("处理： " + errorId + "时出错！");
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
            finally
            {
                if (xlApp != null)
                {
                    //保存关闭
                    xlApp.DisplayAlerts = false;
                    //wb.SaveAs(savePath, XlFileFormat.xlWorkbookNormal);
                    wb.SaveAs(savePath);
                    idiWB.Close();
                    wb.Close();
                    xlApp.DisplayAlerts = true;
                    xlApp.Quit();
                }
            }
        }

        private int getPackageWorkSheetNumByUseNature(string text)
        {
            switch (text)
            {
                case "辅材包":
                    return 2;
                case "主材包":
                    return 3;
                case "软装包":
                    return 4;
                case "其它":
                    return 5;
                default:
                    return 9;
            }
        }

        #region 辅助函数
        private void w7Total(Worksheet ws, int row)
        {
            ws.Cells[row, 12] = string.Format("=J{0}*G{0}", row);
            ws.Cells[row, 13] = string.Format("=K{0}*G{0}", row);
        }

        private void writeListToExcel(List<Material> materials, List<string> orderList, Worksheet ws, ref int process, bool toSum)
        {

            int curRow = 4;

            //开始按顺序规则写入excel，记录已写入excel的材料的id
            List<string> materialWritten = new List<string>();
            foreach (string order in orderList)
            {
                foreach (Material material in materials)
                {
                    if (material.name != null && material.name.Contains(order) && !materialWritten.Contains(material.id))
                    {
                        writeRow(ws, curRow, material);
                        materialWritten.Add(material.id);
                        curRow++;
                        process++;
                        ProgressBarController.setValue(progressBar, process);
                    }
                }
            }

            //写入不在顺序列表的
            foreach (Material material in materials)
            {
                if (!materialWritten.Contains(material.id))
                {
                    writeRow(ws, curRow, material);
                    materialWritten.Add(material.id);
                    curRow++;
                    process++;
                    ProgressBarController.setValue(progressBar, process);
                }
            }

            //if (toSum)
            //{
            //    //合计供货总价
            //    ws.Range[ws.Cells[curRow, 2], ws.Cells[curRow, 11]].Merge();
            //    Range cell = ws.Cells[curRow, 2];
            //    cell.Value = "合计";
            //    cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            //    cell.RowHeight = 36;

            //    if (curRow == 4)//worksheet是空的
            //        ws.Cells[curRow, 19] = 0;
            //    else
            //        ws.Cells[curRow, 19] = string.Format("=SUM(S3:S{0})", curRow - 1);

            //    Range cellRow = ws.Range[ws.Cells[curRow, 1], ws.Cells[curRow, 20]];
            //    Borders borders = cellRow.Borders;
            //    borders.LineStyle = XlLineStyle.xlDot;
            //    borders.Weight = XlBorderWeight.xlHairline;

            //    borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            //    borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            //    borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            //    borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
            //    borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            //    borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
            //}
        }

        private void writeRow(Worksheet ws, int curRow, Material material)
        {
            ws.Rows[curRow].RowHeight = 110;

            ws.Cells[1][curRow] = curRow - 3;
            ws.Cells[2][curRow] = material.id;
            ws.Cells[3][curRow] = material.rooms;
            ws.Cells[4][curRow] = material.idiQuantity;
            ws.Cells[5][curRow] = material.idiUnit;

            //物料名称
            ws.Cells[6][curRow] = material.name;
            if (material.name != null)
            {
                foreach (string key in hilightKeys)
                {
                    if (material.name.Contains(key))
                    {
                        Range celll = ws.Range[ws.Cells[curRow, 1], ws.Cells[curRow, 20]];
                        celll.Interior.Color = System.Drawing.Color.FromArgb(240, 240, 240);
                        break;
                    }
                }
            }


            //单位
            ws.Cells[9][curRow] = material.unit;
            //if ("桶".Equals(material.unit))
            //    ws.Cells[8][curRow].Interior.Color = System.Drawing.Color.FromArgb(255, 255, 0);

            //供货价
            ws.Cells[17][curRow] = material.goujiaPrice;

            //销售价
            ws.Cells[18][curRow] = material.salePrice;

            //品牌
            ws.Cells[11][curRow] = material.brand;

            //型号
            ws.Cells[12][curRow] = material.model;

            //规格
            ws.Cells[13][curRow] = material.dimension;

            //数量

            string tempRule = getQuantityRules(material.name, material.unit, material.dimension, material.id, material.brand);
            string rule = null;
            if (tempRule.Contains("Q15D"))
                rule = tempRule.Replace("Q15D", sum0015D.ToString());
            else if (tempRule.Contains("Q23D"))
                rule = tempRule.Replace("Q23D", sum0023D.ToString());
            else if (tempRule.Contains("Q19D"))
                rule = tempRule.Replace("Q19D", sum0019D.ToString());

            else if (tempRule.Contains("Q26D"))
                rule = tempRule.Replace("Q26D", sum0026D.ToString());
            else if (tempRule.Contains("Q28D"))
                rule = tempRule.Replace("Q28D", sum0028D.ToString());
            else if (tempRule.Contains("Q30D"))
                rule = tempRule.Replace("Q30D", sum0030D.ToString());
            else if (tempRule.Contains("Q32D"))
                rule = tempRule.Replace("Q32D", sum0032D.ToString());
            else if (tempRule.Contains("Q34D"))
                rule = tempRule.Replace("Q34D", sum0034D.ToString());
            else if (tempRule.Contains("Q36D"))
                rule = tempRule.Replace("Q36D", sum0036D.ToString());
            else if (tempRule.Contains("Q39D"))
                rule = tempRule.Replace("Q39D", sum0039D.ToString());
            else if (tempRule.Contains("Q41D"))
                rule = tempRule.Replace("Q41D", sum0041D.ToString());
            else if (tempRule.Contains("Q43D"))
                rule = tempRule.Replace("Q43D", sum0043D.ToString());
            else if (tempRule.Contains("Q45D"))
                rule = tempRule.Replace("Q45D", sum0045D.ToString());
            else if (tempRule.Contains("Q47D"))
                rule = tempRule.Replace("Q47D", sum0047D.ToString());
            else if (tempRule.Contains("Q49D"))
                rule = tempRule.Replace("Q49D", sum0049D.ToString());
            else if (tempRule.Contains("Q53D"))
                rule = tempRule.Replace("Q53D", sum0053D.ToString());
            else if (tempRule.Contains("Q58D"))
                rule = tempRule.Replace("Q58D", sum0058D.ToString());
            else
                rule = tempRule.Replace("Q", "D" + curRow);

            if ("嘉宝莉".Equals(material.brand))
            {
                if ("WMF03-035-0015".Equals(material.id))
                    rule = string.Format("ROUNDUP({0}*2.55/18,0)", "D" + curRow);
                else 
                    rule = getVarnishRule("嘉宝莉", material.dimension, material.checker, material.idiQuantity);
            }

            if ("片".Equals(material.unit) || "卷".Equals(material.unit))
                ws.Cells[7][curRow] = string.Format("=ROUNDUP({0},0)", rule);
            else
                ws.Cells[7][curRow] = string.Format("={0}", rule);

            //供货总价
            ws.Cells[19][curRow] = string.Format("=G{0}*Q{0}", curRow);

            //销售总价
            ws.Cells[20][curRow] = string.Format("=G{0}*R{0}", curRow);

            //技术参数
            ws.Cells[14][curRow] = material.color;

            //备注
            ws.Cells[15][curRow] = material.remark;

            //ws.Rows[curRow].WrapText = true;

            //参考图片
            Range cell = ws.Cells[10][curRow];
            string imgLocalPath = material.image;
            if (!String.IsNullOrEmpty(imgLocalPath))
            {
                //if (cell.RowHeight < 110)
                //    cell.RowHeight = 110;

                Image image = Image.FromFile(imgLocalPath);
                //ws.Shapes.AddPicture(imgLocalPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 12, cell.Top + 5 + curRow * 0.45, 100, 100);
                ws.Shapes.AddPicture(imgLocalPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 12, cell.Top + 5, 100, 100);

                cell.Value = "";
            }

            cell = ws.Range[ws.Cells[curRow, 1], ws.Cells[curRow, 3]];
            if (material.status == -1)
            {
                cell.Interior.Color = System.Drawing.Color.FromArgb(146, 208, 80);
            }
            else if (material.status != 1)
            {
                cell.Interior.Color = System.Drawing.Color.FromArgb(255, 0, 0);
            }

            Range cellRow = ws.Range[ws.Cells[curRow, 1], ws.Cells[curRow, 20]];
            Borders borders = cellRow.Borders;
            borders.LineStyle = XlLineStyle.xlDot;
            borders.Weight = XlBorderWeight.xlHairline;

            borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
            borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
        }

        private string getVarnishRule(string brand, string dimension, string checker, double idiQuantity)
        {
            string formular1 = "IF(XXX*0.3-20*ROUNDDOWN(XXX*0.3/20,0)<12.8,ROUNDDOWN(XXX*0.3/20,0),ROUNDDOWN(XXX*0.3/20,0)+1)";
            string formular2 = "IF(XXX*0.4-20*ROUNDDOWN(XXX*0.4/20,0)<11,ROUNDDOWN(XXX*0.4/20,0),ROUNDDOWN(XXX*0.4/20,0)+1)";
            string formular3 = "IF(YYY*0.4-20*ROUNDDOWN(YYY*0.4/20,0)=0,0,IF(YYY*0.4-20*ROUNDDOWN(YYY*0.4/20,0)<=6.4,1,IF(YYY*0.4-20*ROUNDDOWN(YYY*0.4/20,0)<=12.8,2,0)))";
            string formular4 = "ROUNDUP(XXX*0.4/6.4,0)";
            string formular5 = "IF(YYY*0.4-20*ROUNDDOWN(YYY*0.4/20,0)=0,0,IF(YYY*0.4-20*ROUNDDOWN(YYY*0.4/20,0)<=5.5,1,IF(YYY*0.4-20*ROUNDDOWN(YYY*0.4/20,0)<=11,2,0)))";
            string formular6 = "ROUNDUP(XXX*0.4/5.5,0)";

            switch (brand)
            {
                case "嘉宝莉":
                    if (dimension.Contains("20KG"))
                    {
                        if(string.IsNullOrEmpty(checker))
                            return formular1.Replace("XXX", idiQuantity.ToString());
                        else
                            return formular2.Replace("XXX", idiQuantity.ToString());
                    }
                    else if (dimension.Contains("6.4KG"))
                    {
                        if (Q20kg_JBL.ContainsKey(checker))
                            return formular3.Replace("YYY", Q20kg_JBL[checker].ToString());
                        else
                            return formular4.Replace("XXX",idiQuantity.ToString());
                    }
                    else if (dimension.Contains("5.5KG"))
                    {
                        if (Q20kg_JBL.ContainsKey(checker))
                                return formular5.Replace("YYY", Q20kg_JBL[checker].ToString());
                            else
                                return formular6.Replace("XXX", idiQuantity.ToString());
                    }
                    break;
            }
            return null;
        }


        private void addRowToListNoMatch(string criteria, string rooms, double sum, string idiUnit, JObject jObj, ref List<Material> materials)
        {
            Material material = new Material(criteria, rooms, sum, idiUnit, null, null, null,
                null, null, null, null, null, null, null, null, -1,null);
            materials.Add(material);
        }

        private void addRowToList(string criteria, string rooms,
            double sum, string idiUnit, JObject jObj, ref List<Material> materials)
        {
            //物料名称
            string name = (string)jObj["result"]["name"];

            //单位
            string unit = (string)jObj["result"]["matterUnitName"];

            //图片
            string image = null;

            if (!JTokenIsNullOrEmpty(jObj["result"]["imagePath"]))
            {
                string imgPath = (string)jObj["result"]["imagePath"];

                string tempPath = Path.GetTempPath() + @"Goujia\";

                if (!Directory.Exists(tempPath))
                    Directory.CreateDirectory(tempPath);

                string imgLocalPath = tempPath + StrToMD5(imgPath) + ".png";
                image = imgLocalPath;
                if (!File.Exists(imgLocalPath))
                {
                    try
                    {
                        using (WebClient webClient = new WebClient())
                        {
                            webClient.DownloadFile(imgPath, imgLocalPath);
                            Image imageTemp = Image.FromFile(imgLocalPath);
                            Bitmap bitmap = ResizeImage(imageTemp, 100, 100);
                            imageTemp.Dispose();
                            bitmap.Save(imgLocalPath);
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("下载" + criteria + "图片时发生错误！");
                        MessageBox.Show(e.Message + "\n" + e.StackTrace);
                        image = null;
                    }
                }
            }
            //供货价
            string goujiaPrice = (string)jObj["result"]["displaySalesPrice"];
            //销售价
            string salePrice = (string)jObj["result"]["displayPric"];
            //品牌
            string brand = (string)jObj["result"]["brandName"];

            //型号
            string model = (string)jObj["result"]["model"];

            //规格
            string dimension = (string)jObj["result"]["dimension"];

            //备注
            string remark = (string)jObj["result"]["openMode"] + "\n" + (string)jObj["result"]["remark"];

            //技术参数
            string color = (string)jObj["result"]["color"] + "\n" + (string)jObj["result"]["materials"];

            //启用状态
            int status = (int)jObj["result"]["status"];


            string checker = model + "-" + name + "-" + (string)jObj["result"]["color"];

            Material material = new Material(criteria, rooms, sum, idiUnit, name, unit, image,
                goujiaPrice, salePrice, brand, model, dimension, null, remark, color, status, checker);
            materials.Add(material);

            if ("嘉宝莉".Equals(brand) && dimension.Contains("20KG"))
                addQToDict(brand, checker, sum);
        }

        private void addQToDict(string brand, string checker, double quantity)
        {
            switch (brand)
            {
                case "嘉宝莉":
                    Q20kg_JBL.Add(checker, quantity);
                    break;
            }
        }


        private int getPackageWorkSheetNum(string matName)
        {
            //if (!matName.Contains("台面") && !matName.Contains("不锈钢水槽") && !matName.Contains("摩恩") && packageCheck(chuguiKeys, matName))
            //    return 6;
            //else if (packageCheck(tatamiKeys, matName))
            //    return 7;
            //else if (packageCheck(yimaoguiKeys, matName))
            //    return 8;
            if (packageCheck(fucaiKeys, matName))
                return 2;
            else if (!matName.Contains("洗衣柜") && !matName.Contains("LED吸顶灯") && packageCheck(ruanzhuangKeys, matName)) //为了将 巴比特柜体 放入软装，而不是主材，优先筛选
                return 4;
            else if (!matName.Contains("T5灯管") && packageCheck(zhucaiKeys, matName))
                return 3;
            else if (packageCheck(zigouKeys, matName))
                return 5;
            else
                return 6;
        }

        private bool packageCheck(List<string> package, string matName)
        {
            foreach (string item in package)
            {
                if (matName.Contains(item))
                    return true;
            }
            return false;
        }

        private void initWorkSheet(Worksheet ws, string title)
        {
            ws.Cells.Font.Name = "汉仪旗黑-50S";
            ws.Cells.Font.Size = 10;
            ws.Cells.WrapText = true;

            Range cell = ws.Range[ws.Cells[1, 1], ws.Cells[1, 20]];
            cell.Font.Bold = true;
            cell.WrapText = true;
            cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            cell = ws.Range[ws.Cells[3, 1], ws.Cells[3, 20]];
            cell.Font.Bold = true;
            cell.WrapText = true;
            cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            ws.Range[ws.Cells[1, 1], ws.Cells[1, 2]].Merge();
            ws.Rows[1].RowHeight = 55;
            System.Reflection.Assembly CurrAssembly = System.Reflection.Assembly.LoadFrom(System.Windows.Forms.Application.ExecutablePath);
            System.IO.Stream stream = CurrAssembly.GetManifestResourceStream("GouJiaApp.Resources.logo_sm.png");
            string temp = System.IO.Path.GetTempFileName();
            System.Drawing.Image.FromStream(stream).Save(temp);
            ws.Shapes.AddPicture(temp, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 30, 1, -1, -1);

            ws.Rows[2].RowHeight = 54.95;
            ws.Rows[3].RowHeight = 36;

            cell = ws.Range[ws.Cells[3, 1], ws.Cells[3, 20]];
            cell.Font.Bold = true;
            Borders borders = cell.Borders;
            borders.LineStyle = XlLineStyle.xlDot;
            borders.Weight = XlBorderWeight.xlHairline;
            borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
            borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
            borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
            borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
            cell.Interior.Color = System.Drawing.Color.FromArgb(192, 192, 192);
            cell.Font.Size = 12;

            ws.Range[ws.Cells[1, 6], ws.Cells[1, 20]].Merge();
            cell = ws.Cells[1, 6];
            cell.Value = title;
            cell.Font.Name = "汉仪旗黑-70S";
            cell.Font.Size = 20;
            cell.Font.Bold = true;

            ws.Range[ws.Cells[2, 1], ws.Cells[2, 20]].Merge();
            cell = ws.Cells[2, 1];
            cell.Value = "城运商填写说明：1、只能对“实际下单数量”“二次确认信息”两列单元格进行调整；\n　　　　　　　　 2、调整的信息以红色底色填充；";
            cell.Characters[9, 61].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            cell.Characters[1, 8].Font.Bold = true;
            cell.Font.Size = 12;
            cell = ws.Range[ws.Cells[2, 1], ws.Cells[2, 20]];
            borders = cell.Borders;
            borders.LineStyle = XlLineStyle.xlContinuous;
            borders.Weight = XlBorderWeight.xlThin;
            ws.Columns[1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            cell.HorizontalAlignment = XlHAlign.xlHAlignLeft;

            cell = ws.Cells[3, 1];
            cell.Value = "序号";
            cell.ColumnWidth = 5;

            cell = ws.Cells[3, 2];
            cell.Value = "GES编号";
            cell.ColumnWidth = 15;

            cell = ws.Cells[3, 3];
            cell.Value = "区域";
            cell.ColumnWidth = 8;

            cell = ws.Cells[3, 4];
            cell.Value = "IDI工程量";
            cell.ColumnWidth = 6;
            cell.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 0);
            ws.Columns[4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns[4].NumberFormat = "0.00";

            cell = ws.Cells[3, 5];
            cell.Value = "IDI单位";
            cell.ColumnWidth = 6;
            cell.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 0);
            ws.Columns[5].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            cell = ws.Cells[3, 6];
            cell.Value = "物料名称";
            cell.ColumnWidth = 8;

            cell = ws.Cells[3, 7];
            cell.Value = "数量";
            cell.ColumnWidth = 5.5;
            ws.Columns[7].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns[7].NumberFormat = "0.00";

            cell = ws.Cells[3, 8];
            cell.Value = "实际下单数量";
            cell.ColumnWidth = 6.38;
            ws.Columns[8].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            cell.Interior.Color = System.Drawing.Color.FromArgb(255, 0, 0);

            cell = ws.Cells[3, 9];
            cell.Value = "单位";
            cell.ColumnWidth = 4.5;
            ws.Columns[9].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            cell = ws.Cells[3, 10];
            cell.Value = "参考图片";
            cell.ColumnWidth = 20;

            cell = ws.Cells[3, 11];
            cell.Value = "品牌";
            cell.ColumnWidth = 8;
            ws.Columns[11].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            cell = ws.Cells[3, 12];
            cell.Value = "型号";
            cell.ColumnWidth = 11;

            cell = ws.Cells[3, 13];
            cell.Value = "规格";
            cell.ColumnWidth = 15;

            cell = ws.Cells[3, 14];
            cell.Value = "技术参数";
            cell.ColumnWidth = 16;

            cell = ws.Cells[3, 15];
            cell.Value = "备注";
            cell.ColumnWidth = 26;

            cell = ws.Cells[3, 16];
            cell.Value = "二次确认信息";
            cell.ColumnWidth = 26;

            cell = ws.Cells[3, 17];
            cell.Value = "供货价";
            cell.ColumnWidth = 8;
            ws.Columns[17].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns[17].NumberFormat = "0.00";

            cell = ws.Cells[3, 18];
            cell.Value = "市场参考价";
            cell.ColumnWidth = 8;
            ws.Columns[18].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns[18].NumberFormat = "0.00";

            cell = ws.Cells[3, 19];
            cell.Value = "供货总价";
            cell.ColumnWidth = 8;
            ws.Columns[19].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns[19].NumberFormat = "0.00";

            cell = ws.Cells[3, 20];
            cell.Value = "市场参考总价";
            cell.ColumnWidth = 8;
            ws.Columns[20].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns[20].NumberFormat = "0.00";
        }

        private int slashCount(string str)
        {
            int count = 0;
            foreach (char c in str)
                if (c == '-') count++;
            return count;
        }

        private bool IsOpened(string wbook)
        {
            bool isOpened = true;
            Microsoft.Office.Interop.Excel.Application exApp;
            exApp = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            try
            {
                exApp.Workbooks.get_Item(wbook);
            }
            catch (Exception)
            {
                isOpened = false;
            }
            return isOpened;
        }

        private bool JTokenIsNullOrEmpty(JToken token)
        {
            return (token == null) ||
                   (token.Type == JTokenType.Array && !token.HasValues) ||
                   (token.Type == JTokenType.Object && !token.HasValues) ||
                   (token.Type == JTokenType.String && token.ToString() == String.Empty) ||
                   (token.Type == JTokenType.Null);
        }

        private Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new System.Drawing.Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        private string getQuantityRules(string name, string unit, string size, string id, string brand)
        {
            if (name != null)
            {
                if (name.Contains("瓷片") || name.Contains("墙砖") || name.Contains("抛釉砖"))
                {
                    if (unit != null && unit.Contains("片"))
                    {
                        if (size.Contains("600*600") || size.Contains("600mm*600mm"))
                            return "Q/0.36";
                        else if (size.Contains("300*600") || size.Contains("300mm*600mm"))
                            return "Q/0.18";
                        else if (size.Contains("300*300") || size.Contains("300mm*300mm"))
                            return "Q/0.09";
                        else if (size.Contains("800*800") || size.Contains("800mm*800mm"))
                            return "Q/0.64";
                        else
                            return "Q";
                    }
                    else
                        return "Q";
                }

                else if (name.Contains("轨道"))
                {
                    if (id.Equals("WHM10-013-0006"))
                    {
                        return "ROUNDUP(Q/3,0)";
                    }
                    else
                        return "Q";
                }


                else if (name.Contains("踢脚线"))
                {
                    if (id.Equals("WMZ04-021-0307") || id.Equals("WMZ14-021-0115") || id.Equals("WMZ04-021-0114"))
                        return "ROUNDUP(Q/2.2+2,0)";
                    else if (id.Equals("WMZ14-033-0258") || id.Equals("WMZ14-033-0083") || id.Equals("WMZ14-033-0003"))
                        return "ROUNDUP(Q/2.44+2,0)";
                    else if (id.Equals("WMM01-065-0632") || id.Equals("WMM01-065-0552") || id.Equals("WMM01-065-0129")
                        || id.Equals("WMM01-065-0131") || id.Equals("WMM01-065-0133")
                        || id.Equals("WMM01-065-0499") || id.Equals("WMM01-065-0501")
                        || id.Equals("WMM01-065-0171"))
                        return "ROUNDUP(Q/2.4+2,0)";
                    else
                        return "Q";
                }
                else if (name.Contains("瓷砖胶"))
                {
                    if (id.Equals("WMF02-060-0004"))
                        return "ROUNDUP(Q*10/25,0)";
                    else
                        return "Q";
                }

                else if (name.Contains("瓷砖"))
                {
                    if (unit != null && unit.Contains("片"))
                    {
                        if (size.Contains("165*165") || size.Contains("165mm*165mm"))
                            return "Q/0.03";
                        else if (size.Contains("330*330"))
                            return "Q/0.10";
                        else if (size.Contains("330mm*330mm"))
                            return "Q/0.11";
                        else if (size.Contains("600*600") || size.Contains("600mm*600mm"))
                            return "Q/0.36";
                        else if (size.Contains("800*800") || size.Contains("800mm*800mm"))
                            return "Q/0.64";
                        else if (size.Contains("300*600") || size.Contains("300mm*600mm"))
                            return "Q/0.18";
                        else if (size.Contains("300*300") || size.Contains("300mm*300mm") || size.Contains("600*150mm"))
                            return "Q/0.09";
                        else if (size.Contains("800*150mm"))
                            return "Q/0.12";
                        else
                            return "Q";
                    }
                    else
                        return "Q";
                }


                else if (name.Contains("I型抹灰石膏"))
                {
                    if (id.Equals("WMF02-060-0015"))
                        return "ROUNDUP(Q*10/25,0)";
                    else
                        return "Q";
                }
                else if (name.Contains("壁纸"))
                {
                    if (unit != null && unit.Contains("卷"))
                    {
                        if (size.Contains("0.53*10"))
                            return "Q/5.3";
                        else if (size.Contains("0.68*10"))
                            return "Q/6.8";
                        else
                            return "Q";
                    }
                    else
                        return "Q";
                }

                else if (name.Contains("电线"))
                {
                    if (unit != null && unit.Contains("卷"))
                    {
                        return "ROUNDUP(Q/110,0)";
                    }
                    else
                        return "Q";
                }
                else if (name.Contains("电话线") || name.Contains("网络线"))
                {
                    if (unit != null && unit.Contains("箱"))
                    {
                        return "ROUNDUP(Q/300,0)";
                    }
                    else
                        return "Q";
                }
                else if (name.Contains("石膏板"))
                {
                    if (unit != null && unit.Contains("张"))
                    {
                        return "ROUNDUP(Q*0.35,0)";
                    }
                    else
                        return "Q";
                }
                else if (name.Contains("18厘多层板") || name.Contains("九厘多层板") || name.Contains("杉木集成板"))
                {
                    if (unit != null && unit.Contains("张"))
                    {
                        return "ROUNDUP(Q*0.33,0)";
                    }
                    else
                        return "Q";
                }
                else if (name.Contains("柔性防水浆料"))
                {
                    if (unit != null && unit.Contains("桶"))
                    {
                        return "ROUNDUP(Q*2.55/18,0)";
                    }
                    else
                        return "Q";
                }
                else if (name.Contains("腻子"))
                {
                    if (unit != null && unit.Contains("包"))
                    {
                        return "ROUNDUP(Q*3.5/20,0)";
                    }
                    else
                        return "Q";
                }
                else if (name.Contains("墙纸胶浆"))
                {
                    if (unit != null && unit.Contains("箱"))
                    {
                        return "ROUNDUP(Q*0.1/0.4/24,0)";
                    }
                    else
                        return "Q";
                }
                else if (name.Contains("墙纸胶粉"))
                {
                    if (unit != null && unit.Contains("箱"))
                    {
                        return "ROUNDUP(Q*0.1/0.125/50,0)";
                    }
                    else
                        return "Q";
                }
                else if (name.Contains("可调色"))
                {
                    if (unit != null && unit.Contains("桶"))
                    {
                        if (brand.Equals("多乐士"))
                        {
                            if (size.Contains("18L/桶"))
                                return "IF(Q*0.36-18*ROUNDDOWN(Q*0.36/18,0)<10,ROUNDDOWN(Q*0.36/18,0),ROUNDDOWN(Q*0.36/18,0)+1)";
                            else if (size.Contains("5L/桶"))
                            {
                                if (id.Equals("WMZ19-014-0025"))
                                    return RuleSub("多乐士", "Q26D");
                                else if (id.Equals("WMZ19-014-0027"))
                                    return RuleSub("多乐士", "Q28D");
                                else if (id.Equals("WMZ19-014-0029"))
                                    return RuleSub("多乐士", "Q30D");
                                else if (id.Equals("WMZ19-014-0031"))
                                    return RuleSub("多乐士", "Q32D");
                                else if (id.Equals("WMZ19-014-0033"))
                                    return RuleSub("多乐士", "Q34D");
                                else if (id.Equals("WMZ19-014-0035"))
                                    return RuleSub("多乐士", "Q36D");
                                else if (id.Equals("WMZ19-014-0038"))
                                    return RuleSub("多乐士", "Q39D");
                                else if (id.Equals("WMZ19-014-0040"))
                                    return RuleSub("多乐士", "Q41D");
                                else if (id.Equals("WMZ19-014-0042"))
                                    return RuleSub("多乐士", "Q43D");
                                else if (id.Equals("WMZ19-014-0044"))
                                    return RuleSub("多乐士", "Q45D");
                                else if (id.Equals("WMZ19-014-0046"))
                                    return RuleSub("多乐士", "Q47D");
                                else if (id.Equals("WMZ19-014-0048"))
                                    return RuleSub("多乐士", "Q49D");
                                else if (id.Equals("WMZ19-014-0052"))
                                    return RuleSub("多乐士", "Q53D");
                                else if (id.Equals("WMZ19-014-0057"))
                                    return RuleSub("多乐士", "Q58D");
                                else if (id.Equals("WMZ19-014-0037") || id.Equals("WMZ19-014-0050") || id.Equals("WMZ19-014-0051") || id.Equals("WMZ19-014-0056"))
                                    return "ROUNDUP(Q*0.36/5,0)";
                                else
                                    return "Q";
                                //return RuleSub("多乐士", "Q23D");
                            }
                            else
                                return "Q";
                        }
                        //else if (brand.Equals("嘉宝莉"))
                        //{
                            //if (size.Contains("20KG"))
                            //{
                            //    return "IF(Q*0.4-20*ROUNDDOWN(Q*0.4/20,0)<12.8,ROUNDDOWN(Q*0.4/20,0),ROUNDDOWN(Q*0.4/20,0)+1)";
                            //}

                            //else if (size.Contains("6.4KG"))
                            //{

                            //    //if (id.Equals("WMZ19-035-0057"))
                            //    //    return RuleSub("嘉宝莉", "Q36");
                            //    //else if (id.Equals("WMZ19-035-0058"))
                            //    //    return RuleSub("嘉宝莉", "Q44");
                            //    //else if (id.Equals("WMZ19-035-0059"))
                            //    //    return RuleSub("嘉宝莉", "Q45");
                            //    //else if (id.Equals("WMZ19-035-0060"))
                            //    //    return RuleSub("嘉宝莉", "Q46");
                            //    //else if (id.Equals("WMZ19-035-0061"))
                            //    //    return RuleSub("嘉宝莉", "Q47");
                            //    //else if (id.Equals("WMZ19-035-0062"))
                            //    //    return RuleSub("嘉宝莉", "Q48");
                            //    //else if (id.Equals("WMZ19-035-0063"))
                            //    //    return RuleSub("嘉宝莉", "Q49");
                            //    //else if (id.Equals("WMZ19-035-0064"))
                            //    //    return RuleSub("嘉宝莉", "Q50");
                            //    //else if (id.Equals("WMZ19-035-0065"))
                            //    //    return RuleSub("嘉宝莉", "Q51");
                            //    //else if (id.Equals("WMZ19-035-0066"))
                            //    //    return RuleSub("嘉宝莉", "Q52");
                            //    //else if (id.Equals("WMZ19-035-0067"))
                            //    //    return RuleSub("嘉宝莉", "Q53");
                            //    //else if (id.Equals("WMZ19-035-0068"))
                            //    //    return RuleSub("嘉宝莉", "Q54");
                            //    //else if (id.Equals("WMZ19-035-0069"))
                            //    //    return RuleSub("嘉宝莉", "Q55");
                            //    //else if (id.Equals("WMZ19-035-0070"))
                            //    //    return RuleSub("嘉宝莉", "Q56");
                            //    //else if (id.Equals("WMZ19-035-0071"))
                            //    //    return RuleSub("嘉宝莉", "Q72");
                            //    //else if (id.Equals("WMZ19-035-0073"))
                            //    //    return RuleSub("嘉宝莉", "Q74");

                            //    //else if (id.Equals("WMZ19-035-0103"))
                            //    //    return RuleSub("嘉宝莉", "Q100");
                            //    //else if (id.Equals("WMZ19-035-0104"))
                            //    //    return RuleSub("嘉宝莉", "Q101");
                            //    //else if (id.Equals("WMZ19-035-0105"))
                            //    //    return RuleSub("嘉宝莉", "Q102");
                            //    //else if (id.Equals("WMZ19-035-0107"))
                            //    //    return RuleSub("嘉宝莉", "Q106");
                            //    //else if (id.Equals("WMZ19-035-0109"))
                            //    //    return RuleSub("嘉宝莉", "Q108");

                            //    //else
                            //    //    return "Q";
                            //}

                            //else
                            //    return "Q";
                        //}
                        else
                            return "Q";
                    }
                    else
                        return "Q";
                }
                else if (name.Contains("面漆"))
                {
                    if (unit != null && unit.Contains("桶"))
                    {
                        if (brand.Equals("多乐士"))
                        {
                            if (size.Contains("18L/桶"))
                                return "IF(Q*0.24-18*ROUNDDOWN(Q*0.24/18,0)<10,ROUNDDOWN(Q*0.24/18,0),ROUNDDOWN(Q*0.24/18,0)+1)";
                            else if (size.Contains("5L/桶"))
                                return "IF(Q15D*0.24-18*ROUNDDOWN(Q15D*0.24/18,0)=0,0,IF(Q15D*0.24-18*ROUNDDOWN(Q15D*0.24/18,0)<=5,1,IF(Q15D*0.24-18*ROUNDDOWN(Q15D*0.24/18,0)<=10,2,0)))";
                            else
                                return "Q";
                        }
                        else if (brand.Equals("嘉宝莉"))
                        {
                            if (size.Contains("20KG"))
                                return "IF(Q*0.3-20*ROUNDDOWN(Q*0.3/20,0)<12.8,ROUNDDOWN(Q*0.3/20,0),ROUNDDOWN(Q*0.3/20,0)+1)";
                            else if (size.Contains("6.4KG"))
                                return "IF(Q20*0.3-20*ROUNDDOWN(Q20*0.3/20,0)=0,0,IF(Q20*0.3-20*ROUNDDOWN(Q20*0.3/20,0)<=6.4,1,IF(Q20*0.3-20*ROUNDDOWN(Q20*0.3/20,0)<=12.8,2,0)))";
                            else
                                return "Q";
                        }
                        else
                            return "Q";
                    }
                    else
                        return "Q";
                }
                else if (name.Contains("底漆"))
                {
                    if (unit != null && unit.Contains("桶"))
                    {
                        if (brand.Equals("多乐士"))
                        {
                            if (size.Contains("18L/桶"))
                                return "IF(Q*0.12-18*ROUNDDOWN(Q*0.12/18,0)<10,ROUNDDOWN(Q*0.12/18,0),ROUNDDOWN(Q*0.12/18,0)+1)";
                            else if (size.Contains("5L/桶"))
                                return "IF(Q19D*0.12-18*ROUNDDOWN(Q19D*0.12/18,0)=0,0,IF(Q19D*0.12-18*ROUNDDOWN(Q19D*0.12/18,0)<=5,1,IF(Q19D*0.12-18*ROUNDDOWN(Q19D*0.12/18,0)<=10,2,0)))";
                            else
                                return "Q";
                        }
                        else if (brand.Equals("嘉宝莉"))
                        {
                            if (size.Contains("20KG"))
                                return "IF(Q*0.15-20*ROUNDDOWN(Q*0.15/20,0)<12,ROUNDDOWN(Q*0.15/20,0),ROUNDDOWN(Q*0.15/20,0)+1)";
                            else if (size.Contains("6.4KG"))
                                return "IF(Q23*0.15-20*ROUNDDOWN(Q23*0.15/20,0)=0,0,IF(Q23*0.15-20*ROUNDDOWN(Q23*0.15/20,0)<=6,1,IF(Q23*0.15-20*ROUNDDOWN(Q23*0.15/20,0)<=12,2,0)))";
                            else
                                return "Q";
                        }
                        else
                            return "Q";
                    }
                    else
                        return "Q";
                }
                else if (name.Contains("砂浆"))
                {
                    if (id.Equals("WMF02-060-0003"))
                    {
                        return "ROUNDUP(Q*20/50,0)";
                    }
                    else if (id.Equals("WMF02-060-0006"))
                    {
                        return "ROUNDUP(Q*66.67/50,0)";
                    }
                    else
                        return "Q";
                }

                else if (name.Contains("橱柜 顶线"))
                {
                    if (unit != null && unit.Contains("个"))
                    {
                        return "ROUNDUP(Q/2.4,0)";
                    }
                    else
                        return "Q";
                }
                //else if (name.Contains("抛釉砖"))
                //{
                //    if (unit != null && unit.Contains("片"))
                //    {
                //        if (size.Contains("600*600") || size.Contains("600mm*600mm"))
                //            return "Q/0.36";
                //        if (size.Contains("800*800") || size.Contains("800mm*800mm"))
                //            return "Q/0.64";
                //        else
                //            return "Q";
                //    }
                //    else
                //        return "Q";
                //}
                else if (name.Contains("黑金花"))
                {
                    if (unit != null && unit.Contains("片"))
                    {
                        if (size.Contains("800*800") || size.Contains("800mm*800mm"))
                            return "Q/0.64";
                        else
                            return "Q";
                    }
                    else
                        return "Q";
                }
                else if (name.Contains("莱茵米黄"))
                {
                    if (unit != null && unit.Contains("片"))
                    {
                        if (size.Contains("800*800") || size.Contains("800mm*800mm"))
                            return "Q/0.64";
                        else
                            return "Q";
                    }
                    else
                        return "Q";
                }
                else
                {
                    return "Q";
                }
            }
            else
                return "Q";
        }

        string RuleSub(string name, string sub)
        {
            string temp = "";
            switch (name)
            {
                case "嘉宝莉":
                    temp = "IF(XXX*0.4-20*ROUNDDOWN(XXX*0.4/20,0)=0,0,IF(XXX*0.4-20*ROUNDDOWN(XXX*0.4/20,0)<=6.4,1,IF(XXX*0.4-20*ROUNDDOWN(XXX*0.4/20,0)<=12.8,2,0)))";
                    break;
                case "多乐士":
                    temp = "IF(XXX*0.36-18*ROUNDDOWN(XXX*0.36/18,0)=0,0,IF(XXX*0.36-18*ROUNDDOWN(XXX*0.36/18,0)<=5,1,IF(XXX*0.36-18*ROUNDDOWN(XXX*0.36/18,0)<=10,2,0)))";
                    break;
            }
            return temp.Replace("XXX", sub);
        }
        #endregion

        struct Material
        {
            public string id, rooms, idiUnit, name, unit, image, goujiaPrice, salePrice;
            public double idiQuantity;
            public string brand, model, dimension, techData, remark, color;
            public int status;
            public string checker;//计算乳胶漆公式的时候用

            public Material(string id, string rooms, double idiQuantity,
                string idiUnit, string name, string unit, string image, string goujiaPrice,
                string salePrice, string brand,
                string model, string dimension, string techData, string remark, string color, int status, string checker)
            {
                this.id = id;
                this.rooms = rooms;
                this.idiQuantity = idiQuantity;
                this.idiUnit = idiUnit;
                this.name = name;
                this.unit = unit;
                this.image = image;
                this.salePrice = salePrice;
                this.goujiaPrice = goujiaPrice;
                this.brand = brand;
                this.model = model;
                this.dimension = dimension;
                this.techData = techData;
                this.remark = remark;
                this.color = color;
                this.status = status;
                this.checker = checker;
            }
        }


        public static string StrToMD5(string str)
        {
            byte[] data = Encoding.GetEncoding("GB2312").GetBytes(str);
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] OutBytes = md5.ComputeHash(data);

            string OutString = "";
            for (int i = 0; i < OutBytes.Length; i++)
            {
                OutString += OutBytes[i].ToString("x2");
            }
            // return OutString.ToUpper();
            return OutString.ToLower();
        }
    }
    
}
