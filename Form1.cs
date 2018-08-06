using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KiwoomCode;

namespace hhhh
{
    public partial class Form1 : Form
    {
        private int _scrNum = 5000;
        List<Int32> li = new List<Int32>();
        List<Int32> li2 = new List<Int32>();

        // 화면번호 생산
        private string GetScrNum()
        {
            if (_scrNum < 9999)
                _scrNum++;
            else
                _scrNum = 5000;

            return _scrNum.ToString();
        }

        // 실시간 연결 종료
        private void DisconnectAllRealData()
        {
            for (int i = _scrNum; i > 5000; i--)
            {
                axKHOpenAPI.DisconnectRealData(i.ToString());
            }

            _scrNum = 5000;
        }

        public void Logger(Log type, string format, params Object[] args)
        {
            string message = String.Format(format, args);

            switch (type)
            {
                case Log.조회:
                    lbox_prevalue.Items.Add(message);
                    lbox_prevalue.SelectedIndex = lbox_prevalue.Items.Count - 1;
                    break;
                case Log.최근:
                    lbox_20.Items.Add(message);
                    lbox_20.SelectedIndex = lbox_20.Items.Count - 1;
                    break;
                //case Log.일반:
                //    lst일반.Items.Add(message);
                //    lst일반.SelectedIndex = lst일반.Items.Count - 1;
                //    break;
                //case Log.실시간:
                //    lst일반.Items.Add(message);
                //    lst일반.SelectedIndex = lst일반.Items.Count - 1;
                //    break;
                default: break;
            }
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void tsm_login_Click(object sender, EventArgs e)
        {
            axKHOpenAPI.CommConnect();
        }

        private void tsm_logout_Click(object sender, EventArgs e)
        {
            DisconnectAllRealData();
            axKHOpenAPI.CommTerminate();
        }

        private void tsm_logstate_Click(object sender, EventArgs e)
        {
            if (axKHOpenAPI.GetConnectState() == 0)
            {
                MessageBox.Show("미 접속 중");
            }
            else
            {
                MessageBox.Show("접속 중");
            }
        }

        private void tsm_check_Click(object sender, EventArgs e)
        {
            lv_id.Text = axKHOpenAPI.GetLoginInfo("USER_NAME");
            lv_name.Text = axKHOpenAPI.GetLoginInfo("USER_ID");
            //lv_num.Text = axKHOpenAPI.GetLoginInfo("ACCNO");
            string[] account = axKHOpenAPI.GetLoginInfo("ACCNO").Split(';');
            cbox_account.Items.AddRange(account);
            cbox_account.SelectedIndex = 0;

        }

        private void tsm_end_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void tsm_present_value_Click(object sender, EventArgs e)
        {
            axKHOpenAPI.SetInputValue("종목코드", tbox_code.Text.Trim());
            int nRet = axKHOpenAPI.CommRqData("주식기본정보", "OPT10001", 0, GetScrNum());
            _scrNum++;
        }

        private void tsm_backdata_Click(object sender, EventArgs e)
        {
            axKHOpenAPI.SetInputValue("종목코드", tbox_code.Text.Trim());
            axKHOpenAPI.SetInputValue("조회일자", tbox_date.Text.Trim());
            axKHOpenAPI.SetInputValue("수정주가구분", "1");

            int nRet = axKHOpenAPI.CommRqData("주식일봉차트조회", "OPT10081", 0, GetScrNum());
            _scrNum++;
        }

        private void tsm_back20_Click(object sender, EventArgs e)
        {
            axKHOpenAPI.SetInputValue("종목코드", tbox_code.Text.Trim());
            axKHOpenAPI.SetInputValue("조회일자", tbox_date.Text.Trim());
            axKHOpenAPI.SetInputValue("표시구분", "1");

            int nRet = axKHOpenAPI.CommRqData("일별주가요청", "OPT10086", 0, GetScrNum());
            _scrNum++;
        }

        private void tsm_decision_Click(object sender, EventArgs e)
        {
            decimal result = Correlation(li, li2);
            Logger(Log.조회, "Correlation : {0}", result);
        }

        private void axKHOpenAPI_OnReceiveTrData(object sender, AxKHOpenAPILib._DKHOpenAPIEvents_OnReceiveTrDataEvent e)
        {
            string code = axKHOpenAPI.GetCommData(e.sTrCode, "", 0, "").Trim();

            //OPT10001 : 주식기본정보
            if (e.sRQName == "주식기본정보")
            {
                lbox_prevalue.Items.Clear();
                int nCnt = axKHOpenAPI.GetRepeatCnt(e.sTrCode, e.sRQName);

                Logger(Log.조회, "종목명 : {0}", axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, 0, "종목명").Trim());
                Logger(Log.조회, "현재가 : {0}", Int32.Parse(axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, 0, "현재가").Trim()));
                Logger(Log.조회, "PER : {0}", axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, 0, "PER").Trim());

            }

            //OPT10081 : 주식일봉차트조회
            if (e.sRQName == "주식일봉차트조회")
            {
                int nCnt = axKHOpenAPI.GetRepeatCnt(e.sTrCode, e.sRQName);

                for (int i = 0; i < nCnt; i++)
                {
                    Logger(Log.조회, "{0} | 현재가:{1:N0} | 시가:{2:N0} | 고가:{3:N0} | 저가:{4:N0} ",
                        axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, i, "일자").Trim(),
                        Int32.Parse(axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, i, "현재가").Trim()),
                        Int32.Parse(axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, i, "시가").Trim()),
                        Int32.Parse(axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, i, "고가").Trim()),
                        Int32.Parse(axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, i, "저가").Trim()));
                }
            }

            //OPT10086 : 일별주가요청
            if (e.sRQName == "일별주가요청")
            {
                int nCnt = axKHOpenAPI.GetRepeatCnt(e.sTrCode, e.sRQName);

                for (int i = 0; i < nCnt; i++)
                {
                    Logger(Log.최근, "{0} | 시가:{1:N0} | 고가:{2:N0} | 저가:{3:N0}| 종가:{4:N0}",
                        axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, i, "날짜").Trim(),                        
                        Int32.Parse(axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, i, "시가").Trim()),
                        Int32.Parse(axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, i, "고가").Trim()),
                        Int32.Parse(axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, i, "저가").Trim()),
                        Int32.Parse(axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, i, "종가").Trim()));                    
                    li.Add(Int32.Parse(axKHOpenAPI.CommGetData(e.sTrCode, "", e.sRQName, i, "시가").Trim()));
                    li2.Add(i);
                }
            }
        }

        private decimal Correlation(List<Int32> xList, List<Int32> yList)
        {
            decimal result;
            Int32 multiplyXYSigma = 0;
            Int32 xSigma = 0;
            Int32 ySigma = 0;
            Int32 xPowSigma = 0;
            Int32 yPowSigma = 0;
            Int32 n = xList.Count;

            for (int i = 0; i < xList.Count; i++)
            {
                multiplyXYSigma += (xList[i] * yList[i]);
                xSigma += xList[i];
                ySigma += yList[i];
                xPowSigma += (Int32)Math.Pow(xList[i], 2);
                yPowSigma += (Int32)Math.Pow(yList[i], 2);
            }
            result =
                ((n * multiplyXYSigma) - (xSigma * ySigma)) / ((Int32)Math.Sqrt(((n * xPowSigma)
                - (Int32)Math.Pow(xSigma, 2)) * ((n * yPowSigma)
                - (Int32)Math.Pow(ySigma, 2))));
            return result;
        }

        private void tsm_dreManage_Click(object sender, EventArgs e)
        {
            string manage = axKHOpenAPI.GetMasterConstruction(tbox_code.Text);

            if (manage == "투자주의")
            {
                Logger(Log.조회, "감리구분 : 투자주의");
            }
            if (manage == "정상")
            {
                Logger(Log.조회, "감리구분 : 정상");
            }
            if (manage == "투자경고")
            {
                Logger(Log.조회, "감리구분 : 투자경고");
            }
            if (manage == "투자위험")
            {
                Logger(Log.조회, "감리구분 : 투자위험");
            }
            if (manage == "투자주의환기종목")
            {
                Logger(Log.조회, "감리구분 : 투자주의환기종목");
            }
        }
    }
}