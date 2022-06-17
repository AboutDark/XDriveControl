using FluentModbus;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;

namespace MotorControlF
{
    public partial class Form1 : Form
    {
        ModbusRtuClient modbus = null;
        int id = 0;
        System.Threading.Timer timer = null;
        int delay = 0;
        int chartshow = 120;

        object modbuslock = new object();


        public Form1()
        {
            InitializeComponent();
        }

        private void test()
        {
            var client = new ModbusRtuClient()
            {
                BaudRate = 115200,
                Parity = Parity.None,
                StopBits = StopBits.Two
            };
            client.Connect("COM6", ModbusEndianness.BigEndian);

            var abc = client.ReadHoldingRegisters<UInt16>(1, 3, 2);
            MessageBox.Show(FlipDataInt32(abc.ToArray()).ToString());
            client.WriteMultipleRegisters(1, 3, FlipDataUInt16(7));

            client.Close();
        }

        private Int32 FlipDataInt32(UInt16[] data)
        {
            Int32 ret;
            ret = data[1] << 16;
            ret |= data[0];
            return ret;
        }

        private UInt16[] FlipDataUInt16(Int32 data)
        {
            UInt16[] ret = new UInt16[2];
            ret[1] = Convert.ToUInt16(data >> 16 & 0xffff);
            ret[0] = Convert.ToUInt16(data & 0xffff);
            return ret;
        }

        private String GetInt32(int addr)
        {
            lock (modbuslock)
                return FlipDataInt32(modbus.ReadHoldingRegisters<UInt16>(id, addr, 2).ToArray()).ToString();
        }

        private void SetInt32(int addr, string data)
        {
            lock (modbuslock)
                modbus.WriteMultipleRegisters<UInt16>(id, addr, FlipDataUInt16(Convert.ToInt32(data)));
        }

        private String GetInt16(int addr)
        {
            lock (modbuslock)
                return modbus.ReadHoldingRegisters<UInt16>(id, addr, 1)[0].ToString();
        }

        private void SetInt16(int addr, string data)
        {
            lock (modbuslock)
                modbus.WriteSingleRegister(id, addr, Convert.ToInt16(data));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //chart1.Series[0].Points.Add(Math.Sin(abc));
            //abc += 0.1;
            //count++;

            //if (count > size)
            //{
            //    SetChart(chart1, count - size);
            //}
            //else
            //{
            //    SetChart(chart1, 0);
            //}
        }

        private void SetChart(System.Windows.Forms.DataVisualization.Charting.Chart chart, Int32 iTimeInterval)
        {
            chart.ChartAreas["ChartArea1"].CursorX.AutoScroll = true;
            // 启动用户缩放
            chart.ChartAreas["ChartArea1"].CursorX.IsUserEnabled = true;
            chart.ChartAreas["ChartArea1"].CursorX.IsUserSelectionEnabled = true;
            chart.ChartAreas["ChartArea1"].AxisX.ScaleView.Zoomable = true;
            // 启动滚动条
            chart.ChartAreas["ChartArea1"].AxisX.ScrollBar.Enabled = true;
            // 设置轴间隔，0为自动
            chart.ChartAreas["ChartArea1"].AxisX.Interval = 0;
            // 设置显示位置
            chart.ChartAreas["ChartArea1"].AxisX.ScaleView.Position = iTimeInterval;
            chart.ChartAreas["ChartArea1"].AxisX.ScaleView.Size = 120;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox2.SelectedIndex = 0;
            comboBox4.SelectedIndex = 1;

            comboBox1.Items.AddRange(SerialPort.GetPortNames());
            if (comboBox1.Items.Count > 0)
                comboBox1.SelectedIndex = 0;

            // 设置时钟
            timer = new System.Threading.Timer(RefreshChart, null, 0, Convert.ToInt32(textBox4.Text));
            delay = Convert.ToInt32(textBox4.Text);

            // 设置图标样式
            chart1.ChartAreas["ChartArea1"].CursorX.AutoScroll = true;
            // 启动用户缩放
            chart1.ChartAreas["ChartArea1"].CursorX.IsUserEnabled = true;
            chart1.ChartAreas["ChartArea1"].CursorX.IsUserSelectionEnabled = true;
            chart1.ChartAreas["ChartArea1"].AxisX.ScaleView.Zoomable = true;
            // 启动滚动条
            chart1.ChartAreas["ChartArea1"].AxisX.ScrollBar.Enabled = true;
            // 设置轴间隔，0为自动
            chart1.ChartAreas["ChartArea1"].AxisX.Interval = 0;

            chart2.ChartAreas["ChartArea1"].CursorX.AutoScroll = true;
            // 启动用户缩放
            chart2.ChartAreas["ChartArea1"].CursorX.IsUserEnabled = true;
            chart2.ChartAreas["ChartArea1"].CursorX.IsUserSelectionEnabled = true;
            chart2.ChartAreas["ChartArea1"].AxisX.ScaleView.Zoomable = true;
            // 启动滚动条
            chart2.ChartAreas["ChartArea1"].AxisX.ScrollBar.Enabled = true;
            // 设置轴间隔，0为自动
            chart2.ChartAreas["ChartArea1"].AxisX.Interval = 0;

            // 隐藏系统设置
            // tabControl1.TabPages.RemoveAt(3);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (modbus == null)
            {
                try
                {
                    ModbusConnect();
                    button2.Text = "断开";
                    tabControl1_Selecting(null,null);

                    // 显示版本号
                    UInt32 iver = Convert.ToUInt32(GetInt32(1));
                    version.Text = $"{Convert.ToString(iver >> 24, 16)}.{Convert.ToString(iver << 8 >> 24, 16)}.{Convert.ToString(iver & 0xffff, 16)}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show( $"连接失败 {ex.Message}", "系统提示");
                }
            }
            else
            {
                button2.Text = "连接";
                modbus.Close();
                modbus = null;
            }
        }

        private void ModbusConnect()
        {
            try
            {
                modbus = new ModbusRtuClient();
                // 波特率
                modbus.BaudRate = Convert.ToInt32(textBox1.Text);
                // 奇偶校验
                modbus.Parity = (Parity)comboBox2.SelectedIndex;
                // 停止位
                modbus.StopBits = (StopBits)comboBox4.SelectedIndex;
                // 开始连接
                modbus.Connect(comboBox1.Text, ModbusEndianness.BigEndian);

                id = Convert.ToInt32(textBox5.Text);
            }
            catch (Exception ex)
            {
                modbus = null;
                throw ex;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 关闭时钟
            timer.Dispose();
            // 如果不为null，就关闭连接
            modbus?.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            // 读取目标位置
            textBox6.Text = GetInt32(84);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 写入目标位置
            SetInt32(84, textBox6.Text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            // 读取最大速度
            textBox7.Text = GetInt32(112);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // 写入最大速度
            SetInt32(112, textBox7.Text);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            // 读取DCE KP
            textBox8.Text = GetInt32(33);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            // 读取DCE KP
            SetInt32(33, textBox8.Text);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            // 读取DCE KI
            textBox9.Text = GetInt32(35);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            // 写入DCE KI
            SetInt32(35, textBox9.Text);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            // 读取DCE KV
            textBox10.Text = GetInt32(37);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            // 写入DCE KV
            SetInt32(37, textBox10.Text);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            // 读取DCE KD
            textBox11.Text = GetInt32(39);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            // 写入DCE KD
            SetInt32(39, textBox11.Text);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            // 读取目标位置
            textBox6.Text = GetInt32(84);
            // 读取最大速度
            textBox7.Text = GetInt32(112);
            // 读取DCE KP
            textBox8.Text = GetInt32(33);
            // 读取DCE KI
            textBox9.Text = GetInt32(35);
            // 读取DCE KV
            textBox10.Text = GetInt32(37);
            // 读取DCE KD
            textBox11.Text = GetInt32(39);
        }

        private void button16_Click(object sender, EventArgs e)
        {

            // 写入目标位置
            SetInt32(84, textBox6.Text);
            // 写入最大速度
            SetInt32(112, textBox7.Text);
            // 读取DCE KP
            SetInt32(33, textBox8.Text);
            // 写入DCE KI
            SetInt32(35, textBox9.Text);
            // 写入DCE KV
            SetInt32(37, textBox10.Text);
            // 写入DCE KD
            SetInt32(39, textBox11.Text);
        }

        private void button28_Click(object sender, EventArgs e)
        {
            // 读取 速度
            textBox16.Text = GetInt32(86);
        }

        private void button27_Click(object sender, EventArgs e)
        {
            // 写入 速度
            SetInt32(86, textBox16.Text);
        }

        private void button26_Click(object sender, EventArgs e)
        {
            // 读取PID KP
            textBox15.Text = GetInt32(7);
        }

        private void button25_Click(object sender, EventArgs e)
        {
            // 写入PID KP
            SetInt32(7, textBox15.Text);
        }

        private void button24_Click(object sender, EventArgs e)
        {
            // 读取PID KI
            textBox14.Text = GetInt32(9);
        }

        private void button23_Click(object sender, EventArgs e)
        {
            // 写入PID KI
            SetInt32(9, textBox14.Text);
        }

        private void button20_Click(object sender, EventArgs e)
        {
            // 读取PID KD
            textBox12.Text = GetInt32(11);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            // 写入PID KD
            SetInt32(11, textBox12.Text);
        }

        private void button18_Click(object sender, EventArgs e)
        {            // 读取 速度
            textBox16.Text = GetInt32(86);
            // 读取PID KP
            textBox15.Text = GetInt32(7);
            // 读取PID KI
            textBox14.Text = GetInt32(9);
            // 读取PID KD
            textBox12.Text = GetInt32(11);
        }

        private void button17_Click(object sender, EventArgs e)
        {
            // 写入 速度
            SetInt32(86, textBox16.Text);
            // 写入PID KP
            SetInt32(7, textBox15.Text);
            // 写入PID KI
            SetInt32(9, textBox14.Text);
            // 写入PID KD
            SetInt32(11, textBox12.Text);
        }

        private void button36_Click(object sender, EventArgs e)
        {
            // 读取 力矩
            textBox19.Text = GetInt16(88);
        }

        private void button35_Click(object sender, EventArgs e)
        {
            // 写入 力矩
            SetInt16(88, textBox19.Text);
        }

        private void button34_Click(object sender, EventArgs e)
        {
            // 读取 堵转
            textBox18.Text = GetInt32(105);
        }

        private void button33_Click(object sender, EventArgs e)
        {
            // 写入 堵转
            SetInt32(105, textBox18.Text);
        }

        private void button22_Click(object sender, EventArgs e)
        {
            // 读取 力矩
            textBox19.Text = GetInt16(88);
            // 读取 堵转
            textBox18.Text = GetInt16(105);
            // 读取使能
            textBox2.Text = GetInt16(89);
            // 读取急停
            textBox3.Text = GetInt16(90);
        }

        private void button21_Click(object sender, EventArgs e)
        {
            // 写入 力矩
            SetInt16(88, textBox19.Text);
            // 写入 堵转
            SetInt16(109, textBox18.Text);
            // 写入使能
            SetInt16(89, textBox2.Text);
            // 写入急停
            SetInt16(90, textBox3.Text);
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            switch (tabControl1.SelectedTab.Text)
            {
                case "位置模式":
                    // 将真实位置设置为目标位置
                    SetInt32(84, GetInt32(68));
                    SetInt32(58, 0x20.ToString());
                    button15_Click(null, null);
                    break;
                case "转速模式":
                    SetInt32(58, 0x21.ToString());
                    button18_Click(null, null);
                    break;
                case "力矩模式":
                    SetInt32(58, 0x22.ToString());
                    button22_Click(null, null);
                    break;
                case "系统配置":
                    button30_Click(null, null);
                    break;
                default:
                    break;
            }
        }

        private void RefreshChart(object state)
        {
            try
            {
                // 转速监控
                if (checkBox1.Checked)
                {
                    int buf = Convert.ToInt32( GetInt32(74));
                    chart1.Invoke(new Action(()=> 
                    {
                        chart1.Series[0].Points.Add(buf/51200.0);
                        // 设置显示位置
                        int count = chart1.Series[0].Points.Count;
                        count = count > chartshow ? count - chartshow : 0;

                        chart1.ChartAreas["ChartArea1"].AxisX.ScaleView.Position = count;
                        chart1.ChartAreas["ChartArea1"].AxisX.ScaleView.Size = chartshow;
                    }));
                }
            }
            catch
            {
                try
                {
                    checkBox1.Invoke( new Action(()=>checkBox1.Checked = false));
                }
                catch
                { 
                }
            }

            try
            {
                // 力矩监控
                if (checkBox2.Checked)
                {
                    int buf = Convert.ToInt32( GetInt32(101)); 
                    chart2.Invoke(new Action(() =>
                    {
                        chart2.Series[0].Points.Add(buf);
                        // 设置显示位置
                        int count = chart2.Series[0].Points.Count;
                        count = count > chartshow ? count - chartshow : 0;

                        chart2.ChartAreas["ChartArea1"].AxisX.ScaleView.Position = count;
                        chart2.ChartAreas["ChartArea1"].AxisX.ScaleView.Size = chartshow;
                    }));
                }
            }
            catch
            {
                try
                {
                    checkBox1.Invoke(new Action(() => checkBox2.Checked = false));
                }
                catch
                {
                }
            }
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            try
            {
                timer.Change(0, Convert.ToInt32(textBox4.Text));
                delay = Convert.ToInt32(textBox4.Text);
            }
            catch
            {
                textBox4.Text = delay.ToString();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                chart1.Series[0].Points.Clear();
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                chart2.Series[0].Points.Clear();
            }
        }

        private void button39_Click(object sender, EventArgs e)
        {
            // 读取使能
            textBox2.Text = GetInt16(89);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            // 写入使能
            SetInt16(89, textBox2.Text);
        }

        private void button46_Click(object sender, EventArgs e)
        {
            // 读取急停
            textBox3.Text = GetInt16(90);
        }

        private void button40_Click(object sender, EventArgs e)
        {
            // 写入急停
            SetInt16(90, textBox3.Text);
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            chartshow = (this.Size.Width - 1050)/8 + 120;
        }

        private void chart_GetToolTipText(object sender, System.Windows.Forms.DataVisualization.Charting.ToolTipEventArgs e)
        {
            // 获取图表上的Y轴
            if (e.HitTestResult.ChartElementType == ChartElementType.DataPoint)
            {
                int i = e.HitTestResult.PointIndex;
                DataPoint dp = e.HitTestResult.Series.Points[i];
                e.Text = string.Format("{1:F3}", dp.XValue, dp.YValues[0]);
            }
        }

        private void button30_Click(object sender, EventArgs e)
        {
            // 读取波特率
            textBox21.Text = GetInt32(168);
            // 读取模式
            comboBox5.SelectedIndex = Convert.ToInt16(GetInt16(163));
            // 读取站台号
            textBox13.Text = GetInt16(172);
        }

        private void button29_Click(object sender, EventArgs e)
        {
            // 写入波特率
            SetInt32(168, textBox21.Text);
            // 写入模式
            SetInt16(163, comboBox5.SelectedIndex.ToString());
            // 写入站台号
            SetInt16(172, textBox13.Text);
        }

        private void button42_Click(object sender, EventArgs e)
        {
            // 读取波特率
            textBox21.Text = GetInt32(168);
        }

        private void button41_Click(object sender, EventArgs e)
        {
            // 写入波特率
            SetInt32(168, textBox21.Text);
        }

        private void button38_Click(object sender, EventArgs e)
        {
            // 读取模式
            comboBox5.SelectedIndex = Convert.ToInt16(GetInt16(163));
        }

        private void button37_Click(object sender, EventArgs e)
        {
            // 写入模式
            SetInt16(163, comboBox5.SelectedIndex.ToString());
        }

        private void button44_Click(object sender, EventArgs e)
        {
            // 读取站台号
            textBox13.Text = GetInt16(172);
        }

        private void button43_Click(object sender, EventArgs e)
        {
            // 写入站台号
            SetInt16(172, textBox13.Text);
        }

        private void button45_Click(object sender, EventArgs e)
        {
            // 保存
            SetInt16(3, "1");
        }

        private void button47_Click(object sender, EventArgs e)
        {
            // 重启设备
            SetInt16(2, "1");
        }
    }
}
