using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Security.Cryptography;
using System.IO;

namespace MurbongMacro
{
    [ComVisible(true)]
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]


    public partial class Form1 : Form
    {


        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(ref Point lpPoint);
        [DllImport("user32.dll")]
        private static extern int RegisterHotKey(int hwnd, int id, int fsModifiers, int vk);
        [DllImport("user32.dll")]
        private static extern int UnregisterHotKey(int hwnd, int id);
        [DllImport("kernel32.dll", EntryPoint = "LoadLibrary")]
        private extern static int LoadLibrary(string librayName);
        [DllImport("kernel32.dll", EntryPoint = "GetProcAddress", CharSet = CharSet.Ansi)]
        private extern static IntPtr GetProcAddress(int hwnd, string procedureName);
        [DllImport("kernel32.dll", EntryPoint = "FreeLibrary")]
        private extern static bool FreeLibrary(int hModule);

        
        private delegate void _DD_mov(int x, int y);
        private _DD_mov DD_mov;

        private delegate void _DD_btn(int param);
        private _DD_btn DD_Btn;

        //public delegate void _DD_movR(int dx, int dy);
        //public _DD_movR DD_movR;

        private delegate void _DD_key(int param1, int param2);
        private _DD_key DD_key;

        private delegate void _DD_whl(int param);
        private _DD_whl DD_whl;

        private delegate int _DD_todc(int key);
        private _DD_todc DD_todc;

        private delegate int _DD_str(string str);
        private _DD_str DD_str;

        int hModule = 0x0;
        IntPtr pFuncAddr = IntPtr.Zero;

        public Form1()
        {

            InitializeComponent();
            txtPosX.Text = "0";//초기화
            txtPosY.Text = "0";
            txtTime.Text = "0";
            txtRand1.Text = "0";
            txtRand2.Text = "0";
            txtLoop.Text = "0";
            radioNone.Select();
            cmbAction.SelectedIndex = 0;
            cmbButton.SelectedIndex = 0;

            try
            {
                if (Environment.Is64BitOperatingSystem)
                {

                    if (Environment.Is64BitProcess)
                    {
                        hModule = LoadLibrary(".\\dll\\x64\\ddx64.64.dll");

                    }
                    else
                    {
                        hModule = LoadLibrary(".\\dll\\x64\\ddx64.32.dll");

                    }

                }

                else
                {
                    hModule = LoadLibrary(".\\dll\\x86\\ddx32.dll");
                }

                pFuncAddr = GetProcAddress(hModule, "DD_mov");
                DD_mov = (_DD_mov)Marshal.GetDelegateForFunctionPointer(pFuncAddr, typeof(_DD_mov));

                pFuncAddr = GetProcAddress(hModule, "DD_btn");
                DD_Btn = (_DD_btn)Marshal.GetDelegateForFunctionPointer(pFuncAddr, typeof(_DD_btn));

                pFuncAddr = GetProcAddress(hModule, "DD_whl");
                DD_whl = (_DD_whl)Marshal.GetDelegateForFunctionPointer(pFuncAddr, typeof(_DD_whl));

                pFuncAddr = GetProcAddress(hModule, "DD_key");
                DD_key = (_DD_key)Marshal.GetDelegateForFunctionPointer(pFuncAddr, typeof(_DD_key));

                pFuncAddr = GetProcAddress(hModule, "DD_todc");
                DD_todc = (_DD_todc)Marshal.GetDelegateForFunctionPointer(pFuncAddr, typeof(_DD_todc));

                pFuncAddr = GetProcAddress(hModule, "DD_str");
                DD_str = (_DD_str)Marshal.GetDelegateForFunctionPointer(pFuncAddr, typeof(_DD_str));
                //MessageBox.Show("DLL을 로딩 완료.");

                RegisterHotKey((int)this.Handle, 0, 0x0, (int)Keys.F5);

                RegisterHotKey((int)this.Handle, 1, 0x0, (int)Keys.F6);

                RegisterHotKey((int)this.Handle, 2, 0x0, (int)Keys.F7);
            }
            catch
            {
                MessageBox.Show("DLL을 찾을수 없습니다.");
                Environment.Exit(0);//프로그램 종료.
            }
            string sDirPath;
            sDirPath = Application.StartupPath + ".\\preset";
            DirectoryInfo di = new DirectoryInfo(sDirPath);
            if (di.Exists == false)// Preset 디렉토리가 없으면 새로 만듭니다.
            {
                di.Create();
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

            UnregisterHotKey((int)this.Handle, 0);
            UnregisterHotKey((int)this.Handle, 1);
            UnregisterHotKey((int)this.Handle, 2);
            FreeLibrary(hModule);
            MessageBox.Show("잘가...");

        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == (int)0x312) //핫키가 눌러지면 312 정수 메세지를 받게됨
            {
                if (m.WParam == (IntPtr)0x0) // 그 키의 ID가 0이면
                {
                    Point pt = new Point();
                    GetCursorPos(ref pt);
                    label3.Text = "X :" + pt.X.ToString() + "  Y :" + pt.Y.ToString();// 위치 분석
                    txtPosX.Text = pt.X.ToString();
                    txtPosY.Text = pt.Y.ToString();
                }
                if (m.WParam == (IntPtr)0x1)
                {
                    Point pt = new Point();
                    GetCursorPos(ref pt);
                    if (radioRect.Checked)
                    {
                        int X = pt.X - Convert.ToInt32(txtPosX.Text);
                        int Y = pt.Y - Convert.ToInt32(txtPosY.Text);
                        txtFactor.Text = X + ":" + Y;
                    }
                    else if (radioCircle.Checked)
                    {
                        int X = pt.X - Convert.ToInt32(txtPosX.Text);
                        int Y = pt.Y - Convert.ToInt32(txtPosY.Text);
                        int Radius = (int)Math.Sqrt((double)X * X + Y * Y);

                        txtFactor.Text = "" + Radius;

                    }

                }
                if (m.WParam == (IntPtr)0x2)
                {
                    if (!isRunning)
                    {
                        PresetRun();
                    }
                    else
                    {
                        PresetStop();
                    }
                }
            }
            else if(m.Msg == (int)0x216)
            {

            }


        }

        private ListViewItem ListInsert(string Key, string Interval, string Coordinate, string Action, string Code)
        {
            ListViewItem listViewItem = new ListViewItem(Key);//listView1.Items.Add(Key);
            listViewItem.SubItems.Add(Interval);
            listViewItem.SubItems.Add(Coordinate);
            listViewItem.SubItems.Add(Action);
            listViewItem.SubItems.Add(Code);

            return listViewItem;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)//라디오버튼 None일시 txt박스 잠금
        {
            if (radioNone.Checked == true)
            {
                txtFactor.Enabled = false;
                txtFactor.Text = "";
            }
            else
                txtFactor.Enabled = true;
        }

        private void txtKey_KeyDown(object sender, KeyEventArgs e)//키보드 키 저장
        {
            e.SuppressKeyPress = true;
            txtKey.Text = e.KeyCode.ToString();
            int num = (int)new KeysConverter().ConvertFromString(Convert.ToString((object)e.KeyCode));
            txtCode.Text = num.ToString();
        }

        private void cmbAction_TextChanged(object sender, EventArgs e)
        {
            if (cmbAction.Text == "Loop")
            {
                txtSubLoop.Enabled = true;
            }
            else
            {
                txtSubLoop.Enabled = false;
            }
        }


        private void btnAdd_Click(object sender, EventArgs e)
        {
            string Key = "", Interval = "", Coord = "", Act = "", Code = "";

            Interval = txtTime.Text + "|" + txtRand1.Text + "~" + txtRand2.Text;

            if (cmbAction.Text != "Nothing" && cmbAction.Text != "Loop" && cmbAction.Text != "LoopEnd")
            {
                if (TabControl.SelectedIndex == 0)//KeyBoard
                {
                    if (txtCode.Text == "" || cmbAction.Text == "")
                    {
                        MessageBox.Show("키, 액션을 입력해주세요.");
                    }
                    else if (cmbAction.Text == "Move")
                    {
                        MessageBox.Show("키보드 Move는 아니지않나요...");
                    }
                    else
                    {
                        Key = txtKey.Text;
                        Act = cmbAction.Text;
                        Code = txtCode.Text;
                        Coord = "";
                        listView1.Items.Add(ListInsert(Key, Interval, Coord, Act, Code));
                    }

                }
                if (TabControl.SelectedIndex == 1)//Mouse   
                {
                    if (cmbAction.Text == "")
                    {
                        MessageBox.Show("액션을 입력해주세요.");
                    }
                    else
                    {

                        if (cmbAction.Text == "Move")//Move일때만 좌표 설정하기
                        {
                            String StrFactor = "|";
                            if (radioNone.Checked)
                            {
                                StrFactor += "N|";
                            }
                            else if (radioCircle.Checked)
                            {
                                StrFactor += "C|" + txtFactor.Text;
                            }
                            else if (radioRect.Checked)
                            {
                                StrFactor += "R|" + txtFactor.Text;
                            }
                            Coord = txtPosX.Text + ':' + txtPosY.Text + StrFactor;
                            Act = cmbAction.Text;
                            Code = "";
                            Key = "";
                            listView1.Items.Add(ListInsert(Key, Interval, Coord, Act, Code));
                        }
                        else//아니면 None
                        {
                            if (cmbButton.Text != "")
                            {
                                Key = cmbButton.Text;
                                Act = cmbAction.Text;
                                Coord = "";
                                Code = "";
                                if (Key == "MWUP" || Key == "MWDOWN")
                                {
                                    Act = "";
                                }
                                listView1.Items.Add(ListInsert(Key, Interval, Coord, Act, Code));
                            }
                            else
                            {
                                MessageBox.Show("키를 입력해주세요.");
                            }
                        }


                    }

                }
            }
            else
            {

                Key = "";
                Coord = "";
                Act = cmbAction.Text;
                Code = "";
                if (Act.Substring(0, 4) == "Loop")
                {
                    Key = Act;
                    Act = "";
                    Interval = "";
                    Key += "|" + txtSubLoop.Text;
                }
                listView1.Items.Add(ListInsert(Key, Interval, Coord, Act, Code));
            }

        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                ListViewItem select = listView1.SelectedItems[0];
                int index = select.Index;

                if (index == 0)//순환구조일시 주석 해제
                {
                    //listView1.Items.Remove(select);
                    //listView1.Items.Insert(count - 1, select);

                }
                else
                {
                    listView1.Items.Remove(select);
                    listView1.Items.Insert(index - 1, select);

                }
                listView1.Select();
            }
        }
        private void btnDown_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                ListViewItem select = listView1.SelectedItems[0];
                int index = select.Index;
                int count = listView1.Items.Count;

                if (index == count - 1)
                {
                    //listView1.Items.Remove(select);
                    //listView1.Items.Insert(0, select);

                }
                else
                {
                    listView1.Items.Remove(select);
                    listView1.Items.Insert(index + 1, select);
                }
                listView1.Select();
            }

        }


        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                foreach (ListViewItem K in listView1.SelectedItems)
                {
                    listView1.Items.Remove(K);
                }
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                foreach (ListViewItem K in listView1.SelectedItems)
                {
                    ListViewItem V = (ListViewItem)K.Clone();
                    listView1.Items.Insert(listView1.Items.Count, V);
                }
            }
        }
        /////////////////////////////////////////////////편집///////////////////////////////////////////////////////
        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                ListViewItem select = listView1.SelectedItems[0];
                Form2 F = new Form2();
                F.FormSendEvent += new Form2.FormSendDataHandler(Edit);


                F.SetTextBox(
                      select.SubItems[0].Text
                    , select.SubItems[1].Text
                    , select.SubItems[2].Text
                    , select.SubItems[3].Text
                    , select.SubItems[4].Text);
                F.ShowDialog();
            }
        }
        private void Edit(string msg)
        {
            ListViewItem select = listView1.SelectedItems[0];
            string[] Edits = msg.Split('/');
            for (int i = 0; i < Edits.Length; i++)
            {
                select.SubItems[i].Text = Edits[i];
            }
        }


        /////////////////////////////////////////////////저장///////////////////////////////////////////////////////
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count <= 0)
            {
                MessageBox.Show("내용이 없어요.");
            }
            else
            {

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.InitialDirectory = Application.StartupPath + "\\preset";
                saveFileDialog.Title = "저장";
                saveFileDialog.DefaultExt = "txt";
                saveFileDialog.Filter = "Text(*.txt)|*.txt";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string path = saveFileDialog.FileName;
                    string save = "";

                    foreach (ListViewItem K in listView1.Items)
                    {
                        for (int i = 0; i < 5; i++)
                        {

                            save += K.SubItems[i].Text;
                            if (i < 4)
                            {
                                save += '/';
                            }
                        }
                        save += '\n';
                    }
                    try
                    {
                        StreamWriter streamWriter = new StreamWriter(path, false, Encoding.GetEncoding("euc-kr"));
                        streamWriter.Write(save);
                        streamWriter.Dispose();
                        streamWriter.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }

            }
        }
        /////////////////////////////////////////////////불러오기///////////////////////////////////////////////////////
        private void btnLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "불러오기";
            openFileDialog.InitialDirectory = Application.StartupPath + "\\preset";
            openFileDialog.DefaultExt = "txt";
            openFileDialog.Filter = "Text(*.txt)|*.txt";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog.FileName;
                string load = "";
                try
                {
                    StreamReader streamReader = new StreamReader(path);
                    load = streamReader.ReadToEnd();
                    streamReader.Dispose();
                    streamReader.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                try
                {
                    String[] loadSplit = load.Split('\n');
                    foreach (String loadItem in loadSplit)
                    {
                        if (loadItem != "")
                        {
                            String[] Items = loadItem.Split('/');
                            listView1.Items.Add(ListInsert(Items[0], Items[1], Items[2], Items[3], Items[4]));
                        }

                    }
                }
                catch
                {
                    MessageBox.Show("잘못된 파일형식입니다.");
                }
            }

        }
        /////////////////////////////////////////////////실행///////////////////////////////////////////////////////


        private void ChangeBool(bool Bool)
        {
            btnAdd.Enabled = Bool;
            btnCopy.Enabled = Bool;
            btnDown.Enabled = Bool;
            btnEdit.Enabled = Bool;
            btnUp.Enabled = Bool;
            btnSave.Enabled = Bool;
            btnLoad.Enabled = Bool;
            btnRemove.Enabled = Bool;
        }
        bool isRunning = false;

        private void btnStart_Click(object sender, EventArgs e)
        {

            if (!isRunning)
            {
                PresetRun();
            }
            else
            {
                PresetStop();
            }

        }
        private static DateTime randDelay(String Interval)
        {
            String[] Times = Interval.Split('|');
            Random rand = new Random();
            int Time = Convert.ToInt32(Times[0]);
            int rand1 = Convert.ToInt32(Times[1].Split('~')[0]);
            int rand2 = Convert.ToInt32(Times[1].Split('~')[1]);
            int MS = (Time + rand.Next(rand1, rand2));

            DateTime ThisMoment = DateTime.Now;
            TimeSpan duration = new TimeSpan(0, 0, 0, 0, MS);
            DateTime AfterWards = ThisMoment.Add(duration);

            while (AfterWards >= ThisMoment)
            {
                System.Windows.Forms.Application.DoEvents();
                ThisMoment = DateTime.Now;
            }

            return DateTime.Now;
        }


        List<ListViewItem> Collection = new List<ListViewItem>();
        private void ParseList(List<ListViewItem> List, int Loop, bool Sub)
        {
            for (int i = 0; i < Loop || (Loop == -1 && isRunning); i++)
            {
                int index = 1;
                int LoopCount = 0;
                int EndCount = 0;
                int SubLoop = 0;
                bool Collect = false;
                
                foreach (ListViewItem K in List)
                {
                    if (isRunning)//실행중일때만 한다.
                    {
                        String Key = K.SubItems[0].Text;
                        String Interval = K.SubItems[1].Text;
                        String Coordinate = K.SubItems[2].Text;
                        String Act = K.SubItems[3].Text;
                        String Code = K.SubItems[4].Text;



                        if (Key.Split('|')[0] == "Loop")
                        {
                            LoopCount++;
                            if (Collect == true)
                            {
                                Collection.Add(K); // 콜렉션이
                            }
                            else if (Collect == false)
                            {
                                Collection.Clear();
                                SubLoop = Convert.ToInt32(Key.Split('|')[1]);
                                Collect = true;
                            }

                            continue;
                        }
                        else if (Key == "LoopEnd")
                        {
                            EndCount++;
                            if (EndCount == LoopCount)
                            {
                                Collect = false;
                                LoopCount = 0;
                                EndCount = 0;
                                ParseList(Collection, SubLoop, true);

                            }
                            else
                            {
                                //   MessageBox.Show(Key);
                                Collection.Add(K);
                            }
                            continue;
                        }
                        else if (Collect == true)
                        {
                            //MessageBox.Show(K.Text);
                            Collection.Add(K);
                            continue;
                        }
                        else
                        {
                            if (Act == "Nothing")//아무일도 안합니다.
                            {

                            }
                            else
                            {

                                if (Code == "")//Mouse
                                {
                                    if (Act == "Move")
                                    {
                                        String[] Coords = Coordinate.Split('|');
                                        int X = Convert.ToInt32(Coords[0].Split(':')[0]);
                                        int Y = Convert.ToInt32(Coords[0].Split(':')[1]);
                                        Random rand = new Random((int)DateTime.Now.Ticks);
                                        int randX = 0;
                                        int randY = 0;
                                        if (Coords[1] == "N")
                                        {
                                            //Nothing
                                        }
                                        else if (Coords[1] == "C")
                                        {

                                            int Rad = rand.Next(0, Convert.ToInt32(Coords[2]));
                                            int Degree = rand.Next(0, 360);

                                            randX = (int)(Math.Cos(Math.PI * Degree / 180) * Rad);//매개변수로 랜덤 원 그리기
                                            randY = (int)(Math.Sin(Math.PI * Degree / 180) * Rad);


                                        }
                                        else if (Coords[1] == "R")
                                        {
                                            randX = rand.Next(0, Convert.ToInt32(Coords[2].Split(':')[0]));//랜덤 박스 그리기
                                            randY = rand.Next(0, Convert.ToInt32(Coords[2].Split(':')[1]));
                                        }
                                        DD_mov(X + randX, Y + randY);

                                    }
                                    else if (Act == "Press" || Act == "Pull")
                                    {

                                        int Value = 1;

                                        Value = Act == "Pull" ? Value << 1 : Value; // Value는 Press일때 1, Pull일때 2를 갖는다.

                                        if (Key == "LB")
                                        {
                                        }
                                        else if (Key == "RB")
                                        {
                                            Value = Value << 2;
                                        }
                                        else if (Key == "MB")
                                        {
                                            Value = Value << 4;
                                        }
                                        DD_Btn(Value);//Value Press
                                    }
                                    else if (Act == "")
                                    {
                                        string SubKey;
                                        int SubL;

                                        SubKey = Key.Split('/')[0];
                                        try
                                        {
                                            SubL = Convert.ToInt32(Key.Split('/')[1]);
                                        }
                                        catch
                                        {
                                            SubL = 1;
                                        }

                                        if (SubKey == "MWUP")
                                        {
                                            DD_whl(1);
                                        }
                                        else if (SubKey == "MWDOWN")
                                        {
                                            DD_whl(2);
                                        }
                                    }
                                }
                                else
                                {
                                    int presspull = 1;
                                    if (Act == "Press")
                                    {

                                    }
                                    else
                                    {
                                        presspull = 2;
                                    }
                                    DD_key(DD_todc(Convert.ToInt32(Code)), presspull);
                                }

                            }

                        }

                        if (Loop == -1)
                        {
                            prograsslabel.Text = "무한반복 " + index + "/" + List.Count;
                        }
                        else
                        {
                            prograsslabel.Text = (i + 1) + "/" + Loop + " " + index + "/" + List.Count;//   현재루프/총 루프 + 현재 실행 / 리스트 실행
                        }
                        index++;
                        randDelay(Interval); // 딜레이를 넣는다.


                    }
                    else//실행중 아니면 브레이크
                    {
                        break;
                    }
                }

            }
            if (!Sub)
            {
                PresetStop();
            }
        }


        private void PresetRun()
        {
            if (listView1.Items.Count <= 0)
            {
                return;
            }
            isRunning = true;
            Text = "Murbong's Macro (Run)";
            btnStart.Text = "Stop";
            int loop;
            try
            {
                loop = Convert.ToInt32(txtLoop.Text);
            }
            catch
            {
                txtLoop.Text = "100";
                loop = 100;
            }
            txtLoop.ReadOnly = true;
            ChangeBool(false);
            List<ListViewItem> items = new List<ListViewItem>();
            foreach (ListViewItem K in listView1.Items)
            {
                items.Add(K);
            }

            ParseList(items, loop, false);

        }
        private void PresetStop()
        {
            btnStart.Text = "Start";
            Text = "Murbong's Macro";
            txtLoop.ReadOnly = false;
            ChangeBool(true);
            isRunning = false;
            prograsslabel.Text = "Idle";

        }

        private void OnlyNumAndMinus(object sender, KeyPressEventArgs e)
        {
            bool isValidInput = false;
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                if (e.KeyChar == '-') isValidInput = true;

                if (isValidInput == false) e.Handled = true;
            }

            if (e.KeyChar == '-' && (!string.IsNullOrEmpty((sender as TextBox).Text.Trim()) || (sender as TextBox).Text.IndexOf('-') > -1)) e.Handled = true;

        }


        private void txtTime_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyNumAndMinus(sender, e);
        }
        private void txtRand1_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyNumAndMinus(sender, e);
        }
        private void txtRand2_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyNumAndMinus(sender, e);
        }
        private void txtLoop_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyNumAndMinus(sender, e);
        }


    }
}
