using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Net;

namespace Conector
{
    public partial class FormConector : Form
    {
        private readonly string pathVnc = Environment.ExpandEnvironmentVariables("C:\\Program Files\\uvnc bvba\\UltraVNC\\vncviewer.exe");
        public FormConector()
        {
            InitializeComponent();
        }

        private void HabilitarControles(bool estado)
        {
            switch (estado)
            {
                case false:
                    radioBos.Checked = false;
                    radioBos.Enabled = false;
                    radioStationManager.Checked = false;
                    radioStationManager.Enabled = false;
                    ListaPcManager.Items.Clear();
                    txtBoxBos.Text = string.Empty;
                    WetPos.Enabled = false;
                    WetPos.Checked = false;
                    DryPos.Enabled = false;
                    DryPos.Checked = false;
                    txtBoxEasyPay.Text = string.Empty;
                    EasyPay.Enabled = false;
                    EasyPay.Checked = false;
                    ComboPos.Items.Clear();
                    ComboPos.Enabled = false;

                    comboNuc.Items.Clear();
                    comboNuc.Enabled = false;
                    NUC.Checked = false;
                    NUC.Enabled = false;

                    break;
                case true:
                    radioBos.Enabled = true;
                    radioStationManager.Enabled = true;
                    ListaPcManager.Enabled = true;
                    if (!String.IsNullOrEmpty(ObtenerDetalleFile("DryPOS")))
                        DryPos.Enabled = true;
                    if (!String.IsNullOrEmpty(ObtenerDetalleFile("WetPOS")))
                        WetPos.Enabled = true;
                    if (!String.IsNullOrEmpty(ObtenerDetalleFile("EasyPay")))
                        EasyPay.Enabled = true;
                    if (!String.IsNullOrEmpty(ObtenerDetalleFile("NUC")))
                        NUC.Enabled = true;
                    break;
                default:
                    break;
            }
        }

        private void Buscar()
        {
            HabilitarControles(false);
            //Obtener nro de file
            String File = txtBoxFile.Text;

            if (EsFileValido(File) &&
                IsNumber(File) &&
                !string.IsNullOrEmpty(File) &&
                !string.IsNullOrWhiteSpace(File))
            {
                try
                {
                    OleDbConnection conexionExcel = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\\ENEX\\Excel\\ENEX_Stations.xls';Extended Properties=Excel 8.0;");
                    OleDbDataAdapter consulta = new OleDbDataAdapter(string.Format("select [ID],[File],[Direccion],[Type],[Ciudad],[BosPC],[Password],[ManagerPC],[WetPOS],[DryPOS], [IP] , [User_Eds], [EasyPay], [Modelo_Surtidor], [NUC1],[NUC2] from [Station List$] where [FILE]={0} order by [Type] DESC", File), conexionExcel);
                    consulta.TableMappings.Add("Table", "Net-informations.com");
                    DataSet dataSet = new DataSet();
                    consulta.Fill(dataSet);
                    dataEstaciones.DataSource = dataSet.Tables[0];
                    conexionExcel.Close();
                    switch (dataSet.Tables[0].Rows.Count)
                    {
                        case 0:
                            HabilitarControles(false);
                            MessageBox.Show("No se encuentra File", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            break;
                        default:
                            HabilitarControles(true);
                            break;
                    }
                }
                catch (OleDbException)
                {
                    MessageBox.Show("No existe la BBDD de las EDS. Cierre el programa e inicie nuevamente .", "Error de lectura de BBDD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
                MessageBox.Show("Debes ingresar un Número de File", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void TxtBoxFile_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                e.Handled = true;
                //e.SuppressKeyPress = true;
                Buscar();
            }
        }

        private void BtnBuscar_Click(object sender, EventArgs e)
        {
            Buscar();
        }

        private static bool EsFileValido(string File)
        {
            try
            {
                ushort val;
                val = UInt16.Parse(File);
                if (val > 0 && val < 1000)
                    return true;
            }
            catch (OverflowException)
            {
                return false;
            }

            catch (FormatException)
            {
                return false;
            }
            return false;
        }

        private void BtnConectar_Click(object sender, EventArgs e)
        {
            string File = txtBoxFile.Text;
            if (EsFileValido(File) &&
                IsNumber(File) &&
                !String.IsNullOrEmpty(File) &&
                !string.IsNullOrWhiteSpace(File))
            {
                if (radioBos.Checked)
                {
                    String IpFile = ObtenerDetalleFile("IP"); // Obtengo la ip del BOS
                    String Password = ObtenerDetalleFile("Password"); // Contraseña BOS Escritorio remoto
                    String NroFile = ObtenerDetalleFile("File");  //obtengo nro de file
                    String User_File = ObtenerDetalleFile("User_Eds"); //obtengo User_Eds

                    string pathEscritorioRemoto = Environment.ExpandEnvironmentVariables("%SystemRoot%\\system32\\mstsc.exe");
                    if (!String.IsNullOrEmpty(pathEscritorioRemoto) &&
                        !String.IsNullOrWhiteSpace(pathEscritorioRemoto) &&
                        !String.IsNullOrEmpty(IpFile) &&
                        !String.IsNullOrEmpty(Password) &&
                        !String.IsNullOrEmpty(NroFile)
                        )
                    {
                        //agrega las credenciales del File al almacen de credenciales                    
                        AgregarCredencial(IpFile, NroFile, Password, User_File);
                        //iniciar proceso de escritorio Remoto
                        Process escritorioRemoto = new Process();
                        escritorioRemoto.StartInfo.FileName = pathEscritorioRemoto;
                        //Se le pasan los argumentos a mstsc.exe                    
                        escritorioRemoto.StartInfo.Arguments = string.Format("/v:{0}", IpFile);
                        escritorioRemoto.Start();
                    }
                    //cerrar proceso
                    //escritorioRemoto.Close();
                }
                if (radioStationManager.Checked)
                {
                    if (!string.IsNullOrEmpty(ListaPcManager.Text))
                        ConectorVnc(ListaPcManager.Text);
                    else
                        MessageBox.Show("Debes seleccionar un PC Manager", "Selección de Station Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if (DryPos.Checked)
                {
                    if (!string.IsNullOrEmpty(ComboPos.Text))
                        ConectorVnc(ComboPos.Text);
                    else
                        MessageBox.Show("Debes seleccionar un DryPOS", "Selección de DryPOS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if (WetPos.Checked)
                {
                    if (!string.IsNullOrEmpty(ComboPos.Text))
                        ConectorVnc(ComboPos.Text);
                    else
                        MessageBox.Show("Debes seleccionar un WetPOS", "Selección de WetPOS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if (EasyPay.Checked)
                {
                    Process cmdKey = new Process();
                    cmdKey.StartInfo.FileName = Environment.ExpandEnvironmentVariables("%SystemRoot%\\system32\\cmdkey.exe");
                    cmdKey.StartInfo.UseShellExecute = false;
                    cmdKey.StartInfo.CreateNoWindow = true;
                    cmdKey.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cmdKey.StartInfo.Arguments = string.Format("/generic:TERMSRV/{0} /user:{1} /pass:{2}", txtBoxEasyPay.Text, "", "1");
                    cmdKey.Start();
                    cmdKey.Close();
                    string pathEscritorioRemoto = Environment.ExpandEnvironmentVariables("%SystemRoot%\\system32\\mstsc.exe");
                    //iniciar proceso de escritorio Remoto
                    Process escritorioRemoto = new Process();
                    escritorioRemoto.StartInfo.FileName = pathEscritorioRemoto;
                    //Se le pasan los argumentos a mstsc.exe                    
                    escritorioRemoto.StartInfo.Arguments = string.Format("/v:{0}", txtBoxEasyPay.Text);
                    escritorioRemoto.Start();
                }
                if(NUC.Checked)
                {
                    if (!string.IsNullOrEmpty(comboNuc.Text))
                    {
                        string IpNuc = comboNuc.Text.Split('/')[1].Split(' ')[2];

                        Process cmdKey = new Process();
                        cmdKey.StartInfo.FileName = Environment.ExpandEnvironmentVariables("%SystemRoot%\\system32\\cmdkey.exe");
                        cmdKey.StartInfo.UseShellExecute = false;
                        cmdKey.StartInfo.CreateNoWindow = true;
                        cmdKey.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        //agrega la credencial como local para el usuario .\orpak
                        cmdKey.StartInfo.Arguments = string.Format("/generic:TERMSRV/{0} /user:{1}\\{2} /pass:{3}", IpNuc, Environment.MachineName, "", "");
                        cmdKey.Start();
                        cmdKey.Close();
                        string pathEscritorioRemoto = Environment.ExpandEnvironmentVariables("%SystemRoot%\\system32\\mstsc.exe");
                        //iniciar proceso de escritorio Remoto
                        Process escritorioRemoto = new Process();
                        escritorioRemoto.StartInfo.FileName = pathEscritorioRemoto;
                        //Se le pasan los argumentos a mstsc.exe                    
                        escritorioRemoto.StartInfo.Arguments = string.Format("/v:{0}", IpNuc);
                        escritorioRemoto.Start();
                    }                        
                    else
                        MessageBox.Show("Debes seleccionar un NUC", "Selección de NUC", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Debes ingresar un Número de File", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ConectorVnc(string ip)
        {
            try
            {
                Process Vnc = new Process();
                Vnc.StartInfo.FileName = pathVnc;
                Vnc.StartInfo.Arguments = ip;
                Vnc.Start();
            }
            catch (Exception)
            {
                MessageBox.Show("No tienes VNC en el equipo", "vnc", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void AgregarCredencial(String ip, String NroFile, String password, String User_Eds)
        {
            Process cmdKey = new Process();
            cmdKey.StartInfo.FileName = Environment.ExpandEnvironmentVariables("%SystemRoot%\\system32\\cmdkey.exe");
            cmdKey.StartInfo.UseShellExecute = false;
            cmdKey.StartInfo.CreateNoWindow = true;
            cmdKey.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

            //Caso donde el file todavía no tiene usuario EDS\XXX
            if (User_Eds.Length == 0)
            {
                //temporal para modificar el nro del file segun su largo (si es file 5 vs file 100)
                var temp = "";
                //formatea usuario de acuerdo al largo de nro de file
                switch (NroFile.Length)
                {
                    case 1:
                        temp = string.Format("CLBO000{0}\\CLAD000{0}", NroFile);
                        break;
                    case 2:
                        temp = string.Format("CLBO00{0}\\CLAD00{0}", NroFile);
                        break;
                    case 3:
                        temp = string.Format("CLBO0{0}\\CLAD0{0}", NroFile);
                        break;
                    default:
                        break;
                }
                cmdKey.StartInfo.Arguments = string.Format("/generic:TERMSRV/{0} /user:{1} /pass:{2}", ip, temp, password);
                cmdKey.Start();
                cmdKey.Close();
            }
            else // BOS En dominio eds
            {
                cmdKey.StartInfo.Arguments = string.Format("/generic:TERMSRV/{0} /user:{1} /pass:{2}", ip, User_Eds, password);
                cmdKey.Start();
                cmdKey.Close();
            }
        }

        private void FormConector_Load(object sender, EventArgs e)
        {
            string[] args = Environment.GetCommandLineArgs();
            switch (args.Length)
            {
                //Se inicia solo el .exe y sí hay vpn y red (situacion normal)
                case 1:
                    //copia ENEX_Stations.xls desde 10.34.128.119\\Soporte TI
                    Boolean _error = CopiarExcel();
                    if (_error)
                    {
                        //Limpia ListaPcManager
                        ListaPcManager.Items.Clear();
                        ListaPcManager.ResetText();
                        //LLena datagrid
                        AgregarTodoFile();
                    }
                    else
                    {
                        txtBoxFile.Enabled = false;
                        btnBuscar.Enabled = false;
                        btnConectar.Enabled = false;
                        BtnTodo.Enabled = false;
                        label2.Visible = true;
                        label3.Visible = true;
                    }
                    break;
                case 2: // Parametro 1 : conector.exe parametro 2: offline
                    if (args[1].Equals("offline"))
                    {
                        //Limpia ListaPcManager
                        ListaPcManager.Items.Clear();
                        ListaPcManager.ResetText();
                        //LLena datagrid
                        AgregarTodoFile();
                    }
                    else
                    {
                        MessageBox.Show("Parámetro Incorrecto", "param", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        txtBoxFile.Enabled = false;
                        btnBuscar.Enabled = false;
                        btnConectar.Enabled = false;
                        BtnTodo.Enabled = false;
                    }
                    /*foreach (string arg in args)
                    {if (arg.Equals("offline")){}}*/
                    break;
                default:
                    MessageBox.Show("Parámetro Incorrecto", "param", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtBoxFile.Enabled = false;
                    btnBuscar.Enabled = false;
                    btnConectar.Enabled = false;
                    BtnTodo.Enabled = false;
                    break;
            }
        }
        private Boolean CopiarExcel()
        {
            //Process process = new Process();
            //process.StartInfo.FileName = Environment.ExpandEnvironmentVariables("%SystemRoot%\\system32\\cmdkey.exe");
            //process.StartInfo.Arguments = "/generic:TERMSRV/10.20.0.7 /user:clad0002 /pass:F002.112015.A1";
            //process.Start();
            /*Process cmdKey = new Process();
            cmdKey.StartInfo.FileName = Environment.ExpandEnvironmentVariables("%SystemRoot%\\system32\\cmdkey.exe");
            cmdKey.StartInfo.UseShellExecute = false;
            cmdKey.StartInfo.CreateNoWindow = true;
            cmdKey.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

         

            string path = "C:\\ENEX\\Excel\\";
            if (!Directory.Exists(path))
               Directory.CreateDirectory(path);

            string remoteUri = "";
            string destFileName = "C:\\ENEX\\Excel\\ENEX_Stations.xls";
            // Create a new WebClient instance.
            WebClient myWebClient = new WebClient();
            // Download the Web resource and save it into the current filesystem folder.
            myWebClient.DownloadFile(remoteUri, destFileName);

            //string sourceFileName = "\\10.20.62.7\\Users\\Clad0124\\AppData\\Local\\Packages\\windows_ie_ac_002\\ENEX_Stations.xls"; // file 124
            //string destFileName = "C:\\ENEX\\Excel\\ENEX_Stations.xls";
            //string path = "C:\\ENEX\\Excel\\";

            //BBDD remota existe
            
            /*
            if (!Directory.Exists("\\\\10.20.62.7\\Users\\Clad0124\\AppData\\Local\\Packages\\windows_ie_ac_002\\") || !File.Exists(sourceFileName))
            {
                MessageBox.Show("Error al acceder a la BBDD Remota", "sin acceso a EDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }*/
            //Crea directorio local ENEX\\Excel
            /*if (!Directory.Exists(path))
                Directory.CreateDirectory(path);*/
            try
            {
                //File.Copy(sourceFileName, destFileName, overwrite: true);
            }
            catch (UnauthorizedAccessException)
            {
                //MessageBox.Show("No se pudo acceder a datos de eds. Borrar carpeta C:\\ENEX", "Error al copiar data de EDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //return false;
            }
            catch (IOException)
            {
                //MessageBox.Show("Debes conectar a la VPN", "Sin conexión remota", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //return false;
            }
            finally { }
            return true;
        }

        private void AgregarTodoFile()
        {
            //Conectando a archivo Excel
            OleDbConnection ConexionExcel = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\\ENEX\\Excel\\ENEX_Stations.xls';Extended Properties=Excel 8.0;");
            //Consulta select a excel
            OleDbDataAdapter Consulta = new OleDbDataAdapter("select [ID], [File],[Direccion],[Type],[Ciudad],[BosPC],[Password],[ManagerPC],[WetPOS],[DryPOS],[IP],[User_Eds],[EasyPay],[Modelo_Surtidor],[NUC1], [NUC2] from [Station List$]", ConexionExcel);
            Consulta.TableMappings.Add("Table", "Net-informations.com");
            DataSet dataSet = new DataSet();
            Consulta.Fill(dataSet);
            dataEstaciones.DataSource = dataSet.Tables[0];
            ConexionExcel.Close();
            for (int i = 0; i < 16; i++)            
                dataEstaciones.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataEstaciones.Columns["ID"].Visible = false;
            dataEstaciones.Columns["IP"].Visible = false;
            dataEstaciones.Columns["NUC1"].Visible = false;
            dataEstaciones.Columns["NUC2"].Visible = false;
        }

        private string ObtenerDetalleFile(string Variable)
        {
            var Detalle = String.Empty;
            try
            {
                int rowIndex = dataEstaciones.SelectedCells[0].RowIndex;

                DataGridViewRow dataGridViewRow = dataEstaciones.Rows[rowIndex];
                switch (Variable)
                {
                    case "ManagerPC":
                        Detalle = Convert.ToString(dataGridViewRow.Cells["ManagerPC"].Value);
                        break;
                    case "File":
                        Detalle = Convert.ToString(dataGridViewRow.Cells["File"].Value);
                        break;
                    case "IP":
                        Detalle = Convert.ToString(dataGridViewRow.Cells["IP"].Value);
                        break;
                    case "Password":
                        Detalle = Convert.ToString(dataGridViewRow.Cells["Password"].Value);
                        break;
                    case "BosPC":
                        Detalle = Convert.ToString(dataGridViewRow.Cells["BosPC"].Value);
                        break;
                    case "WetPOS":
                        Detalle = Convert.ToString(dataGridViewRow.Cells["WetPOS"].Value);
                        break;
                    case "DryPOS":
                        Detalle = Convert.ToString(dataGridViewRow.Cells["DryPOS"].Value);
                        break;
                    case "User_Eds":
                        Detalle = Convert.ToString(dataGridViewRow.Cells["User_Eds"].Value);
                        break;
                    case "EasyPay":
                        Detalle = Convert.ToString(dataGridViewRow.Cells["EasyPay"].Value);
                        break;
                    case "NUC":
                        Detalle = Convert.ToString(dataGridViewRow.Cells["NUC1"].Value) + Convert.ToString(dataGridViewRow.Cells["NUC2"].Value);
                        break;
                    default: break;
                }
            }
#pragma warning disable CS0168 // La variable 'e' se ha declarado pero nunca se usa
            catch (System.ArgumentOutOfRangeException e)
#pragma warning restore CS0168 // La variable 'e' se ha declarado pero nunca se usa
            {
            }
            return Detalle;
        }

        private void radioStationManager_CheckedChanged(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txtBoxFile.Text))
            {
                //Limpiar lista pc manager
                ListaPcManager.Items.Clear();
                ListaPcManager.ResetText();
                ComboPos.Items.Clear();
                ComboPos.Enabled = false;
                //limpia lista NUC
                comboNuc.Items.Clear();
                comboNuc.Enabled = false;
                AgregarIpPcManager(ObtenerDetalleFile("ManagerPC"));
            }
        }

        private void AgregarIpPcManager(string ipmanager)
        {
            string[] ListaIp = ipmanager.Split('/');
            //se agrega cada ip al listBox
            foreach (var ip in ListaIp)
                ListaPcManager.Items.Add(ip);
        }

        private void radioBos_CheckedChanged(object sender, EventArgs e)
        {
            txtBoxBos.Text = ObtenerDetalleFile("BosPC");

            //Limpia lista POS
            ComboPos.Items.Clear();
            ComboPos.Enabled = false;

            //limpia lista NUC
            comboNuc.Items.Clear();
            comboNuc.Enabled = false;
        }

        private void BtnTodo_Click(object sender, EventArgs e)
        {
            //vaciar controles
            HabilitarControles(false);
            //Agrega lista de files
            AgregarTodoFile();
        }

        public static bool IsNumber(string s)
        {
            if (string.IsNullOrEmpty(s) || string.IsNullOrWhiteSpace(s))
                return false;
            foreach (char c in s)
                if (!char.IsDigit(c))
                    return false;
            return true;
        }

        private void DryPos_CheckedChanged(object sender, EventArgs e)
        {
            //Limpio lista combo
            ComboPos.Enabled = true;
            ComboPos.Items.Clear();

            //limpia lista NUC
            comboNuc.Items.Clear();
            comboNuc.Enabled = false;

            //Traer datos del file
            String DryPOS = ObtenerDetalleFile("DryPOS");
            //Se separa lista de las ip
            string[] Lista = DryPOS.Split('/');
            //se agrega cada ip al combo
            foreach (var ip in Lista)
                ComboPos.Items.Add(ip);
        }

        private void WetPos_CheckedChanged(object sender, EventArgs e)
        {
            //Limpio lista combo
            ComboPos.Enabled = true;
            ComboPos.Items.Clear();

            //limpia lista NUC
            comboNuc.Items.Clear();
            comboNuc.Enabled = false;

            String WetPOS = ObtenerDetalleFile("WetPOS");
            string[] Lista = WetPOS.Split('/');
            //se agrega cada ip al combo
            foreach (var ip in Lista)
                ComboPos.Items.Add(ip);
        }

        private void BtnPing_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBoxBos.Text))
            {
                return;
            }
            System.Diagnostics.ProcessStartInfo proc = new System.Diagnostics.ProcessStartInfo();
            proc.FileName = @"C:\windows\system32\cmd.exe";
            proc.Arguments = string.Format("/c ping -t {0}", txtBoxBos.Text);
            System.Diagnostics.Process.Start(proc);
        }


        private void BtnRho_Click(object sender, EventArgs e)
        {
            if (CheckIngresar.Checked)
            {
                string pathEscritorioRemoto = Environment.ExpandEnvironmentVariables("%SystemRoot%\\system32\\mstsc.exe");

                Process cmdKey = new Process();
                cmdKey.StartInfo.FileName = Environment.ExpandEnvironmentVariables("%SystemRoot%\\system32\\cmdkey.exe");
                cmdKey.StartInfo.Arguments = string.Format("/generic:TERMSRV/{0} /user:{1} /pass:{2}", "", "", "");
                cmdKey.Start();

                //iniciar proceso de escritorio Remoto
                Process escritorioRemoto = new Process();
                escritorioRemoto.StartInfo.FileName = pathEscritorioRemoto;
                //Se le pasan los argumentos a mstsc.exe                    
                escritorioRemoto.StartInfo.Arguments = string.Format("/v:{0}", "");
                escritorioRemoto.Start();

                CheckIngresar.Checked = false;
            }
        }

        private void FormConector_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void btnUsuario_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBoxBos.Text))
            { return; }

            string NroFile = txtBoxFile.Text;
            //temporal para modificar el nro del file segun su largo (si es file 5 vs file 100)
            var temp = "";
            //formatea usuario de acuerdo al largo de nro de file
            switch (NroFile.Length)
            {
                case 1:
                    temp = string.Format("CLBO000{0}\\CLAD000{0}", NroFile);
                    break;
                case 2:
                    temp = string.Format("CLBO00{0}\\CLAD00{0}", NroFile);
                    break;
                case 3:
                    temp = string.Format("CLBO0{0}\\CLAD0{0}", NroFile);
                    break;
                default: break;
            }
            try
            {
                Clipboard.Clear();
                Clipboard.SetText(temp);
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                msg += Environment.NewLine;
                msg += Environment.NewLine;
                msg += "The problem:";
                msg += Environment.NewLine;
                msg += getOpenClipboardWindowText();
                MessageBox.Show(msg);
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern IntPtr GetOpenClipboardWindow();
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern int GetWindowText(int hwnd, StringBuilder text, int count);
        private string getOpenClipboardWindowText()
        {
            IntPtr hwnd = GetOpenClipboardWindow();
            StringBuilder sb = new StringBuilder(501);
            GetWindowText(hwnd.ToInt32(), sb, 500);
            return sb.ToString();
            // example:
            // skype_plugin_core_proxy_window: 02490E80
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBoxBos.Text))
                return;
            ConectorVnc(txtBoxBos.Text);
        }

        private void txtBoxFile_TextChanged(object sender, MouseEventArgs e)
        {
            txtBoxFile.Text = string.Empty;
        }

        private void dataEstaciones_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ListaPcManager.Items.Clear();
            ListaPcManager.ResetText();
            
            //sacando el check de los radios
            WetPos.Checked = false;
            DryPos.Checked = false;
            NUC.Checked = false;
            radioStationManager.Checked = false;
            radioBos.Checked = false;
            EasyPay.Checked = false;

            DryPos.Enabled = false;
            WetPos.Enabled = false;
            EasyPay.Enabled = false;
            NUC.Enabled = false;

            ComboPos.Items.Clear();
            ComboPos.Enabled = false;
            
            comboNuc.Items.Clear();
            comboNuc.Enabled = false;

            if (!String.IsNullOrEmpty(ObtenerDetalleFile("DryPOS")) && !string.IsNullOrEmpty(txtBoxFile.Text))
                DryPos.Enabled = true;
            if (!String.IsNullOrEmpty(ObtenerDetalleFile("WetPOS")) && !string.IsNullOrEmpty(txtBoxFile.Text))
                WetPos.Enabled = true;
            if (!String.IsNullOrEmpty(ObtenerDetalleFile("EasyPay")) && !string.IsNullOrEmpty(txtBoxFile.Text))
                EasyPay.Enabled = true;
            if (!String.IsNullOrEmpty(ObtenerDetalleFile("NUC")) && !string.IsNullOrEmpty(txtBoxFile.Text))
                NUC.Enabled = true;

            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dataEstaciones.Rows[e.RowIndex];
                AgregarIpPcManager(row.Cells["ManagerPC"].Value.ToString());
                txtBoxBos.Text = row.Cells["BosPC"].Value.ToString();
                txtBoxEasyPay.Text = row.Cells["EasyPay"].Value.ToString();
            }
        }

        private void radioEasyPay_CheckedChanged(object sender, EventArgs e)
        {
            txtBoxEasyPay.Text = ObtenerDetalleFile("EasyPay");
            ComboPos.Items.Clear();
            ComboPos.Enabled = false;
            comboNuc.Items.Clear();
            comboNuc.Enabled = false;
        }

        private void NUC_Click(object sender, EventArgs e)
        {
            //Limpia Lisra
            comboNuc.Enabled = true;
            comboNuc.Items.Clear();

            //Limpio lista combo
            ComboPos.Enabled = false;
            ComboPos.Items.Clear();

            string NUC = ObtenerDetalleFile("NUC");
            string[] ListaNuc = NUC.Split('*');
            //Se agrega nuc a la lista
            foreach (var detalle in ListaNuc)
            {
                string[] detalleNuc = detalle.Split('/');
                comboNuc.Items.Add("Pc: " + detalleNuc[0] + " / " + "IP: " + detalleNuc[1] + " / " + "Nro: " + detalleNuc[2] + " / " + detalleNuc[3]);
            }
        }

        private void radioStationManager_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txtBoxFile.Text))
            {
                //Limpiar lista pc manager
                ListaPcManager.Items.Clear();
                ListaPcManager.ResetText();
                ComboPos.Items.Clear();
                ComboPos.Enabled = false;
                //limpia lista NUC
                comboNuc.Items.Clear();
                comboNuc.Enabled = false;
                AgregarIpPcManager(ObtenerDetalleFile("ManagerPC"));
            }
        }
    }
}
