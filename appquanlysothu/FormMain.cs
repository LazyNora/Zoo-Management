using ADOX;
using ClosedXML.Excel;
using DataGridViewAutoFilter;
using LiveCharts;
using LiveCharts.Helpers;
using LiveCharts.Wpf;
using Microsoft.Win32;
using MoreLinq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using Axis = LiveCharts.Wpf.Axis;
using Button = System.Windows.Forms.Button;
using Color = System.Drawing.Color;
using DataTable = System.Data.DataTable;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using SeriesCollection = LiveCharts.SeriesCollection;
using TextBox = System.Windows.Forms.TextBox;

namespace appquanlysothu
{
	public partial class FormMain : Form
	{
		// Object of Creature model
		private Creature creatureObject = new Creature();

		// List to hold creature details
		private List<Creature> creatureList = new List<Creature>();

		private int borderSize = 2;
		private Size formSize;
		private UserPreferenceChangedEventHandler UserPreferenceChanged;

		public FormMain()
		{
			InitializeComponent();
			this.Padding = new Padding(borderSize);//Border size
			this.BackColor = Color.FromArgb(98, 102, 244);//Border color
			LoadTheme();
			UserPreferenceChanged = new UserPreferenceChangedEventHandler(SystemEvents_UserPreferenceChanged);
			SystemEvents.UserPreferenceChanged += UserPreferenceChanged;
			this.Disposed += new EventHandler(Form_Disposed);
			panel2.Visible = false;
			panelData.Visible = true;
			panelCS.Visible = false;
			panelChart.Visible = false;
		}

		private void LoadTheme()
		{
			var themeColor = WinTheme.GetAccentColor();//Windows Accent Color
			var lightColor = ControlPaint.Light(themeColor);
			var darkColor = ControlPaint.Dark(themeColor);
			panelMenu.BackColor = themeColor;

			//Buttons
			foreach (Button button in this.Controls.OfType<Button>())
			{
				button.BackColor = themeColor;
			}
			foreach (Button button in this.panelData.Controls.OfType<Button>())
			{
				button.FlatAppearance.MouseOverBackColor = themeColor;
				button.FlatAppearance.MouseDownBackColor = lightColor;
			}
		}
		private void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
		{
			if (e.Category == UserPreferenceCategory.General || e.Category == UserPreferenceCategory.VisualStyle)
			{
				LoadTheme();
			}
		}
		private void Form_Disposed(object sender, EventArgs e)
		{
			SystemEvents.UserPreferenceChanged -= UserPreferenceChanged;
		}

		[DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
		private extern static void ReleaseCapture();
		[DllImport("user32.DLL", EntryPoint = "SendMessage")]
		private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
		private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
		{

		}
		private void CreateGridView()
		{
			dataGridView1.AutoGenerateColumns = false;
			DataGridViewAutoFilterTextBoxColumn IdColumn = new DataGridViewAutoFilterTextBoxColumn();
			IdColumn.Name = "IdColumn";
			IdColumn.DataPropertyName = "Id";
			IdColumn.HeaderText = "Id";
			IdColumn.HeaderCell.Style.Font = new Font(dataGridView1.Font, FontStyle.Regular);
			IdColumn.ReadOnly = false;
			IdColumn.FilteringEnabled = true;
			dataGridView1.Columns.Add(IdColumn);

			DataGridViewAutoFilterTextBoxColumn NameColumn = new DataGridViewAutoFilterTextBoxColumn();
			NameColumn.Name = "NameColumn";
			NameColumn.DataPropertyName = "Name";
			NameColumn.HeaderText = "Tên";
			NameColumn.HeaderCell.Style.Font = new Font(dataGridView1.Font, FontStyle.Regular);
			NameColumn.ReadOnly = false;
			NameColumn.FilteringEnabled = true;
			dataGridView1.Columns.Add(NameColumn);

			DataGridViewAutoFilterTextBoxColumn TypeColumn = new DataGridViewAutoFilterTextBoxColumn();
			TypeColumn.Name = "TypeColumn";
			TypeColumn.DataPropertyName = "Type";
			TypeColumn.HeaderText = "Loài";
			TypeColumn.HeaderCell.Style.Font = new Font(dataGridView1.Font, FontStyle.Regular);
			TypeColumn.ReadOnly = false;
			TypeColumn.FilteringEnabled = true;
			dataGridView1.Columns.Add(TypeColumn);

			DataGridViewAutoFilterTextBoxColumn BarnColumn = new DataGridViewAutoFilterTextBoxColumn();
			BarnColumn.Name = "BarnColumn";
			BarnColumn.DataPropertyName = "Barn";
			BarnColumn.HeaderText = "Chuồng";
			BarnColumn.HeaderCell.Style.Font = new Font(dataGridView1.Font, FontStyle.Regular);
			BarnColumn.ReadOnly = false;
			BarnColumn.FilteringEnabled = true;
			dataGridView1.Columns.Add(BarnColumn);

			DataGridViewAutoFilterTextBoxColumn AgeColumn = new DataGridViewAutoFilterTextBoxColumn();
			AgeColumn.Name = "AgeColumn";
			AgeColumn.DataPropertyName = "Age";
			AgeColumn.HeaderText = "Tuổi";
			AgeColumn.HeaderCell.Style.Font = new Font(dataGridView1.Font, FontStyle.Regular);
			AgeColumn.DefaultCellStyle.BackColor = Color.FromArgb(255, 224, 224, 224);
			AgeColumn.ReadOnly = true;
			AgeColumn.FilteringEnabled = true;
			dataGridView1.Columns.Add(AgeColumn);

			DataGridViewComboBoxColumn SexColumn = new DataGridViewComboBoxColumn();
			SexColumn.Items.AddRange(new KeyValuePair<int, string>(0, "-Select-"), new KeyValuePair<int, string>(1, "Đực"), new KeyValuePair<int, string>(2, "Cái"));
			SexColumn.Name = "SexColumn";
			SexColumn.DataPropertyName = "Sex";
			SexColumn.HeaderText = "Giới Tính";
			SexColumn.HeaderCell.Style.Font = new Font(dataGridView1.Font, FontStyle.Regular);
			SexColumn.ReadOnly = false;
			SexColumn.ValueMember = "Key";
			SexColumn.DisplayMember = "Value";
			SexColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
			dataGridView1.Columns.Add(SexColumn);

			DataGridViewComboBoxColumn ConditionColumn = new DataGridViewComboBoxColumn();
			ConditionColumn.Items.AddRange(new KeyValuePair<int, string>(0, "-Select-"), new KeyValuePair<int, string>(1, "Bình Thường"), new KeyValuePair<int, string>(2, "Bệnh"), new KeyValuePair<int, string>(3, "Có Thai"));
			ConditionColumn.Name = "ConditionColumn";
			ConditionColumn.DataPropertyName = "Condition";
			ConditionColumn.HeaderText = "Tình Trạng Sức Khỏe";
			ConditionColumn.HeaderCell.Style.Font = new Font(dataGridView1.Font, FontStyle.Regular);
			ConditionColumn.ReadOnly = false;
			ConditionColumn.DisplayMember = "Value";
			ConditionColumn.ValueMember = "Key";
			ConditionColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
			dataGridView1.Columns.Add(ConditionColumn);

			DataGridViewComboBoxColumn CarnHerbivoreColumn = new DataGridViewComboBoxColumn();
			CarnHerbivoreColumn.Items.AddRange(new KeyValuePair<int, string>(0, "-Select-"), new KeyValuePair<int, string>(1, "Ăn Cỏ"), new KeyValuePair<int, string>(2, "Ăn Thịt"), new KeyValuePair<int, string>(3, "Ăn Tạp"));
			CarnHerbivoreColumn.Name = "CarnHerbivoreColumn";
			CarnHerbivoreColumn.DataPropertyName = "Carn_herbivore";
			CarnHerbivoreColumn.HeaderText = "Chủng Loại";
			CarnHerbivoreColumn.HeaderCell.Style.Font = new Font(dataGridView1.Font, FontStyle.Regular);
			CarnHerbivoreColumn.ReadOnly = false;
			CarnHerbivoreColumn.DisplayMember = "Value";
			CarnHerbivoreColumn.ValueMember = "Key";
			CarnHerbivoreColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
			dataGridView1.Columns.Add(CarnHerbivoreColumn);

			DataGridViewAutoFilterTextBoxColumn WeightColumn = new DataGridViewAutoFilterTextBoxColumn();
			WeightColumn.Name = "WeightColumn";
			WeightColumn.DataPropertyName = "Weight";
			WeightColumn.HeaderText = "Cân Nặng";
			WeightColumn.HeaderCell.Style.Font = new Font(dataGridView1.Font, FontStyle.Regular);
			WeightColumn.ReadOnly = false;
			WeightColumn.FilteringEnabled = true;
			dataGridView1.Columns.Add(WeightColumn);

			DataGridViewAutoFilterTextBoxColumn BirthColumn = new DataGridViewAutoFilterTextBoxColumn();
			BirthColumn.Name = "BirthColumn";
			BirthColumn.DataPropertyName = "Birth";
			BirthColumn.HeaderText = "Ngày Sinh";
			BirthColumn.HeaderCell.Style.Font = new Font(dataGridView1.Font, FontStyle.Regular);
			BirthColumn.ReadOnly = false;
			BirthColumn.FilteringEnabled = true;
			dataGridView1.Columns.Add(BirthColumn);

			DataGridViewAutoFilterTextBoxColumn EntryColumn = new DataGridViewAutoFilterTextBoxColumn();
			EntryColumn.Name = "EntryColumn";
			EntryColumn.DataPropertyName = "Entry";
			EntryColumn.HeaderText = "Ngày Nhập";
			EntryColumn.HeaderCell.Style.Font = new Font(dataGridView1.Font, FontStyle.Regular);
			EntryColumn.ReadOnly = false;
			EntryColumn.FilteringEnabled = true;
			dataGridView1.Columns.Add(EntryColumn);

			DataGridViewTextBoxColumn NoteColumn = new DataGridViewTextBoxColumn();
			NoteColumn.Name = "NoteColumn";
			NoteColumn.DataPropertyName = "Note";
			NoteColumn.HeaderText = "Ghi Chú";
			NoteColumn.HeaderCell.Style.Font = new Font(dataGridView1.Font, FontStyle.Regular);
			NoteColumn.ReadOnly = false;
			dataGridView1.Columns.Add(NoteColumn);
		}
		private void Form1_Load(object sender, EventArgs e)
		{
			formSize = this.ClientSize;
			dataGridView1.Columns.Clear();
			CreateGridView();
			txtFileName.Text = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\data.xlsx";
			Import();
			DataTable dT = new DataTable();
			dT.Columns.Add("ID");
			dT.Columns.Add("Value");
			dT.Rows.Add(0, "                ---- Select ----");
			dT.Rows.Add(1, "Giới Tính");
			dT.Rows.Add(2, "Tình Trạng Sức Khỏe");
			dT.Rows.Add(3, "Chủng Loại");
			cbPie.DataSource = dT;
			cbPie.DisplayMember = "Value";
			cbPie.ValueMember = "ID";
			cbPie.SelectedIndex = 0;
			pieChart1.LegendLocation = LegendLocation.Right;
			DataTable dT2 = new DataTable();
			dT2.Columns.Add("ID");
			dT2.Columns.Add("Value");
			dT2.Rows.Add(0, "                ---- Select ----");
			dT2.Rows.Add(1, "Chuồng");
			dT2.Rows.Add(2, "Loài");
			dT2.Rows.Add(3, "Năm Nhập");
			cbBar.DataSource = dT2;
			cbBar.DisplayMember = "Value";
			cbBar.ValueMember = "ID";
			cbBar.SelectedIndex = 0;
			cartesianChart1.Visible = false;
		}

		private void ShowAllLabel_Click(object sender, EventArgs e)
		{
			DataGridViewAutoFilterTextBoxColumn.RemoveFilter(dataGridView1);
		}
		private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
		{
			string filterStatus = DataGridViewAutoFilterColumnHeaderCell.GetFilterStatus(dataGridView1);
			if (string.IsNullOrEmpty(filterStatus))
			{
				ShowAllLabel.Visible = false;
				FilterStatusLabel.Visible = false;
			}
			else
			{
				ShowAllLabel.Visible = true;
				FilterStatusLabel.Visible = true;
				FilterStatusLabel.Text = filterStatus;
			}
		}

		private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Alt && (e.KeyCode == System.Windows.Forms.Keys.Down || e.KeyCode == System.Windows.Forms.Keys.Up) && dataGridView1.CurrentCell != null && dataGridView1.CurrentCell.OwningColumn.HeaderCell is DataGridViewAutoFilterColumnHeaderCell filterCell)
			{
				filterCell.ShowDropDownList();
				e.Handled = true;
			}
		}

		private DataTable GetDataTableFromDGV(DataGridView dgv)
		{
			var dt = new DataTable();
			dt.TableName = "List";
			foreach (DataGridViewColumn column in dgv.Columns)
			{
				if (column.Visible)
				{
					dt.Columns.Add(column.Name);
				}
			}
			object[] cellValue = new object[dgv.Columns.Count];
			foreach (DataGridViewRow row in dgv.Rows)
			{
				for (int i = 0; i < row.Cells.Count; i++)
				{
					cellValue[i] = row.Cells[i].Value;
				}
				dt.Rows.Add(cellValue);
			}
			return dt;
		}
		private void GetListFromDGV(DataGridView dgv)
		{
			int count = creatureList.Count;
			foreach (DataGridViewRow dr in dgv.Rows)
			{
				if (dr.Index == count)
				{
					continue;
				}
				int sex = int.Parse(dr.Cells[5].Value == null ? "0" : dr.Cells[5].Value.ToString());
				int con = int.Parse(dr.Cells[6].Value == null ? "0" : dr.Cells[6].Value.ToString());
				int carn = int.Parse(dr.Cells[7].Value == null ? "0" : dr.Cells[7].Value.ToString());
				creatureList.Add(new Creature(dr.Cells[0].Value == null ? "" : dr.Cells[0].Value.ToString(), dr.Cells[1].Value == null ? "" : dr.Cells[1].Value.ToString(), dr.Cells[2].Value == null ? "" : dr.Cells[2].Value.ToString(), dr.Cells[3].Value == null ? "" : dr.Cells[3].Value.ToString(), sex, con, carn, float.Parse(dr.Cells[8].Value == null ? "0" : dr.Cells[8].Value.ToString()), DateTime.Parse(dr.Cells[10].Value == null ? "01/01/0001" : dr.Cells[10].Value.ToString()), DateTime.Parse(dr.Cells[9].Value == null ? "01/01/0001" : dr.Cells[9].Value.ToString()), dr.Cells[11].Value == null ? "" : dr.Cells[11].Value.ToString()));
			}
			creatureList.RemoveRange(0, count);
		}

		private void panelTitle_MouseDown(object sender, MouseEventArgs e)
		{
			ReleaseCapture();
			SendMessage(this.Handle, 0x112, 0xf012, 0);
		}
		//Overridden methods
		protected override void WndProc(ref Message m)
		{
			const int WM_NCCALCSIZE = 0x0083;//Standar Title Bar - Snap Window
			const int WM_SYSCOMMAND = 0x0112;
			const int SC_MINIMIZE = 0xF020; //Minimize form (Before)
			const int SC_RESTORE = 0xF120; //Restore form (Before)
			const int WM_NCHITTEST = 0x0084;//Win32, Mouse Input Notification: Determine what part of the window corresponds to a point, allows to resize the form.
			const int resizeAreaSize = 10;

			#region Form Resize
			// Resize/WM_NCHITTEST values
			const int HTCLIENT = 1; //Represents the client area of the window
			const int HTLEFT = 10;  //Left border of a window, allows resize horizontally to the left
			const int HTRIGHT = 11; //Right border of a window, allows resize horizontally to the right
			const int HTTOP = 12;   //Upper-horizontal border of a window, allows resize vertically up
			const int HTTOPLEFT = 13;//Upper-left corner of a window border, allows resize diagonally to the left
			const int HTTOPRIGHT = 14;//Upper-right corner of a window border, allows resize diagonally to the right
			const int HTBOTTOM = 15; //Lower-horizontal border of a window, allows resize vertically down
			const int HTBOTTOMLEFT = 16;//Lower-left corner of a window border, allows resize diagonally to the left
			const int HTBOTTOMRIGHT = 17;//Lower-right corner of a window border, allows resize diagonally to the right

			///<Doc> More Information: https://docs.microsoft.com/en-us/windows/win32/inputdev/wm-nchittest </Doc>

			if (m.Msg == WM_NCHITTEST)
			{ //If the windows m is WM_NCHITTEST
				base.WndProc(ref m);
				if (this.WindowState == FormWindowState.Normal)//Resize the form if it is in normal state
				{
					if ((int)m.Result == HTCLIENT)//If the result of the m (mouse pointer) is in the client area of the window
					{
						Point screenPoint = new Point(m.LParam.ToInt32()); //Gets screen point coordinates(X and Y coordinate of the pointer)                           
						Point clientPoint = this.PointToClient(screenPoint); //Computes the location of the screen point into client coordinates                          

						if (clientPoint.Y <= resizeAreaSize)//If the pointer is at the top of the form (within the resize area- X coordinate)
						{
							if (clientPoint.X <= resizeAreaSize) //If the pointer is at the coordinate X=0 or less than the resizing area(X=10) in 
								m.Result = (IntPtr)HTTOPLEFT; //Resize diagonally to the left
							else if (clientPoint.X < (this.Size.Width - resizeAreaSize))//If the pointer is at the coordinate X=11 or less than the width of the form(X=Form.Width-resizeArea)
								m.Result = (IntPtr)HTTOP; //Resize vertically up
							else //Resize diagonally to the right
								m.Result = (IntPtr)HTTOPRIGHT;
						}
						else if (clientPoint.Y <= (this.Size.Height - resizeAreaSize)) //If the pointer is inside the form at the Y coordinate(discounting the resize area size)
						{
							if (clientPoint.X <= resizeAreaSize)//Resize horizontally to the left
								m.Result = (IntPtr)HTLEFT;
							else if (clientPoint.X > (this.Width - resizeAreaSize))//Resize horizontally to the right
								m.Result = (IntPtr)HTRIGHT;
						}
						else
						{
							if (clientPoint.X <= resizeAreaSize)//Resize diagonally to the left
								m.Result = (IntPtr)HTBOTTOMLEFT;
							else if (clientPoint.X < (this.Size.Width - resizeAreaSize)) //Resize vertically down
								m.Result = (IntPtr)HTBOTTOM;
							else //Resize diagonally to the right
								m.Result = (IntPtr)HTBOTTOMRIGHT;
						}
					}
				}
				return;
			}
			#endregion

			//Remove border and keep snap window
			if (m.Msg == WM_NCCALCSIZE && m.WParam.ToInt32() == 1)
			{
				return;
			}

			//Keep form size when it is minimized and restored. Since the form is resized because it takes into account the size of the title bar and borders.
			if (m.Msg == WM_SYSCOMMAND)
			{
				int wParam = (m.WParam.ToInt32() & 0xFFF0);

				if (wParam == SC_MINIMIZE)  //Before
					formSize = this.ClientSize;
				if (wParam == SC_RESTORE)// Restored form(Before)
					this.Size = formSize;
			}
			base.WndProc(ref m);
		}

		private void Form1_Resize(object sender, EventArgs e)
		{
			AdjustForm();
		}
		private void AdjustForm()
		{
			switch (this.WindowState)
			{
				case FormWindowState.Maximized: //Maximized form (After)
					this.Padding = new Padding(8, 8, 8, 0);
					break;
				case FormWindowState.Normal: //Restored form (After)
					if (this.Padding.Top != borderSize)
						this.Padding = new Padding(borderSize);
					break;
			}
		}
		private void CollapseMenu()
		{
			if (this.panelMenu.Width > 200) //Collapse menu
			{
				panelMenu.Width = 100;
				pictureBox1.Visible = false;
				label1.Visible = false;
				btnMenu.Dock = DockStyle.Top;
				foreach (Button menuButton in panelMenu.Controls.OfType<Button>())
				{
					menuButton.Text = "";
					menuButton.ImageAlign = ContentAlignment.MiddleCenter;
					menuButton.Padding = new Padding(0);
				}
			}
			else
			{ //Expand menu
				panelMenu.Width = 230;
				pictureBox1.Visible = true;
				label1.Visible = true;
				btnMenu.Dock = DockStyle.None;
				foreach (Button menuButton in panelMenu.Controls.OfType<Button>())
				{
					menuButton.Text = "   " + menuButton.Tag.ToString();
					menuButton.ImageAlign = ContentAlignment.MiddleLeft;
					menuButton.Padding = new Padding(10, 0, 0, 0);
				}
			}
		}

		private void btnMinimize_Click(object sender, EventArgs e)
		{
			formSize = this.ClientSize;
			this.WindowState = FormWindowState.Minimized;
		}

		private void btnMaximize_Click(object sender, EventArgs e)
		{
			if (this.WindowState == FormWindowState.Normal)
			{
				formSize = this.ClientSize;
				this.WindowState = FormWindowState.Maximized;
			}
			else
			{
				this.WindowState = FormWindowState.Normal;
				this.Size = formSize;
			}
		}

		private void btnClose_Click(object sender, EventArgs e)
		{
			Application.Exit();
		}

		private void btnMenu_Click(object sender, EventArgs e)
		{
			CollapseMenu();
		}

		private void label1_Click(object sender, EventArgs e)
		{
			CollapseMenu();
		}

		private void Info_Click(object sender, EventArgs e)
		{
			panel2.Visible = true;
			panelData.Visible = false;
			panelCS.Visible = false;
			panelChart.Visible = false;
		}
		private void DataBase_Click(object sender, EventArgs e)
		{
			panel2.Visible = false;
			panelCS.Visible = false;
			panelData.Visible = true;
			panelChart.Visible = false;
		}

		private void CommingSoon_Click(object sender, EventArgs e)
		{
			panelCS.Visible = true;
			panel2.Visible = false;
			panelData.Visible = false;
			panelChart.Visible = false;
		}
		public static string SelectedTable = string.Empty;
		private void browseBtn_Click(object sender, EventArgs e)
		{
			System.Windows.Forms.OpenFileDialog fdlg = new System.Windows.Forms.OpenFileDialog();
			fdlg.Title = "Select file";
			fdlg.InitialDirectory = @"c:\";
			fdlg.FileName = txtFileName.Text;
			fdlg.Filter = "Excel |*.xlsx";    //"Excel Sheet(*.xls)|*.xls|All Files(*.*)|*.*"
			fdlg.FilterIndex = 1;
			fdlg.RestoreDirectory = true;
			if (fdlg.ShowDialog() == DialogResult.OK)
			{
				dataGridView1.Columns.Clear();
				txtFileName.Text = fdlg.FileName;
				CreateGridView();
				Import();
				Application.DoEvents();
			}
		}

		private void Import()
		{
			if (txtFileName.Text.Trim() != string.Empty)
			{
				try
				{
					string[] strTables = GetTableExcel(txtFileName.Text);
					if (strTables[0] != null && strTables[0] != string.Empty && items < 2)
					{
						creatureList.Clear();
						DataTable dt = GetDataTableExcel(txtFileName.Text, strTables[0]);
						foreach (DataRow dr in dt.Rows)
						{
							int sex = 0, con = 0, carn = 0;
							int.TryParse(dr[5].ToString(), out sex);
							int.TryParse(dr[6].ToString(), out con);
							int.TryParse(dr[7].ToString(), out carn);
							creatureList.Add(new Creature(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), sex, con, carn, float.Parse(dr[8].ToString()), DateTime.Parse(dr[9].ToString()), DateTime.Parse(dr[10].ToString()), dr[11].ToString()));
						}
						creatureBindingSource.DataSource = creatureList.ToDataTable();
						dataGridView1.DataSource = creatureBindingSource;
					}
					else
					{
						frmSelectTables objSelectTable = new frmSelectTables(strTables);
						objSelectTable.ShowDialog(this);
						objSelectTable.Dispose();
						if ((SelectedTable != string.Empty) && (SelectedTable != null))
						{
							creatureList.Clear();
							DataTable dt = GetDataTableExcel(txtFileName.Text, SelectedTable);
							foreach (DataRow dr in dt.Rows)
							{
								int sex = 0, con = 0, carn = 0;
								int.TryParse(dr[5].ToString(), out sex);
								int.TryParse(dr[6].ToString(), out con);
								int.TryParse(dr[7].ToString(), out carn);
								creatureList.Add(new Creature(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), sex, con, carn, float.Parse(dr[8].ToString()), DateTime.Parse(dr[9].ToString()), DateTime.Parse(dr[10].ToString()), dr[11].ToString()));
							}
							creatureBindingSource.DataSource = creatureList.ToDataTable();
							dataGridView1.DataSource = creatureBindingSource;
						}
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message.ToString());
				}
			}
		}

		public static DataTable GetDataTableExcel(string strFileName, string Table)
		{
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + strFileName + "; Extended Properties = \"Excel 12.0;HDR=Yes;IMEX=1\";");
			conn.Open();
			string strQuery = "SELECT * FROM [" + Table + "]";
			System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter(strQuery, conn);
			System.Data.DataSet ds = new System.Data.DataSet();
			adapter.Fill(ds);
			return ds.Tables[0];
		}
		private static int items = 0;
		public static string[] GetTableExcel(string strFileName)
		{
			string[] strTables = new string[100];
			Catalog oCatlog = new Catalog();
			ADOX.Table oTable = new ADOX.Table();
			ADODB.Connection oConn = new ADODB.Connection();
			oConn.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + strFileName + "; Extended Properties = \"Excel 12.0;HDR=Yes;IMEX=1\";", "", "", 0);
			oCatlog.ActiveConnection = oConn;
			if (oCatlog.Tables.Count > 0)
			{
				int item = 0;
				foreach (ADOX.Table tab in oCatlog.Tables)
				{
					if (tab.Type == "TABLE")
					{
						strTables[item] = tab.Name;
						item++;
					}
				}
				items = item;
			}
			return strTables;
		}

		private void txtFileName_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == System.Windows.Forms.Keys.Enter && txtFileName.Text != "")
			{
				dataGridView1.Columns.Clear();
				CreateGridView();
				Import();
				Application.DoEvents();
			}
		}

		private void saveBtn_Click(object sender, EventArgs e)
		{
			using (System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog() { Filter = "Excel Workbook|*.xlsx" })
			{
				if (sfd.ShowDialog() == DialogResult.OK)
				{
					try
					{
						using (XLWorkbook workbook = new XLWorkbook())
						{
							workbook.Worksheets.Add(GetDataTableFromDGV(dataGridView1), "DataTable");
							workbook.SaveAs(sfd.FileName);
						}
						MessageBox.Show("Exported", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					catch (Exception ex)
					{
						MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
		}

		private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			DisplayTextBox(e.RowIndex, e.ColumnIndex);
			RI = e.RowIndex;
			CI = e.ColumnIndex;
			if (dataGridView1.CurrentRow.IsNewRow && e.ColumnIndex != 4)
			{
				creatureList.Add(new Creature());
				creatureBindingSource.DataSource = creatureList.ToDataTable();
				dataGridView1.DataSource = creatureBindingSource;
			}
		}

		private void DisplayTextBox(int rowindex, int columnindex)
		{
			if (rowindex >= 0)
			{
				//gets a collection that contains all the rows
				DataGridViewRow row = this.dataGridView1.Rows[rowindex];
				//populate the textbox from specific value of the coordinates of column and row.
				tbID.Text = row.Cells[0].Value.ToString() == string.Empty ? "" : row.Cells[0].Value.ToString();
				tbName.Text = row.Cells[1].Value.ToString() == string.Empty ? "" : row.Cells[1].Value.ToString();
				tbType.Text = row.Cells[2].Value.ToString() == string.Empty ? "" : row.Cells[2].Value.ToString();
				tbBarn.Text = row.Cells[3].Value.ToString() == string.Empty ? "" : row.Cells[3].Value.ToString();
				lbAge.Text = row.Cells[4].Value.ToString() == string.Empty ? "" : row.Cells[4].Value.ToString();
				DataTable dT = new DataTable();
				dT.Columns.Add("ID");
				dT.Columns.Add("Value");
				dT.Rows.Add(0, "-Select-");
				dT.Rows.Add(1, "Đực");
				dT.Rows.Add(2, "Cái");
				cbSex.DataSource = dT;
				cbSex.DisplayMember = "Value";
				cbSex.ValueMember = "ID";
				cbSex.SelectedIndex = int.Parse(row.Cells[5].Value.ToString() == string.Empty ? "0" : row.Cells[5].Value.ToString());
				DataTable dT2 = new DataTable();
				dT2.Columns.Add("ID");
				dT2.Columns.Add("Value");
				dT2.Rows.Add(0, "-Select-");
				dT2.Rows.Add(1, "Bình Thường");
				dT2.Rows.Add(2, "Bệnh");
				dT2.Rows.Add(3, "Có Thai");
				cbCon.DataSource = dT2;
				cbCon.DisplayMember = "Value";
				cbCon.ValueMember = "ID";
				cbCon.SelectedIndex = int.Parse(row.Cells[6].Value.ToString() == string.Empty ? "0" : row.Cells[6].Value.ToString());
				DataTable dT3 = new DataTable();
				dT3.Columns.Add("ID");
				dT3.Columns.Add("Value");
				dT3.Rows.Add(0, "-Select-");
				dT3.Rows.Add(1, "Ăn Cỏ");
				dT3.Rows.Add(2, "Ăn Thịt");
				dT3.Rows.Add(3, "Ăn Tạp");
				cbCarn.DataSource = dT3;
				cbCarn.DisplayMember = "Value";
				cbCarn.ValueMember = "ID";
				cbCarn.SelectedIndex = int.Parse(row.Cells[7].Value.ToString() == string.Empty ? "0" : row.Cells[7].Value.ToString());
				tbWeight.Text = row.Cells[8].Value.ToString() == string.Empty ? "" : row.Cells[8].Value.ToString();
				tbBirth.Text = row.Cells[9].Value.ToString() == string.Empty ? "" : DateTime.Parse(row.Cells[9].Value.ToString()).ToString("dd/MM/yyyy");
				tbEntry.Text = row.Cells[10].Value.ToString() == string.Empty ? "" : DateTime.Parse(row.Cells[10].Value.ToString()).ToString("dd/MM/yyyy");
				tbNote.Text = row.Cells[11].Value.ToString() == string.Empty ? "" : row.Cells[11].Value.ToString();

			}
		}
		private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
		{
			GetListFromDGV(dataGridView1);
			dataGridView1.Columns.Clear();
			CreateGridView();
			creatureBindingSource.DataSource = creatureList.ToDataTable();
			dataGridView1.DataSource = creatureBindingSource;
			DisplayTextBox(e.RowIndex, e.ColumnIndex);
		}

		private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
		{
			var result = MessageBox.Show("Are you sure you want to delete this row?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
			if (result == DialogResult.No)
			{
				e.Cancel = true;
			}
		}
		private void dataGridView1_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
		{
			GetListFromDGV(dataGridView1);
			dataGridView1.Columns.Clear();
			CreateGridView();
			creatureList.RemoveAt(RI);
			creatureBindingSource.DataSource = creatureList.ToDataTable();
			dataGridView1.DataSource = creatureBindingSource;
		}

		private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
		{
			if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
			{
				if (e.ColumnIndex == 5 || e.ColumnIndex == 6 || e.ColumnIndex == 7)
				{
					e.Paint(e.CellBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.Border);
					using (Pen p = new Pen(Color.WhiteSmoke, 1))
					{
						System.Drawing.Rectangle rect = e.CellBounds;
						rect.Width -= 2;
						rect.Height -= 2;
						e.Graphics.DrawRectangle(p, rect);
					}
					e.Handled = true;
				}
			}
		}
		private int RI = -1;
		private int CI = -1;
		private void tbID_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == System.Windows.Forms.Keys.Enter && tbID.Text != "" && RI != -1 && CI != -1)
			{
				dataGridView1.Rows[RI].Cells[0].Value = tbID.Text;
				dataGridView1.Rows[RI].Cells[1].Value = tbName.Text;
				dataGridView1.Rows[RI].Cells[11].Value = tbNote.Text;
				if (tbBarn.Text != "")
					dataGridView1.Rows[RI].Cells[3].Value = tbBarn.Text;
				if (tbType.Text != "")
					dataGridView1.Rows[RI].Cells[2].Value = tbType.Text;
				if (tbWeight.Text != "" && float.TryParse(tbWeight.Text, out _))
					dataGridView1.Rows[RI].Cells[8].Value = tbWeight.Text;
				if (tbEntry.Text != "" && DateTime.TryParse(tbEntry.Text, out _))
					dataGridView1.Rows[RI].Cells[10].Value = tbEntry.Text;
				if (tbBirth.Text != "" && DateTime.TryParse(tbBirth.Text, out _))
					dataGridView1.Rows[RI].Cells[19].Value = tbBirth.Text;
			}
		}

		private void tbName_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == System.Windows.Forms.Keys.Enter && RI != -1 && CI != -1)
			{
				dataGridView1.Rows[RI].Cells[1].Value = tbName.Text;
				dataGridView1.Rows[RI].Cells[0].Value = tbID.Text;
				dataGridView1.Rows[RI].Cells[11].Value = tbNote.Text;
				if (tbBarn.Text != "")
					dataGridView1.Rows[RI].Cells[3].Value = tbBarn.Text;
				if (tbType.Text != "")
					dataGridView1.Rows[RI].Cells[2].Value = tbType.Text;
				if (tbWeight.Text != "" && float.TryParse(tbWeight.Text, out _))
					dataGridView1.Rows[RI].Cells[8].Value = tbWeight.Text;
				if (tbEntry.Text != "" && DateTime.TryParse(tbEntry.Text, out _))
					dataGridView1.Rows[RI].Cells[10].Value = tbEntry.Text;
				if (tbBirth.Text != "" && DateTime.TryParse(tbBirth.Text, out _))
					dataGridView1.Rows[RI].Cells[19].Value = tbBirth.Text;
			}
		}

		private void tbBarn_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == System.Windows.Forms.Keys.Enter && tbBarn.Text != "" && RI != -1 && CI != -1)
			{
				dataGridView1.Rows[RI].Cells[3].Value = tbBarn.Text;
				dataGridView1.Rows[RI].Cells[0].Value = tbID.Text;
				dataGridView1.Rows[RI].Cells[1].Value = tbName.Text;
				dataGridView1.Rows[RI].Cells[11].Value = tbNote.Text;
				if (tbType.Text != "")
					dataGridView1.Rows[RI].Cells[2].Value = tbType.Text;
				if (tbWeight.Text != "" && float.TryParse(tbWeight.Text, out _))
					dataGridView1.Rows[RI].Cells[8].Value = tbWeight.Text;
				if (tbEntry.Text != "" && DateTime.TryParse(tbEntry.Text, out _))
					dataGridView1.Rows[RI].Cells[10].Value = tbEntry.Text;
				if (tbBirth.Text != "" && DateTime.TryParse(tbBirth.Text, out _))
					dataGridView1.Rows[RI].Cells[19].Value = tbBirth.Text;
			}
		}

		private void tbType_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == System.Windows.Forms.Keys.Enter && tbType.Text != "" && RI != -1 && CI != -1)
			{
				dataGridView1.Rows[RI].Cells[2].Value = tbType.Text;
				dataGridView1.Rows[RI].Cells[0].Value = tbID.Text;
				dataGridView1.Rows[RI].Cells[1].Value = tbName.Text;
				dataGridView1.Rows[RI].Cells[11].Value = tbNote.Text;
				if (tbBarn.Text != "")
					dataGridView1.Rows[RI].Cells[3].Value = tbBarn.Text;
				if (tbWeight.Text != "" && float.TryParse(tbWeight.Text, out _))
					dataGridView1.Rows[RI].Cells[8].Value = tbWeight.Text;
				if (tbEntry.Text != "" && DateTime.TryParse(tbEntry.Text, out _))
					dataGridView1.Rows[RI].Cells[10].Value = tbEntry.Text;
				if (tbBirth.Text != "" && DateTime.TryParse(tbBirth.Text, out _))
					dataGridView1.Rows[RI].Cells[19].Value = tbBirth.Text;
			}
		}

		private void cbSex_SelectionChangeCommitted(object sender, EventArgs e)
		{
			if (cbSex.SelectedIndex != 0)
			{
				dataGridView1.Rows[RI].Cells[5].Value = cbSex.SelectedIndex;
				dataGridView1.Rows[RI].Cells[0].Value = tbID.Text;
				dataGridView1.Rows[RI].Cells[1].Value = tbName.Text;
				dataGridView1.Rows[RI].Cells[11].Value = tbNote.Text;
				if (tbBarn.Text != "")
					dataGridView1.Rows[RI].Cells[3].Value = tbBarn.Text;
				if (tbType.Text != "")
					dataGridView1.Rows[RI].Cells[2].Value = tbType.Text;
				if (tbWeight.Text != "" && float.TryParse(tbWeight.Text, out _))
					dataGridView1.Rows[RI].Cells[8].Value = tbWeight.Text;
				if (tbEntry.Text != "" && DateTime.TryParse(tbEntry.Text, out _))
					dataGridView1.Rows[RI].Cells[10].Value = tbEntry.Text;
				if (tbBirth.Text != "" && DateTime.TryParse(tbBirth.Text, out _))
					dataGridView1.Rows[RI].Cells[19].Value = tbBirth.Text;
			}
		}

		private void tbWeight_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == System.Windows.Forms.Keys.Enter && tbWeight.Text != "" && RI != -1 && CI != -1 && float.TryParse(tbWeight.Text, out _))
			{
				dataGridView1.Rows[RI].Cells[8].Value = tbWeight.Text;
				dataGridView1.Rows[RI].Cells[0].Value = tbID.Text;
				dataGridView1.Rows[RI].Cells[1].Value = tbName.Text;
				dataGridView1.Rows[RI].Cells[11].Value = tbNote.Text;
				if (tbBarn.Text != "")
					dataGridView1.Rows[RI].Cells[3].Value = tbBarn.Text;
				if (tbType.Text != "")
					dataGridView1.Rows[RI].Cells[2].Value = tbType.Text;
				if (tbEntry.Text != "" && DateTime.TryParse(tbEntry.Text, out _))
					dataGridView1.Rows[RI].Cells[10].Value = tbEntry.Text;
				if (tbBirth.Text != "" && DateTime.TryParse(tbBirth.Text, out _))
					dataGridView1.Rows[RI].Cells[19].Value = tbBirth.Text;
			}
		}

		private void cbCon_SelectionChangeCommitted(object sender, EventArgs e)
		{
			if (cbCon.SelectedIndex != 0)
			{
				dataGridView1.Rows[RI].Cells[6].Value = cbCon.SelectedIndex;
				dataGridView1.Rows[RI].Cells[0].Value = tbID.Text;
				dataGridView1.Rows[RI].Cells[1].Value = tbName.Text;
				dataGridView1.Rows[RI].Cells[11].Value = tbNote.Text;
				if (tbBarn.Text != "")
					dataGridView1.Rows[RI].Cells[3].Value = tbBarn.Text;
				if (tbType.Text != "")
					dataGridView1.Rows[RI].Cells[2].Value = tbType.Text;
				if (tbWeight.Text != "" && float.TryParse(tbWeight.Text, out _))
					dataGridView1.Rows[RI].Cells[8].Value = tbWeight.Text;
				if (tbEntry.Text != "" && DateTime.TryParse(tbEntry.Text, out _))
					dataGridView1.Rows[RI].Cells[10].Value = tbEntry.Text;
				if (tbBirth.Text != "" && DateTime.TryParse(tbBirth.Text, out _))
					dataGridView1.Rows[RI].Cells[19].Value = tbBirth.Text;
			}
		}

		private void cbCarn_SelectionChangeCommitted(object sender, EventArgs e)
		{
			if (cbCarn.SelectedIndex != 0)
			{
				dataGridView1.Rows[RI].Cells[7].Value = cbCarn.SelectedIndex;
				dataGridView1.Rows[RI].Cells[0].Value = tbID.Text;
				dataGridView1.Rows[RI].Cells[1].Value = tbName.Text;
				dataGridView1.Rows[RI].Cells[11].Value = tbNote.Text;
				if (tbBarn.Text != "")
					dataGridView1.Rows[RI].Cells[3].Value = tbBarn.Text;
				if (tbType.Text != "")
					dataGridView1.Rows[RI].Cells[2].Value = tbType.Text;
				if (tbWeight.Text != "" && float.TryParse(tbWeight.Text, out _))
					dataGridView1.Rows[RI].Cells[8].Value = tbWeight.Text;
				if (tbEntry.Text != "" && DateTime.TryParse(tbEntry.Text, out _))
					dataGridView1.Rows[RI].Cells[10].Value = tbEntry.Text;
				if (tbBirth.Text != "" && DateTime.TryParse(tbBirth.Text, out _))
					dataGridView1.Rows[RI].Cells[19].Value = tbBirth.Text;
			}
		}

		private void tbEntry_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == System.Windows.Forms.Keys.Enter && tbEntry.Text != "" && RI != -1 && CI != -1 && DateTime.TryParse(tbEntry.Text, out _))
			{
				dataGridView1.Rows[RI].Cells[10].Value = tbEntry.Text;
				dataGridView1.Rows[RI].Cells[0].Value = tbID.Text;
				dataGridView1.Rows[RI].Cells[1].Value = tbName.Text;
				dataGridView1.Rows[RI].Cells[11].Value = tbNote.Text;
				if (tbBarn.Text != "")
					dataGridView1.Rows[RI].Cells[3].Value = tbBarn.Text;
				if (tbType.Text != "")
					dataGridView1.Rows[RI].Cells[2].Value = tbType.Text;
				if (tbWeight.Text != "" && float.TryParse(tbWeight.Text, out _))
					dataGridView1.Rows[RI].Cells[8].Value = tbWeight.Text;
				if (tbBirth.Text != "" && DateTime.TryParse(tbBirth.Text, out _))
					dataGridView1.Rows[RI].Cells[19].Value = tbBirth.Text;
			}
		}

		private void tbBirth_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == System.Windows.Forms.Keys.Enter && tbBirth.Text != "" && RI != -1 && CI != -1 && DateTime.TryParse(tbBirth.Text, out _))
			{
				dataGridView1.Rows[RI].Cells[9].Value = tbBirth.Text;
				dataGridView1.Rows[RI].Cells[0].Value = tbID.Text;
				dataGridView1.Rows[RI].Cells[1].Value = tbName.Text;
				dataGridView1.Rows[RI].Cells[11].Value = tbNote.Text;
				if (tbBarn.Text != "")
					dataGridView1.Rows[RI].Cells[3].Value = tbBarn.Text;
				if (tbType.Text != "")
					dataGridView1.Rows[RI].Cells[2].Value = tbType.Text;
				if (tbWeight.Text != "" && float.TryParse(tbWeight.Text, out _))
					dataGridView1.Rows[RI].Cells[8].Value = tbWeight.Text;
				if (tbEntry.Text != "" && DateTime.TryParse(tbEntry.Text, out _))
					dataGridView1.Rows[RI].Cells[10].Value = tbEntry.Text;
			}
		}

		private void tbNote_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == System.Windows.Forms.Keys.Enter && RI != -1 && CI != -1)
			{
				dataGridView1.Rows[RI].Cells[11].Value = tbNote.Text;
				dataGridView1.Rows[RI].Cells[0].Value = tbID.Text;
				dataGridView1.Rows[RI].Cells[1].Value = tbName.Text;
				if (tbBarn.Text != "")
					dataGridView1.Rows[RI].Cells[3].Value = tbBarn.Text;
				if (tbType.Text != "")
					dataGridView1.Rows[RI].Cells[2].Value = tbType.Text;
				if (tbWeight.Text != "" && float.TryParse(tbWeight.Text, out _))
					dataGridView1.Rows[RI].Cells[8].Value = tbWeight.Text;
				if (tbEntry.Text != "" && DateTime.TryParse(tbEntry.Text, out _))
					dataGridView1.Rows[RI].Cells[10].Value = tbEntry.Text;
				if (tbBirth.Text != "" && DateTime.TryParse(tbBirth.Text, out _))
					dataGridView1.Rows[RI].Cells[19].Value = tbBirth.Text;
			}
		}

		Func<ChartPoint, string> lablePoint = chartpoint => string.Format("{0} ({1:P})", chartpoint.Y, chartpoint.Participation);
		Func<ChartPoint, string> lablePoint2 = chartpoint => string.Format("{0}", chartpoint.Y);

		private void cbPie_SelectionChangeCommitted(object sender, EventArgs e)
		{
			SeriesCollection series = new SeriesCollection();
			int csex1 = 0, csex2 = 0, ccon1 = 0, ccon2 = 0, ccon3 = 0, ccarn1 = 0, ccarn2 = 0, ccarn3 = 0;
			foreach (var obj in creatureList)
			{
				if (obj.Sex == 1) csex1++;
				if (obj.Sex == 2) csex2++;
				if (obj.Condition == 1) ccon1++;
				if (obj.Condition == 2) ccon2++;
				if (obj.Condition == 3) ccon3++;
				if (obj.Carn_herbivore == 1) ccarn1++;
				if (obj.Carn_herbivore == 2) ccarn2++;
				if (obj.Carn_herbivore == 3) ccarn3++;
			}
			if (cbPie.SelectedIndex == 1)
			{
				series.Clear();
				series.Add(new PieSeries() { Title = "Đực", Values = new ChartValues<int> { csex1 }, DataLabels = true, LabelPoint = lablePoint, Fill = System.Windows.Media.Brushes.DarkRed });
				series.Add(new PieSeries() { Title = "Cái", Values = new ChartValues<int> { csex2 }, DataLabels = true, LabelPoint = lablePoint, Fill = System.Windows.Media.Brushes.Bisque, Foreground = System.Windows.Media.Brushes.Black });
				pieChart1.InnerRadius = 0;
				pieChart1.Series = series;
			}
			else if (cbPie.SelectedIndex == 2)
			{
				series.Clear();
				series.Add(new PieSeries() { Title = "Bình Thường", Values = new ChartValues<int> { ccon1 }, DataLabels = true, LabelPoint = lablePoint, Fill = System.Windows.Media.Brushes.LightSeaGreen, Foreground = System.Windows.Media.Brushes.Black });
				series.Add(new PieSeries() { Title = "Bệnh", Values = new ChartValues<int> { ccon2 }, DataLabels = true, LabelPoint = lablePoint, Fill = System.Windows.Media.Brushes.Gold, Foreground = System.Windows.Media.Brushes.Black });
				series.Add(new PieSeries() { Title = "Có Thai", Values = new ChartValues<int> { ccon3 }, DataLabels = true, LabelPoint = lablePoint, Fill = System.Windows.Media.Brushes.Plum, Foreground = System.Windows.Media.Brushes.Black });
				pieChart1.InnerRadius = 100;
				pieChart1.Series = series;
			}
			else if (cbPie.SelectedIndex == 3)
			{
				series.Clear();
				series.Add(new PieSeries() { Title = "Ăn Cỏ", Values = new ChartValues<int> { ccon1 }, DataLabels = true, LabelPoint = lablePoint, Fill = System.Windows.Media.Brushes.MediumAquamarine, Foreground = System.Windows.Media.Brushes.Black });
				series.Add(new PieSeries() { Title = "Ăn Thịt", Values = new ChartValues<int> { ccon2 }, DataLabels = true, LabelPoint = lablePoint, Fill = System.Windows.Media.Brushes.Tomato, Foreground = System.Windows.Media.Brushes.Black });
				series.Add(new PieSeries() { Title = "Ăn Tạp", Values = new ChartValues<int> { ccon3 }, DataLabels = true, LabelPoint = lablePoint, Fill = System.Windows.Media.Brushes.LemonChiffon, Foreground = System.Windows.Media.Brushes.Black });
				pieChart1.InnerRadius = 0;
				pieChart1.Series = series;
			}
		}

		private void Chart_Click(object sender, EventArgs e)
		{
			panel2.Visible = false;
			panelData.Visible = false;
			panelCS.Visible = false;
			panelChart.Visible = true;
		}
		public List<string> strings = new List<string>();
		public List<int> counts = new List<int>();
		private void cbBar_SelectionChangeCommitted(object sender, EventArgs e)
		{
			SeriesCollection series = new SeriesCollection();
			cartesianChart1.Visible = true;
			if (cbBar.SelectedIndex == 1)
			{
				series.Clear();
				strings.Clear();
				counts.Clear();
				foreach (var obj in creatureList)
				{
					if (strings.Contains(obj.Barn) == false)
					{
						int count = 0;
						strings.Add(obj.Barn);
						foreach (var obj1 in creatureList)
						{
							if (obj1.Barn == obj.Barn)
							{
								count++;
							}
						}
						counts.Add(count);
					}
				}
				Lables = strings;
				Results = counts.AsChartValues();
				series.Add(new ColumnSeries() { Title = "Tổng: ", Values = Results, DataLabels = true, LabelPoint = lablePoint2, Fill = System.Windows.Media.Brushes.Khaki });
				cartesianChart1.AxisX.Clear();
				cartesianChart1.AxisY.Clear();
				cartesianChart1.AxisX.Add(new Axis { Title = "Chuồng", Labels = Lables, DisableAnimations = true, Separator = new Separator { Step = 1 }, Foreground = System.Windows.Media.Brushes.Black });
				cartesianChart1.AxisY.Add(new Axis { LabelFormatter = null, Foreground = System.Windows.Media.Brushes.Black });
				cartesianChart1.Series = series;
			}
			else if (cbBar.SelectedIndex == 2)
			{
				series.Clear();
				strings.Clear();
				counts.Clear();
				foreach (var obj in creatureList)
				{
					if (strings.Contains(obj.Type) == false)
					{
						int count = 0;
						strings.Add(obj.Type);
						foreach (var obj1 in creatureList)
						{
							if (obj1.Type == obj.Type)
							{
								count++;
							}
						}
						counts.Add(count);
					}
				}
				Lables = strings;
				Results = counts.AsChartValues();
				series.Add(new ColumnSeries() { Title = "Tổng: ", Values = Results, DataLabels = true, LabelPoint = lablePoint2, Fill = System.Windows.Media.Brushes.LightSeaGreen });
				cartesianChart1.AxisX.Clear();
				cartesianChart1.AxisY.Clear();
				cartesianChart1.AxisX.Add(new Axis { Title = "Loài", LabelsRotation = -20, DisableAnimations = true, Labels = Lables, Foreground = System.Windows.Media.Brushes.Black, Separator = new Separator { Step = 1 } });
				cartesianChart1.AxisY.Add(new Axis { LabelFormatter = null, Foreground = System.Windows.Media.Brushes.Black });
				cartesianChart1.Series = series;
			}
			else if (cbBar.SelectedIndex == 3)
			{
				series.Clear();
				strings.Clear();
				counts.Clear();
				foreach (var obj in creatureList)
				{
					if (strings.Contains(obj.Entry.Year.ToString()) == false)
					{
						int count = 0;
						strings.Add(obj.Entry.Year.ToString());
						foreach (var obj1 in creatureList)
						{
							if (obj1.Entry.Year.ToString() == obj.Entry.Year.ToString())
							{
								count++;
							}
						}
						counts.Add(count);
					}
				}
				Lables = strings;
				Results = counts.AsChartValues();
				series.Add(new ColumnSeries() { Title = "Tổng: ", Values = Results, DataLabels = true, LabelPoint = lablePoint2, Fill = System.Windows.Media.Brushes.PaleVioletRed });
				cartesianChart1.AxisX.Clear();
				cartesianChart1.AxisY.Clear();
				cartesianChart1.AxisX.Add(new Axis { Title = "Năm", LabelsRotation = -20, DisableAnimations = true, Separator = new Separator { Step = 1 }, Labels = Lables, Foreground = System.Windows.Media.Brushes.Black });
				cartesianChart1.AxisY.Add(new Axis { LabelFormatter = null, Foreground = System.Windows.Media.Brushes.Black });
				cartesianChart1.Series = series;
			}
		}
		public ChartValues<int> Results = new ChartValues<int>();
		public List<string> Lables = new List<string>();
		private void tbSearch_TextChanged(object sender, EventArgs e)
		{
			if (cbBar.SelectedIndex == 3)
			{
				strings.Clear();
				counts.Clear();
				foreach (var obj in creatureList)
				{
					if (strings.Contains(obj.Entry.Year.ToString()) == false)
					{
						int count = 0;
						strings.Add(obj.Entry.Year.ToString());
						foreach (var obj1 in creatureList)
						{
							if (obj1.Entry.Year.ToString() == obj.Entry.Year.ToString())
							{
								count++;
							}
						}
						counts.Add(count);
					}
				}
			}
			else if (cbBar.SelectedIndex == 2)
			{
				strings.Clear();
				counts.Clear();
				foreach (var obj in creatureList)
				{
					if (strings.Contains(obj.Type) == false)
					{
						int count = 0;
						strings.Add(obj.Type);
						foreach (var obj1 in creatureList)
						{
							if (obj1.Type == obj.Type)
							{
								count++;
							}
						}
						counts.Add(count);
					}
				}
			}
			else if (cbBar.SelectedIndex == 1)
			{
				strings.Clear();
				counts.Clear();
				foreach (var obj in creatureList)
				{
					if (strings.Contains(obj.Barn) == false)
					{
						int count = 0;
						strings.Add(obj.Barn);
						foreach (var obj1 in creatureList)
						{
							if (obj1.Barn == obj.Barn)
							{
								count++;
							}
						}
						counts.Add(count);
					}
				}
			}
			var q = ((TextBox)sender).Text ?? string.Empty;
			q = q.ToUpper();

			var records = strings.Select((value, index) => new { value, index }).Where(x => x.value.ToUpper().Contains(q)).OrderByDescending(y => counts[y.index]).Take(15).ToArray();

			Results.Clear();
			Lables.Clear();
			foreach (var record in records)
			{
				Results.Add(counts[record.index]);
				Lables.Add(record.value);
			}
		}

		private void cbPie_SelectedIndexChanged(object sender, EventArgs e)
		{

		}
	}
}