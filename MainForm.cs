using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DuctOrderApp.Helpers;
using DuctOrderApp.Models;
using DuctOrderApp.Services;

namespace DuctOrderApp.Forms
{
    /// <summary>
    /// Main application window.
    /// All UI is built in code – no Designer file required.
    /// </summary>
    public class MainForm : Form
    {
        // ── Services ─────────────────────────────────────────────────────────
        private readonly ExcelService _excel = new ExcelService();

        // ── Cached order list ─────────────────────────────────────────────────
        private List<OrderModel> _orders = new List<OrderModel>();

        // ── Input controls ────────────────────────────────────────────────────
        private TextBox   txtClient;
        private TextBox   txtOrderName;
        private ComboBox  cboDuctType;
        private CheckBox  chkUrgent;
        private CheckBox  chkDone;
        private Label     lblDateValue;

        // ── Toolbar / search ─────────────────────────────────────────────────
        private TextBox   txtSearch;
        private CheckBox  chkFilterUrgent;

        // ── Buttons ───────────────────────────────────────────────────────────
        private Button    btnSave;
        private Button    btnClear;
        private Button    btnLoad;
        private Button    btnEdit;
        private Button    btnDelete;
        private Button    btnExport;

        // ── Grid ─────────────────────────────────────────────────────────────
        private DataGridView grid;

        // ── Stats labels ─────────────────────────────────────────────────────
        private Label lblStatTotal;
        private Label lblStatUrgent;
        private Label lblStatDone;
        private Label lblStatPending;

        // ── Edit state ───────────────────────────────────────────────────────
        private bool _editMode    = false;
        private int  _editIndex   = -1;   // index in _orders list

        // ─────────────────────────────────────────────────────────────────────
        public MainForm()
        {
            InitializeComponents();
            LoadAndDisplayOrders();
        }

        // ═════════════════════════════════════════════════════════════════════
        //  UI CONSTRUCTION
        // ═════════════════════════════════════════════════════════════════════

        private void InitializeComponents()
        {
            // ── Form ─────────────────────────────────────────────────────────
            this.Text            = "Duct Order Management System";
            this.Size            = new Size(1200, 820);
            this.MinimumSize     = new Size(900, 650);
            this.StartPosition   = FormStartPosition.CenterScreen;
            this.BackColor       = ThemeHelper.BgForm;
            this.Font            = new Font("Segoe UI", 9f);

            // ── Title bar panel ───────────────────────────────────────────────
            var pnlTitle = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 56,
                BackColor = Color.FromArgb(20, 20, 20)
            };

            var lblTitle = new Label
            {
                Text      = "⬡  DUCT ORDER MANAGEMENT SYSTEM",
                Font      = ThemeHelper.FontTitle,
                ForeColor = ThemeHelper.Accent,
                AutoSize  = false,
                Dock      = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding   = new Padding(16, 0, 0, 0)
            };
            pnlTitle.Controls.Add(lblTitle);

            // ── Main split: left form panel  |  right grid panel ─────────────
            var splitMain = new SplitContainer
            {
                Dock             = DockStyle.Fill,
                Orientation      = Orientation.Vertical,
                SplitterDistance = 370,
                BackColor        = ThemeHelper.BgForm,
                Panel1MinSize    = 300,
                Panel2MinSize    = 400
            };

            BuildInputPanel(splitMain.Panel1);
            BuildGridPanel(splitMain.Panel2);

            this.Controls.Add(splitMain);
            this.Controls.Add(pnlTitle);   // added last → rendered on top (Dock.Top)
        }

        // ── Left panel (inputs + buttons + stats) ─────────────────────────────
        private void BuildInputPanel(SplitterPanel parent)
        {
            parent.BackColor = ThemeHelper.BgControl;
            parent.Padding   = new Padding(16);

            // ---- Stats strip (top) ------------------------------------------
            var pnlStats = new Panel
            {
                Dock      = DockStyle.Bottom,
                Height    = 72,
                BackColor = Color.FromArgb(28, 28, 28),
                Padding   = new Padding(10, 8, 10, 8)
            };
            lblStatTotal   = MakeStatLabel("TOTAL\n0",   ThemeHelper.Accent);
            lblStatUrgent  = MakeStatLabel("URGENT\n0",  ThemeHelper.AccentDanger);
            lblStatDone    = MakeStatLabel("DONE\n0",    ThemeHelper.AccentSuccess);
            lblStatPending = MakeStatLabel("PENDING\n0", Color.FromArgb(255, 180, 0));

            var statsFlow = new FlowLayoutPanel
            {
                Dock       = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents  = false,
            };
            statsFlow.Controls.AddRange(new Control[] { lblStatTotal, lblStatUrgent, lblStatDone, lblStatPending });
            pnlStats.Controls.Add(statsFlow);
            parent.Controls.Add(pnlStats);

            // ---- GroupBox: Order Details ------------------------------------
            var grp = new GroupBox
            {
                Text    = "  ORDER DETAILS",
                Dock    = DockStyle.Fill,
                Padding = new Padding(12, 18, 12, 12),
            };
            ThemeHelper.StyleGroupBox(grp);

            int labelW = 110;
            int inputH = 28;
            int gap    = 8;
            int y      = 28;

            // Client
            grp.Controls.Add(MakeLabel("CLIENT", labelW, y));
            txtClient = new TextBox { Left = labelW, Top = y, Width = grp.Width - labelW - 16, Height = inputH, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            ThemeHelper.StyleTextBox(txtClient);
            grp.Controls.Add(txtClient);
            y += inputH + gap;

            // Order Name
            grp.Controls.Add(MakeLabel("ORDER NAME", labelW, y));
            txtOrderName = new TextBox { Left = labelW, Top = y, Width = grp.Width - labelW - 16, Height = inputH, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            ThemeHelper.StyleTextBox(txtOrderName);
            grp.Controls.Add(txtOrderName);
            y += inputH + gap;

            // Duct Type
            grp.Controls.Add(MakeLabel("DUCT TYPE", labelW, y));
            cboDuctType = new ComboBox
            {
                Left   = labelW, Top = y, Width = grp.Width - labelW - 16, Height = inputH,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                DropDownStyle = ComboBoxStyle.DropDown
            };
            cboDuctType.Items.AddRange(new object[]
            {
                "Rectangular Duct", "Round Duct", "Oval Duct",
                "Flexible Duct", "Spiral Duct", "Flat Oval Duct",
                "Acoustic Duct", "Insulated Duct", "Custom"
            });
            ThemeHelper.StyleComboBox(cboDuctType);
            grp.Controls.Add(cboDuctType);
            y += inputH + gap;

            // Date Created (auto)
            grp.Controls.Add(MakeLabel("DATE CREATED", labelW, y));
            lblDateValue = new Label
            {
                Left      = labelW, Top = y, Height = inputH,
                Width     = grp.Width - labelW - 16,
                Anchor    = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                TextAlign = ContentAlignment.MiddleLeft,
                Text      = DateTime.Now.ToString("yyyy-MM-dd"),
                ForeColor = ThemeHelper.Accent,
                Font      = ThemeHelper.FontInput,
                BackColor = Color.Transparent
            };
            grp.Controls.Add(lblDateValue);
            y += inputH + gap + 4;

            // Checkboxes row
            chkUrgent = new CheckBox { Text = "⚡ URGENT", Left = labelW, Top = y, AutoSize = true };
            ThemeHelper.StyleCheckBox(chkUrgent);
            chkUrgent.ForeColor = ThemeHelper.AccentDanger;
            grp.Controls.Add(chkUrgent);

            chkDone = new CheckBox { Text = "✔ DONE", Left = labelW + 140, Top = y, AutoSize = true };
            ThemeHelper.StyleCheckBox(chkDone);
            chkDone.ForeColor = ThemeHelper.AccentSuccess;
            grp.Controls.Add(chkDone);
            y += 30 + gap * 2;

            // ---- Buttons ---------------------------------------------------
            int btnW = (grp.Width - labelW - 16 - 8) / 2;  // two-column layout

            btnSave = new Button { Text = "💾  SAVE ORDER",  Left = labelW,        Top = y, Width = grp.Width - labelW - 16, Height = 40, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            ThemeHelper.StyleButtonPrimary(btnSave);
            btnSave.Height = 40;
            grp.Controls.Add(btnSave);
            y += 44 + gap;

            btnClear  = new Button { Text = "🗑  CLEAR",    Left = labelW,        Top = y, Width = (grp.Width - labelW - 16 - 8) / 2, Height = 36, Anchor = AnchorStyles.Top | AnchorStyles.Left };
            btnLoad   = new Button { Text = "🔄  RELOAD",   Left = labelW + (grp.Width - labelW - 16 - 8) / 2 + 8, Top = y, Width = (grp.Width - labelW - 16 - 8) / 2, Height = 36, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            ThemeHelper.StyleButtonSecondary(btnClear);
            ThemeHelper.StyleButtonSecondary(btnLoad);
            grp.Controls.Add(btnClear);
            grp.Controls.Add(btnLoad);

            parent.Controls.Add(grp);

            // ---- Wire events ------------------------------------------------
            btnSave.Click  += BtnSave_Click;
            btnClear.Click += (s, e) => ClearForm();
            btnLoad.Click  += (s, e) => LoadAndDisplayOrders();
        }

        // ── Right panel (search bar + grid + action buttons) ──────────────────
        private void BuildGridPanel(SplitterPanel parent)
        {
            parent.BackColor = ThemeHelper.BgForm;
            parent.Padding   = new Padding(8, 16, 16, 16);

            // ---- Top toolbar ------------------------------------------------
            var pnlToolbar = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 46,
                BackColor = ThemeHelper.BgForm,
                Padding   = new Padding(0, 0, 0, 8)
            };

            txtSearch = new TextBox
            {
                PlaceholderText = "🔍  Search client or order…",
                Left   = 0, Top = 4, Width = 260, Height = 30,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            ThemeHelper.StyleTextBox(txtSearch);
            txtSearch.TextChanged += (s, e) => ApplyFilter();

            chkFilterUrgent = new CheckBox
            {
                Text     = "⚡ Urgent only",
                Left     = 272, Top = 8,
                AutoSize = true,
                ForeColor = ThemeHelper.AccentDanger
            };
            ThemeHelper.StyleCheckBox(chkFilterUrgent);
            chkFilterUrgent.CheckedChanged += (s, e) => ApplyFilter();

            // Action buttons on the right
            btnEdit   = new Button { Text = "✏ Edit",   Width = 88, Height = 30, Top = 4, Anchor = AnchorStyles.Top | AnchorStyles.Right, Right = parent.Width - 8 };
            btnDelete = new Button { Text = "✖ Delete", Width = 88, Height = 30, Top = 4, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnExport = new Button { Text = "📄 Export", Width = 88, Height = 30, Top = 4, Anchor = AnchorStyles.Top | AnchorStyles.Right };

            // position from right
            btnExport.Left = parent.Width - 288;
            btnDelete.Left = parent.Width - 192;
            btnEdit.Left   = parent.Width - 96;

            ThemeHelper.StyleButtonSecondary(btnEdit);
            ThemeHelper.StyleButtonDanger(btnDelete);
            ThemeHelper.StyleButtonSuccess(btnExport);

            pnlToolbar.Controls.AddRange(new Control[] { txtSearch, chkFilterUrgent, btnExport, btnDelete, btnEdit });
            parent.Controls.Add(pnlToolbar);

            // ---- DataGridView -----------------------------------------------
            grid = new DataGridView { Dock = DockStyle.Fill };
            ThemeHelper.StyleDataGridView(grid);
            parent.Controls.Add(grid);

            // ---- Wire events ------------------------------------------------
            btnEdit.Click   += BtnEdit_Click;
            btnDelete.Click += BtnDelete_Click;
            btnExport.Click += BtnExport_Click;
        }

        // ─────────────────────────────────────────────────────────────────────
        //  HELPERS – UI factories
        // ─────────────────────────────────────────────────────────────────────

        private Label MakeLabel(string text, int width, int top)
        {
            var lbl = new Label
            {
                Text      = text,
                Left      = 0, Top = top,
                Width     = width - 8, Height = 26,
                TextAlign = ContentAlignment.MiddleRight,
                BackColor = Color.Transparent
            };
            ThemeHelper.StyleLabel(lbl);
            return lbl;
        }

        private Label MakeStatLabel(string text, Color color)
        {
            return new Label
            {
                Text      = text,
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = color,
                BackColor = Color.Transparent,
                Width     = 72, Height = 50,
                TextAlign = ContentAlignment.MiddleCenter,
                Margin    = new Padding(4, 0, 4, 0)
            };
        }

        // ═════════════════════════════════════════════════════════════════════
        //  BUSINESS LOGIC
        // ═════════════════════════════════════════════════════════════════════

        /// <summary>Validate, save, reload.</summary>
        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidateInputs()) return;

                var order = BuildOrderFromForm();

                if (_editMode && _editIndex >= 0)
                {
                    // Replace existing entry
                    _orders[_editIndex] = order;
                    _excel.OverwriteOrders(_orders);
                    ExitEditMode();
                    ShowSuccess("Order updated successfully.");
                }
                else
                {
                    _excel.SaveOrder(order);
                    _orders.Add(order);
                    ShowSuccess("Order saved successfully.");
                }

                ClearForm();
                RefreshGrid(_orders);
                UpdateStats();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving order:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (grid.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a row to edit.", "No Selection",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Resolve the actual index in _orders (filter might have changed row positions)
            int gridRowIndex = grid.SelectedRows[0].Index;
            var row          = grid.Rows[gridRowIndex];
            int actualIndex  = FindOrderIndex(row);

            if (actualIndex < 0) return;

            _editMode  = true;
            _editIndex = actualIndex;

            var o = _orders[actualIndex];
            txtClient.Text        = o.Client;
            txtOrderName.Text     = o.OrderName;
            cboDuctType.Text      = o.DuctType;
            chkUrgent.Checked     = o.Urgent;
            chkDone.Checked       = o.Done;
            lblDateValue.Text     = o.DateCreated;

            btnSave.Text          = "💾  UPDATE ORDER";
            ThemeHelper.StyleButtonSuccess(btnSave);

            txtClient.Focus();
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (grid.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a row to delete.", "No Selection",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var confirm = MessageBox.Show(
                "Are you sure you want to delete this order?",
                "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes) return;

            try
            {
                int gridRowIndex = grid.SelectedRows[0].Index;
                int actualIndex  = FindOrderIndex(grid.Rows[gridRowIndex]);

                if (actualIndex < 0) return;

                _orders.RemoveAt(actualIndex);
                _excel.OverwriteOrders(_orders);
                RefreshGrid(_orders);
                UpdateStats();
                ShowSuccess("Order deleted.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error deleting order:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>Simple CSV export (PDF requires a 3rd-party library).</summary>
        private void BtnExport_Click(object sender, EventArgs e)
        {
            using (var dlg = new SaveFileDialog
            {
                Title    = "Export Orders",
                Filter   = "CSV file (*.csv)|*.csv",
                FileName = $"DuctOrders_{DateTime.Now:yyyyMMdd}"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                try
                {
                    var lines = new System.Collections.Generic.List<string>
                    {
                        "CLIENT,ORDER NAME,DUCT TYPE,URGENT,DATE CREATED,DONE"
                    };
                    foreach (var o in _orders)
                        lines.Add($"\"{o.Client}\",\"{o.OrderName}\",\"{o.DuctType}\"," +
                                  $"{(o.Urgent ? "YES" : "NO")},{o.DateCreated},{(o.Done ? "YES" : "NO")}");

                    System.IO.File.WriteAllLines(dlg.FileName, lines);
                    ShowSuccess($"Exported {_orders.Count} orders to CSV.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Export error:\n{ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // ── Data helpers ──────────────────────────────────────────────────────

        private void LoadAndDisplayOrders()
        {
            try
            {
                _excel.EnsureFileExists();
                _orders = _excel.LoadOrders();
                RefreshGrid(_orders);
                UpdateStats();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading orders:\n{ex.Message}", "Load Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshGrid(List<OrderModel> list)
        {
            grid.DataSource = null;
            grid.Rows.Clear();
            grid.Columns.Clear();

            if (list.Count == 0) return;

            // Build columns
            AddColumn("CLIENT",        "Client",       false);
            AddColumn("ORDER NAME",    "OrderName",    false);
            AddColumn("DUCT TYPE",     "DuctType",     false);
            AddColumn("URGENT",        "Urgent",       true,  80);
            AddColumn("DATE CREATED",  "DateCreated",  true,  110);
            AddColumn("DONE",          "Done",         true,  70);

            // Populate rows
            foreach (var o in list)
            {
                var row = new DataGridViewRow();
                row.CreateCells(grid,
                    o.Client,
                    o.OrderName,
                    o.DuctType,
                    o.Urgent  ? "YES" : "NO",
                    o.DateCreated,
                    o.Done    ? "YES" : "NO");

                // Highlight urgent rows
                if (o.Urgent)
                    row.DefaultCellStyle.ForeColor = ThemeHelper.AccentDanger;
                if (o.Done)
                    row.DefaultCellStyle.ForeColor = ThemeHelper.AccentSuccess;

                grid.Rows.Add(row);
            }
        }

        private void AddColumn(string header, string name, bool fixedWidth, int width = 0)
        {
            var col = new DataGridViewTextBoxColumn
            {
                HeaderText = header,
                Name       = name,
                SortMode   = DataGridViewColumnSortMode.Automatic
            };
            if (fixedWidth)
            {
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                col.Width        = width;
            }
            else
            {
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
            grid.Columns.Add(col);
        }

        private void ApplyFilter()
        {
            var query    = txtSearch.Text.Trim().ToLower();
            var urgentOnly = chkFilterUrgent.Checked;

            var filtered = _orders.Where(o =>
                (string.IsNullOrEmpty(query) ||
                 o.Client.ToLower().Contains(query) ||
                 o.OrderName.ToLower().Contains(query) ||
                 o.DuctType.ToLower().Contains(query))
                && (!urgentOnly || o.Urgent)
            ).ToList();

            RefreshGrid(filtered);
        }

        private void UpdateStats()
        {
            int total   = _orders.Count;
            int urgent  = _orders.Count(o => o.Urgent);
            int done    = _orders.Count(o => o.Done);
            int pending = total - done;

            lblStatTotal.Text   = $"TOTAL\n{total}";
            lblStatUrgent.Text  = $"URGENT\n{urgent}";
            lblStatDone.Text    = $"DONE\n{done}";
            lblStatPending.Text = $"PENDING\n{pending}";
        }

        private OrderModel BuildOrderFromForm()
        {
            return new OrderModel
            {
                Client      = txtClient.Text.Trim(),
                OrderName   = txtOrderName.Text.Trim(),
                DuctType    = cboDuctType.Text.Trim(),
                Urgent      = chkUrgent.Checked,
                DateCreated = _editMode ? lblDateValue.Text : DateTime.Now.ToString("yyyy-MM-dd"),
                Done        = chkDone.Checked
            };
        }

        private bool ValidateInputs()
        {
            if (string.IsNullOrWhiteSpace(txtClient.Text))
            {
                ShowValidationError("Client name is required.", txtClient);
                return false;
            }
            if (string.IsNullOrWhiteSpace(txtOrderName.Text))
            {
                ShowValidationError("Order name is required.", txtOrderName);
                return false;
            }
            if (string.IsNullOrWhiteSpace(cboDuctType.Text))
            {
                ShowValidationError("Duct type is required.", cboDuctType);
                return false;
            }
            return true;
        }

        private void ShowValidationError(string message, Control focusControl)
        {
            MessageBox.Show(message, "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            focusControl?.Focus();
        }

        private void ClearForm()
        {
            txtClient.Text    = string.Empty;
            txtOrderName.Text = string.Empty;
            cboDuctType.Text  = string.Empty;
            chkUrgent.Checked = false;
            chkDone.Checked   = false;
            lblDateValue.Text = DateTime.Now.ToString("yyyy-MM-dd");
            ExitEditMode();
            txtClient.Focus();
        }

        private void ExitEditMode()
        {
            _editMode   = false;
            _editIndex  = -1;
            btnSave.Text = "💾  SAVE ORDER";
            ThemeHelper.StyleButtonPrimary(btnSave);
        }

        /// <summary>
        /// Find the actual index in _orders for a given DataGridView row
        /// by matching all visible cell values (works even when filtered).
        /// </summary>
        private int FindOrderIndex(DataGridViewRow row)
        {
            string client    = row.Cells["Client"].Value?.ToString()    ?? string.Empty;
            string orderName = row.Cells["OrderName"].Value?.ToString() ?? string.Empty;
            string date      = row.Cells["DateCreated"].Value?.ToString() ?? string.Empty;

            return _orders.FindIndex(o =>
                o.Client      == client &&
                o.OrderName   == orderName &&
                o.DateCreated == date);
        }

        private void ShowSuccess(string message)
        {
            // Transient toast-style tooltip on the title bar area
            var lbl = new Label
            {
                Text      = "  ✔  " + message,
                BackColor = ThemeHelper.AccentSuccess,
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 9f, FontStyle.Bold),
                AutoSize  = false,
                Height    = 30,
                Dock      = DockStyle.Bottom,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding   = new Padding(8, 0, 0, 0)
            };
            this.Controls.Add(lbl);
            lbl.BringToFront();

            var timer = new Timer { Interval = 2500 };
            timer.Tick += (s, ev) =>
            {
                timer.Stop();
                this.Controls.Remove(lbl);
                lbl.Dispose();
            };
            timer.Start();
        }
    }
}
