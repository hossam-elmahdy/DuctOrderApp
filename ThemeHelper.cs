using System.Drawing;
using System.Windows.Forms;

namespace DuctOrderApp.Helpers
{
    /// <summary>
    /// Centralises all dark-mode colour and style constants.
    /// </summary>
    public static class ThemeHelper
    {
        // ── Palette ──────────────────────────────────────────────────────────
        public static readonly Color BgForm        = Color.FromArgb(35,  35,  35);
        public static readonly Color BgControl     = Color.FromArgb(45,  45,  45);
        public static readonly Color BgInput       = Color.FromArgb(55,  55,  55);
        public static readonly Color BgHeader      = Color.FromArgb(25,  25,  25);
        public static readonly Color Accent        = Color.FromArgb(30, 136, 229);   // Blue
        public static readonly Color AccentHover   = Color.FromArgb(21,  101, 192);
        public static readonly Color AccentDanger  = Color.FromArgb(229,  57,  53);  // Red (delete)
        public static readonly Color AccentSuccess = Color.FromArgb(67,  160,  71);  // Green
        public static readonly Color TextPrimary   = Color.White;
        public static readonly Color TextSecondary = Color.FromArgb(180, 180, 180);
        public static readonly Color GridLine      = Color.FromArgb(60,  60,  60);
        public static readonly Color RowAlt        = Color.FromArgb(50,  50,  50);
        public static readonly Color RowSelected   = Color.FromArgb(30, 100, 200);

        // ── Typography ───────────────────────────────────────────────────────
        public static readonly Font FontLabel   = new Font("Segoe UI", 9f,  FontStyle.Bold);
        public static readonly Font FontInput   = new Font("Segoe UI", 9.5f);
        public static readonly Font FontButton  = new Font("Segoe UI", 9.5f, FontStyle.Bold);
        public static readonly Font FontTitle   = new Font("Segoe UI", 14f, FontStyle.Bold);
        public static readonly Font FontGrid    = new Font("Segoe UI", 9f);

        // ── Control Styling ──────────────────────────────────────────────────

        public static void StyleTextBox(TextBox tb)
        {
            tb.BackColor  = BgInput;
            tb.ForeColor  = TextPrimary;
            tb.Font       = FontInput;
            tb.BorderStyle = BorderStyle.FixedSingle;
        }

        public static void StyleComboBox(ComboBox cb)
        {
            cb.BackColor  = BgInput;
            cb.ForeColor  = TextPrimary;
            cb.Font       = FontInput;
            cb.FlatStyle  = FlatStyle.Flat;
        }

        public static void StyleCheckBox(CheckBox chk)
        {
            chk.BackColor = Color.Transparent;
            chk.ForeColor = TextPrimary;
            chk.Font      = FontInput;
        }

        public static void StyleLabel(Label lbl)
        {
            lbl.BackColor = Color.Transparent;
            lbl.ForeColor = TextSecondary;
            lbl.Font      = FontLabel;
        }

        /// <summary>Primary action button (blue accent).</summary>
        public static void StyleButtonPrimary(Button btn)
        {
            ApplyFlatButton(btn, Accent);
        }

        /// <summary>Secondary / neutral button.</summary>
        public static void StyleButtonSecondary(Button btn)
        {
            ApplyFlatButton(btn, Color.FromArgb(70, 70, 70));
        }

        /// <summary>Danger button (red accent).</summary>
        public static void StyleButtonDanger(Button btn)
        {
            ApplyFlatButton(btn, AccentDanger);
        }

        /// <summary>Success button (green accent).</summary>
        public static void StyleButtonSuccess(Button btn)
        {
            ApplyFlatButton(btn, AccentSuccess);
        }

        private static void ApplyFlatButton(Button btn, Color backColor)
        {
            btn.FlatStyle               = FlatStyle.Flat;
            btn.BackColor               = backColor;
            btn.ForeColor               = Color.White;
            btn.Font                    = FontButton;
            btn.FlatAppearance.BorderSize = 0;
            btn.Cursor                  = Cursors.Hand;
            btn.Height                  = 36;
        }

        public static void StyleDataGridView(DataGridView dgv)
        {
            dgv.BackgroundColor          = BgControl;
            dgv.GridColor                = GridLine;
            dgv.BorderStyle              = BorderStyle.None;
            dgv.CellBorderStyle          = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dgv.EnableHeadersVisualStyles = false;
            dgv.Font                     = FontGrid;
            dgv.RowHeadersVisible        = false;
            dgv.SelectionMode            = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect              = false;
            dgv.ReadOnly                 = true;
            dgv.AllowUserToAddRows       = false;
            dgv.AllowUserToDeleteRows    = false;
            dgv.AutoSizeColumnsMode      = DataGridViewAutoSizeColumnsMode.Fill;

            // Column headers
            dgv.ColumnHeadersDefaultCellStyle.BackColor  = BgHeader;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor  = TextSecondary;
            dgv.ColumnHeadersDefaultCellStyle.Font       = new Font("Segoe UI", 9f, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = BgHeader;
            dgv.ColumnHeadersHeight = 34;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            // Default cell style
            dgv.DefaultCellStyle.BackColor          = BgControl;
            dgv.DefaultCellStyle.ForeColor          = TextPrimary;
            dgv.DefaultCellStyle.SelectionBackColor = RowSelected;
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.DefaultCellStyle.Padding            = new Padding(4, 0, 4, 0);

            // Alternate row
            dgv.AlternatingRowsDefaultCellStyle.BackColor          = RowAlt;
            dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = RowSelected;
            dgv.AlternatingRowsDefaultCellStyle.ForeColor          = TextPrimary;

            dgv.RowTemplate.Height = 28;
        }

        public static void StyleGroupBox(GroupBox gb)
        {
            gb.BackColor = BgControl;
            gb.ForeColor = TextSecondary;
            gb.Font      = FontLabel;
        }

        public static void StylePanel(Panel pnl)
        {
            pnl.BackColor = BgControl;
        }
    }
}
