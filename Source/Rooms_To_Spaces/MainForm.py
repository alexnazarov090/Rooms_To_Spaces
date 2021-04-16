import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._tableLayoutPanel = System.Windows.Forms.TableLayoutPanel()
		self._file_path_label = System.Windows.Forms.Label()
		self._file_path_textBox = System.Windows.Forms.TextBox()
		self._browse_button = System.Windows.Forms.Button()
		self._shar_pars_label = System.Windows.Forms.Label()
		self._shar_pars_checkedListBox = System.Windows.Forms.CheckedListBox()
		self._shar_pars_check_all_btn = System.Windows.Forms.Button()
		self._shar_pars_check_none_btn = System.Windows.Forms.Button()
		self._builtin_pars_label = System.Windows.Forms.Label()
		self._builtin_pars_checkedListBox = System.Windows.Forms.CheckedListBox()
		self._builtin_pars_check_all_btn = System.Windows.Forms.Button()
		self._builtin_pars_check_none_btn = System.Windows.Forms.Button()
		self._progressBar = System.Windows.Forms.ProgressBar()
		self._run_button = System.Windows.Forms.Button()
		self._space_id_label = System.Windows.Forms.Label()
		self._space_id_comboBox = System.Windows.Forms.ComboBox()
		self._write_exl_button = System.Windows.Forms.Button()
		self._write_exl_checkBox = System.Windows.Forms.CheckBox()
		self._del_spaces_button = System.Windows.Forms.Button()
		self._tableLayoutPanel.SuspendLayout()
		self.SuspendLayout()
		# 
		# tableLayoutPanel
		# 
		self._tableLayoutPanel.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
		self._tableLayoutPanel.BackColor = System.Drawing.SystemColors.Control
		self._tableLayoutPanel.ColumnCount = 3
		self._tableLayoutPanel.ColumnStyles.Add(System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.3333321))
		self._tableLayoutPanel.ColumnStyles.Add(System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.3333321))
		self._tableLayoutPanel.ColumnStyles.Add(System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.3333321))
		self._tableLayoutPanel.Controls.Add(self._file_path_label, 0, 2)
		self._tableLayoutPanel.Controls.Add(self._run_button, 2, 13)
		self._tableLayoutPanel.Controls.Add(self._file_path_textBox, 0, 3)
		self._tableLayoutPanel.Controls.Add(self._browse_button, 2, 3)
		self._tableLayoutPanel.Controls.Add(self._shar_pars_label, 0, 4)
		self._tableLayoutPanel.Controls.Add(self._shar_pars_checkedListBox, 0, 5)
		self._tableLayoutPanel.Controls.Add(self._shar_pars_check_all_btn, 2, 5)
		self._tableLayoutPanel.Controls.Add(self._shar_pars_check_none_btn, 2, 6)
		self._tableLayoutPanel.Controls.Add(self._builtin_pars_label, 0, 8)
		self._tableLayoutPanel.Controls.Add(self._builtin_pars_checkedListBox, 0, 9)
		self._tableLayoutPanel.Controls.Add(self._builtin_pars_check_all_btn, 2, 9)
		self._tableLayoutPanel.Controls.Add(self._builtin_pars_check_none_btn, 2, 10)
		self._tableLayoutPanel.Controls.Add(self._progressBar, 0, 14)
		self._tableLayoutPanel.Controls.Add(self._space_id_label, 0, 0)
		self._tableLayoutPanel.Controls.Add(self._space_id_comboBox, 0, 1)
		self._tableLayoutPanel.Controls.Add(self._write_exl_button, 0, 13)
		self._tableLayoutPanel.Controls.Add(self._write_exl_checkBox, 2, 11)
		self._tableLayoutPanel.Controls.Add(self._del_spaces_button, 1, 13)
		self._tableLayoutPanel.Font = System.Drawing.Font("Microsoft Sans Serif", 9.75, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 204)
		self._tableLayoutPanel.Location = System.Drawing.Point(12, 12)
		self._tableLayoutPanel.Name = "tableLayoutPanel"
		self._tableLayoutPanel.Padding = System.Windows.Forms.Padding(3, 0, 3, 0)
		self._tableLayoutPanel.RowCount = 15
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle())
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle())
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle())
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle())
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle())
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 7.728637))
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 7.7286377))
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 19.3548374))
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.564821))
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 7.486645))
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 7.486645))
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 3.74332261))
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.7160378))
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 17.6824284))
		self._tableLayoutPanel.RowStyles.Add(System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 7.507986))
		self._tableLayoutPanel.Size = System.Drawing.Size(460, 537)
		self._tableLayoutPanel.TabIndex = 0
		# 
		# file_path_label
		# 
		self._file_path_label.Dock = System.Windows.Forms.DockStyle.Bottom
		self._file_path_label.Location = System.Drawing.Point(6, 50)
		self._file_path_label.Name = "file_path_label"
		self._file_path_label.Size = System.Drawing.Size(145, 20)
		self._file_path_label.TabIndex = 0
		self._file_path_label.Text = "Excel File Path:"
		# 
		# file_path_textBox
		# 
		self._tableLayoutPanel.SetColumnSpan(self._file_path_textBox, 2)
		self._file_path_textBox.Dock = System.Windows.Forms.DockStyle.Fill
		self._file_path_textBox.Location = System.Drawing.Point(6, 73)
		self._file_path_textBox.Name = "file_path_textBox"
		self._file_path_textBox.ReadOnly = True
		self._file_path_textBox.Size = System.Drawing.Size(296, 22)
		self._file_path_textBox.TabIndex = 1
		# 
		# browse_button
		# 
		self._browse_button.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
		self._browse_button.BackColor = System.Drawing.SystemColors.HotTrack
		self._browse_button.FlatStyle = System.Windows.Forms.FlatStyle.Popup
		self._browse_button.ForeColor = System.Drawing.SystemColors.HighlightText
		self._browse_button.Location = System.Drawing.Point(354, 73)
		self._browse_button.Name = "browse_button"
		self._browse_button.Size = System.Drawing.Size(100, 27)
		self._browse_button.TabIndex = 2
		self._browse_button.Text = "Browse"
		self._browse_button.UseVisualStyleBackColor = False
		# 
		# shar_pars_label
		# 
		self._shar_pars_label.Dock = System.Windows.Forms.DockStyle.Bottom
		self._shar_pars_label.Location = System.Drawing.Point(6, 103)
		self._shar_pars_label.Name = "shar_pars_label"
		self._shar_pars_label.Size = System.Drawing.Size(145, 20)
		self._shar_pars_label.TabIndex = 3
		self._shar_pars_label.Text = "Shared Parameters:"
		# 
		# shar_pars_checkedListBox
		# 
		self._shar_pars_checkedListBox.CheckOnClick = True
		self._tableLayoutPanel.SetColumnSpan(self._shar_pars_checkedListBox, 2)
		self._shar_pars_checkedListBox.Dock = System.Windows.Forms.DockStyle.Fill
		self._shar_pars_checkedListBox.FormattingEnabled = True
		self._shar_pars_checkedListBox.Location = System.Drawing.Point(6, 126)
		self._shar_pars_checkedListBox.Name = "shar_pars_checkedListBox"
		self._tableLayoutPanel.SetRowSpan(self._shar_pars_checkedListBox, 3)
		self._shar_pars_checkedListBox.Size = System.Drawing.Size(296, 136)
		self._shar_pars_checkedListBox.TabIndex = 4
		# 
		# shar_pars_check_all_btn
		# 
		self._shar_pars_check_all_btn.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
		self._shar_pars_check_all_btn.BackColor = System.Drawing.SystemColors.HotTrack
		self._shar_pars_check_all_btn.FlatStyle = System.Windows.Forms.FlatStyle.Popup
		self._shar_pars_check_all_btn.ForeColor = System.Drawing.SystemColors.HighlightText
		self._shar_pars_check_all_btn.Location = System.Drawing.Point(354, 126)
		self._shar_pars_check_all_btn.Name = "shar_pars_check_all_btn"
		self._shar_pars_check_all_btn.Size = System.Drawing.Size(100, 25)
		self._shar_pars_check_all_btn.TabIndex = 5
		self._shar_pars_check_all_btn.Text = "Check All"
		self._shar_pars_check_all_btn.UseVisualStyleBackColor = False
		# 
		# shar_pars_check_none_btn
		# 
		self._shar_pars_check_none_btn.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right
		self._shar_pars_check_none_btn.BackColor = System.Drawing.SystemColors.HotTrack
		self._shar_pars_check_none_btn.FlatStyle = System.Windows.Forms.FlatStyle.Popup
		self._shar_pars_check_none_btn.ForeColor = System.Drawing.SystemColors.HighlightText
		self._shar_pars_check_none_btn.Location = System.Drawing.Point(354, 157)
		self._shar_pars_check_none_btn.Name = "shar_pars_check_none_btn"
		self._shar_pars_check_none_btn.Size = System.Drawing.Size(100, 25)
		self._shar_pars_check_none_btn.TabIndex = 6
		self._shar_pars_check_none_btn.Text = "Check None"
		self._shar_pars_check_none_btn.UseVisualStyleBackColor = False
		# 
		# builtin_pars_label
		# 
		self._tableLayoutPanel.SetColumnSpan(self._builtin_pars_label, 2)
		self._builtin_pars_label.Dock = System.Windows.Forms.DockStyle.Bottom
		self._builtin_pars_label.Location = System.Drawing.Point(6, 265)
		self._builtin_pars_label.Name = "builtin_pars_label"
		self._builtin_pars_label.Size = System.Drawing.Size(296, 18)
		self._builtin_pars_label.TabIndex = 7
		self._builtin_pars_label.Text = "Built-in Parameters:"
		# 
		# builtin_pars_checkedListBox
		# 
		self._builtin_pars_checkedListBox.CheckOnClick = True
		self._tableLayoutPanel.SetColumnSpan(self._builtin_pars_checkedListBox, 2)
		self._builtin_pars_checkedListBox.Dock = System.Windows.Forms.DockStyle.Fill
		self._builtin_pars_checkedListBox.FormattingEnabled = True
		self._builtin_pars_checkedListBox.Location = System.Drawing.Point(6, 286)
		self._builtin_pars_checkedListBox.Name = "builtin_pars_checkedListBox"
		self._tableLayoutPanel.SetRowSpan(self._builtin_pars_checkedListBox, 4)
		self._builtin_pars_checkedListBox.Size = System.Drawing.Size(296, 138)
		self._builtin_pars_checkedListBox.TabIndex = 8
		# 
		# builtin_pars_check_all_btn
		# 
		self._builtin_pars_check_all_btn.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
		self._builtin_pars_check_all_btn.BackColor = System.Drawing.SystemColors.Highlight
		self._builtin_pars_check_all_btn.FlatStyle = System.Windows.Forms.FlatStyle.Popup
		self._builtin_pars_check_all_btn.ForeColor = System.Drawing.SystemColors.HighlightText
		self._builtin_pars_check_all_btn.Location = System.Drawing.Point(354, 286)
		self._builtin_pars_check_all_btn.Name = "builtin_pars_check_all_btn"
		self._builtin_pars_check_all_btn.Size = System.Drawing.Size(100, 24)
		self._builtin_pars_check_all_btn.TabIndex = 9
		self._builtin_pars_check_all_btn.Text = "Check All"
		self._builtin_pars_check_all_btn.UseVisualStyleBackColor = False
		# 
		# builtin_pars_check_none_btn
		# 
		self._builtin_pars_check_none_btn.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right
		self._builtin_pars_check_none_btn.BackColor = System.Drawing.SystemColors.Highlight
		self._builtin_pars_check_none_btn.FlatStyle = System.Windows.Forms.FlatStyle.Popup
		self._builtin_pars_check_none_btn.ForeColor = System.Drawing.SystemColors.HighlightText
		self._builtin_pars_check_none_btn.Location = System.Drawing.Point(354, 316)
		self._builtin_pars_check_none_btn.Name = "builtin_pars_check_none_btn"
		self._builtin_pars_check_none_btn.Size = System.Drawing.Size(100, 24)
		self._builtin_pars_check_none_btn.TabIndex = 10
		self._builtin_pars_check_none_btn.Text = "Check None"
		self._builtin_pars_check_none_btn.UseVisualStyleBackColor = False
		# 
		# progressBar
		# 
		self._progressBar.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
		self._progressBar.BackColor = System.Drawing.SystemColors.Control
		self._tableLayoutPanel.SetColumnSpan(self._progressBar, 3)
		self._progressBar.Location = System.Drawing.Point(6, 512)
		self._progressBar.Name = "progressBar"
		self._progressBar.Size = System.Drawing.Size(448, 22)
		self._progressBar.TabIndex = 12
		# 
		# run_button
		# 
		self._run_button.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
		self._run_button.AutoSize = True
		self._run_button.BackColor = System.Drawing.SystemColors.Highlight
		self._run_button.Enabled = False
		self._run_button.FlatStyle = System.Windows.Forms.FlatStyle.Popup
		self._run_button.Font = System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 204)
		self._run_button.ForeColor = System.Drawing.SystemColors.HighlightText
		self._run_button.Location = System.Drawing.Point(308, 447)
		self._run_button.Name = "run_button"
		self._run_button.Size = System.Drawing.Size(146, 50)
		self._run_button.TabIndex = 11
		self._run_button.Text = "Create/update spaces"
		self._run_button.UseVisualStyleBackColor = False
		# 
		# space_id_label
		# 
		self._tableLayoutPanel.SetColumnSpan(self._space_id_label, 2)
		self._space_id_label.Dock = System.Windows.Forms.DockStyle.Bottom
		self._space_id_label.Location = System.Drawing.Point(6, 0)
		self._space_id_label.Name = "space_id_label"
		self._space_id_label.Size = System.Drawing.Size(296, 20)
		self._space_id_label.TabIndex = 13
		self._space_id_label.Text = "Space ID Parameter:"
		# 
		# space_id_comboBox
		# 
		self._tableLayoutPanel.SetColumnSpan(self._space_id_comboBox, 2)
		self._space_id_comboBox.Dock = System.Windows.Forms.DockStyle.Bottom
		self._space_id_comboBox.DropDownHeight = 150
		self._space_id_comboBox.FormattingEnabled = True
		self._space_id_comboBox.IntegralHeight = False
		self._space_id_comboBox.ItemHeight = 16
		self._space_id_comboBox.Location = System.Drawing.Point(6, 23)
		self._space_id_comboBox.MaxDropDownItems = 10
		self._space_id_comboBox.Name = "space_id_comboBox"
		self._space_id_comboBox.Size = System.Drawing.Size(296, 24)
		self._space_id_comboBox.TabIndex = 14
		# 
		# write_exl_button
		# 
		self._write_exl_button.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
		self._write_exl_button.AutoSize = True
		self._write_exl_button.BackColor = System.Drawing.SystemColors.Highlight
		self._write_exl_button.Enabled = False
		self._write_exl_button.FlatStyle = System.Windows.Forms.FlatStyle.Popup
		self._write_exl_button.Font = System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 204)
		self._write_exl_button.ForeColor = System.Drawing.SystemColors.HighlightText
		self._write_exl_button.Location = System.Drawing.Point(6, 447)
		self._write_exl_button.Name = "write_exl_button"
		self._write_exl_button.Size = System.Drawing.Size(145, 50)
		self._write_exl_button.TabIndex = 15
		self._write_exl_button.Text = "Write to Excel"
		self._write_exl_button.UseVisualStyleBackColor = False
		# 
		# write_exl_checkBox
		# 
		self._write_exl_checkBox.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
		self._write_exl_checkBox.CheckAlign = System.Drawing.ContentAlignment.TopRight
		self._write_exl_checkBox.Enabled = False
		self._write_exl_checkBox.Location = System.Drawing.Point(308, 359)
		self._write_exl_checkBox.Name = "write_exl_checkBox"
		self._write_exl_checkBox.Padding = System.Windows.Forms.Padding(3, 0, 3, 0)
		self._tableLayoutPanel.SetRowSpan(self._write_exl_checkBox, 2)
		self._write_exl_checkBox.Size = System.Drawing.Size(146, 65)
		self._write_exl_checkBox.TabIndex = 16
		self._write_exl_checkBox.Text = "Write to Excel when creating/updating spaces?"
		self._write_exl_checkBox.TextAlign = System.Drawing.ContentAlignment.TopLeft
		self._write_exl_checkBox.UseVisualStyleBackColor = True
		# 
		# del_spaces_button
		# 
		self._del_spaces_button.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
		self._del_spaces_button.AutoSize = True
		self._del_spaces_button.BackColor = System.Drawing.SystemColors.Highlight
		self._del_spaces_button.Enabled = False
		self._del_spaces_button.FlatStyle = System.Windows.Forms.FlatStyle.Popup
		self._del_spaces_button.Font = System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 204)
		self._del_spaces_button.ForeColor = System.Drawing.SystemColors.HighlightText
		self._del_spaces_button.Location = System.Drawing.Point(157, 447)
		self._del_spaces_button.Name = "del_spaces_button"
		self._del_spaces_button.Size = System.Drawing.Size(145, 50)
		self._del_spaces_button.TabIndex = 17
		self._del_spaces_button.Text = "Delete Spaces"
		self._del_spaces_button.UseVisualStyleBackColor = False
		# 
		# MainForm
		# 
		self.BackColor = System.Drawing.SystemColors.ControlLightLight
		self.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		self.ClientSize = System.Drawing.Size(484, 561)
		self.Controls.Add(self._tableLayoutPanel)
		self.MinimumSize = System.Drawing.Size(500, 600)
		self.Name = "MainForm"
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		self.Text = "Create/update Spaces from Rooms"
		self._tableLayoutPanel.ResumeLayout(False)
		self._tableLayoutPanel.PerformLayout()
		self.ResumeLayout(False)
