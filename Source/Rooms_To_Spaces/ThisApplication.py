#!/usr/bin/env python
# -*- coding: utf-8 -*-

import clr
import System
clr.AddReference("RevitAPI.dll")
clr.AddReference("RevitAPIUI.dll")
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference("Microsoft.Office.Interop.Excel")

from Autodesk.Revit import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Macros import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI.Selection import *
from System.Collections.Generic import *
from System.Runtime.InteropServices import Marshal
from System.Threading import Thread
from System.ComponentModel import BackgroundWorker
import System.Windows.Forms as WinForms
import System.Drawing
import Microsoft.Office.Interop.Excel as Excel
import sys
path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(path)
import os
sys.path.append(os.path.realpath(__file__))
from MainForm import MainForm
import string
import re
from itertools import izip, product
import traceback
import logging
import json
import io

# Logger
dir_path = os.path.dirname(os.path.realpath(__file__))
logging.basicConfig(level=logging.DEBUG,
                    filename=r'{}'.format(os.path.join(dir_path, 'ThisApplication.log')),
                    filemode='w',
                    format='%(asctime)s %(name)s %(levelname)s:%(message)s')
logger = logging.getLogger(__name__)


class ThisApplication (ApplicationEntryPoint):
    # region Revit Macros generated code
    def FinishInitialization(self):
        ApplicationEntryPoint.FinishInitialization(self)
        self.InternalStartup()

    def OnShutdown(self):
        self.InternalShutdown()
        ApplicationEntryPoint.OnShutdown(self)

    def InternalStartup(self):
        self.Startup()

    def InternalShutdown(self):
        self.Shutdown()
    # endregion

    def Startup(self):
        self

    def Shutdown(self):
        self

    # Transaction mode

    def GetTransactionMode(self):
        return Attributes.TransactionMode.Manual

    # Addin Id
    def GetAddInId(self):
        return '2E9CE0D0-6090-4FF2-94A0-06E025DBC3CD'

    def Rooms_To_Spaces(self):

        # region Initial parameters
        search_id = "REFERENCED_ROOM_UNIQUE_ID"
        excel_parameters = [search_id, "ID_revit", "101_MW_AR_BUILDINGBLOCKS", "102_MW_AR_LEVEL", "103_MW_AR_REGION", "106_MW_AR_ENDNUMBER", "Number", "Name", "111_Name_ENG", "105_MW_AR_REGION_NAME_ENG",
                           "110_MW_CRCLASS", "Area", "Unbounded Height", "112_MW_PS_PRCLASS", "117_MW_PS_CONTRPOINT", "113_MW_PS_TEMP_MAX_CLASS", "114_MW_PS_TEMP_MIN_CLASS", "115_MW_PS_HUMID_MAX_CLASS",
                           "116_MW_PS_HUMID_MIN_CLASS", "ROOM_OVERFLOW", "108_MW_AR_FIRECAT", "127_MW_PS_SUBREION", "120_MW_FP_PEOPLE", "121_FP_PEOPLE_WORKMODE"]
        New_MEPSpaces = {}
        Exist_MEPSpaces = {}
        # link_name = "C_CNP_AR_detached.rvt"
        # endregion

        # region Get Document and Application
        doc = self.ActiveUIDocument.Document
        uidoc = UIDocument(doc)
        app = self.Application
        uiapp = UIApplication(app)
        docCreation = doc.Create
        excel = Excel.ApplicationClass()
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo("en-US")
        # endregion

        def create_spaces_by_rooms(rooms):

            # Initiate the transacton group
            with TransactionGroup(doc, "Batch create spaces/Transfer parameters") as tg:
                tg.Start()

                Room_UniqueIds = [room.UniqueId for room in rooms if room.Area > 0 and room.Location != None]

                # Define if there're spaces in a model
                if _space_collector.GetElementCount() == 0:
                    # If there are no spaces
                    # Create a space
                    for room in rooms:
                        if room.Area > 0 and room.UniqueId not in New_MEPSpaces:
                            space_creator(room, lvls_dict)

                # If there are spaces in the model
                else:
                    space_check(Room_UniqueIds)
                    # For each room in RVT link rooms take room UniqueId and check if there's space with the same Id among the existing MEP spaces
                    for room in rooms:
                        # If there's such space get it
                        if room.UniqueId in Exist_MEPSpaces:
                            exst_space = Exist_MEPSpaces[room.UniqueId]
                            space_updater(room, exst_space)
                            
                        # If there's no such space
                        else:
                            # Create a space
                            if room.Area > 0 and room.UniqueId not in New_MEPSpaces:
                                space_creator(room, lvls_dict)

                tg.Assimilate()

        def space_updater(room, exst_space):
            ''' 
            Function updates existing spaces and moves them if necessary
            '''
            # And extract its coordinates
            try:
                exst_space_coords = exst_space.Location.Point
                # Get room coordinates
                room_coords = room.Location.Point
                # Compare two sets of coordinates
                # If they are almost the same
                if exst_space_coords.IsAlmostEqualTo(room_coords):
                    with Transaction(doc, "Update parameters") as tr:
                        tr.Start()
                        options = tr.GetFailureHandlingOptions()
                        failureHandler = ErrorSwallower()
                        options.SetFailuresPreprocessor(failureHandler)
                        options.SetClearAfterRollback(True)
                        tr.SetFailureHandlingOptions(options)
                        # Transfer room parameters to a corresponding space
                        para_setter(room, exst_space)
                        tr.Commit()

                # Otherwise, move the existing space according to room coordinates
                else:
                    move_vector = room.Location.Point - exst_space.Location.Point
                    with Transaction(doc, "Move spaces") as tr:
                        tr.Start()

                        options = tr.GetFailureHandlingOptions()
                        failureHandler = ErrorSwallower()
                        options.SetFailuresPreprocessor(failureHandler)
                        options.SetClearAfterRollback(True)
                        tr.SetFailureHandlingOptions(options)

                        exst_space.Location.Move(move_vector)
                        # Transfer room parameters to a corresponding space
                        para_setter(room, exst_space)
                        tr.Commit()

            except System.MissingMemberException:
                pass
                
        def space_check(Room_UniqueIds):
            ''' 
            Function checks for not placed spaces, obsolete spaces and deletes them
            '''
            # Collect all existing MEP spaces in this document
            for space in _space_collector.ToElements():
                # Get "REFERENCED_ROOM_UNIQUE_ID" parameter of each space
                ID_par_list = space.GetParameters(search_id)
                ref_id_par_val = ID_par_list[0].AsString()
                # Check if REFERENCED_ROOM_UNIQUE_ID is in room Room_UniqueIds
                # If it's not the case delete the corresponding space
                if space.Area == 0 or ref_id_par_val not in Room_UniqueIds:
                    with Transaction(doc, "Delete spaces") as tr:
                        tr.Start()
                        try:
                            doc.Delete(space.Id)
                        except:
                            pass
                        tr.Commit()
                else:
                    # Otherwise cast it into Existing MEP spaces dictionary
                    Exist_MEPSpaces[ref_id_par_val] = Exist_MEPSpaces.get(
                        ref_id_par_val, space)

        def para_setter(r, space):
            ''' 
            Function transers parameters from the room to newly created spaces
            r: room, space: MEPSpace
            '''

            # For each parameter in room parameters define if it's shared or builtin parameter
            for par in r.Parameters:
                if par.IsShared:
                    # If room parameter is shared get space from Spaces dictionary by UniqueId and extract corresponding space parameter from it by room parameter GUID
                    # Depending on the room parameter storage type set its value to the space parameter
                    try:
                        space_par = space.get_Parameter(par.GUID)
                        if par.StorageType == StorageType.String and par.HasValue == True:
                            space_par.Set(par.AsString())

                        elif par.StorageType == StorageType.Integer and par.HasValue == True:
                            space_par.Set(par.AsInteger())

                        elif par.StorageType == StorageType.Double and par.HasValue == True:
                            space_par.Set(par.AsDouble())
                    except:
                        pass

                elif par.Definition.BuiltInParameter != BuiltInParameter.INVALID:
                    # If room parameter is builtin get space from Spaces dictionary by UniqueId and extract corresponding space parameter from it by builtin parameter
                    # Depending on the room parameter storage type set its value to the space parameter
                    try:
                        space_par = space.get_Parameter(par.Definition)
                        if par.StorageType == StorageType.String and par.HasValue == True:
                            space_par.Set(par.AsString())

                        elif par.StorageType == StorageType.Integer and par.HasValue == True:
                            space_par.Set(par.AsInteger())

                        elif par.StorageType == StorageType.Double and par.HasValue == True:
                            space_par.Set(par.AsDouble())
                    except:
                        pass

        def space_creator(r, lvls):
            ''' 
            Function creates new spaces 
            r: room, lvls: level elevations and levels dictionary
            '''
            try:
                # Get the room level
                point = r.Location.Point
                room_lvl = r.Level
                # Get a level from the lvls dictionary by room level elevation
                lvl = lvls[room_lvl.Elevation]
                # Create space by coordinates and level taken from room
                with Transaction(doc, "Batch create spaces") as tr:
                    tr.Start()
                    
                    options = tr.GetFailureHandlingOptions()
                    failureHandler = ErrorSwallower()
                    options.SetFailuresPreprocessor(failureHandler)
                    options.SetClearAfterRollback(True)
                    tr.SetFailureHandlingOptions(options)
                    space = docCreation.NewSpace(lvl, UV(point.X, point.Y))
                    
                    # Get "REFERENCED_ROOM_UNIQUE_ID" parameter
                    ID_par_list = space.GetParameters(search_id)
                    ref_id_par = ID_par_list[0]

                    # Assign room UniqueID to "REFERENCED_ROOM_UNIQUE_ID" parameter
                    if ref_id_par:
                        ref_id_par.Set(r.UniqueId)
                        New_MEPSpaces[r.UniqueId] = New_MEPSpaces.get(r.UniqueId, space)
                        para_setter(r, space)
                    

                    # Get "ID_revit" parameter of a MEP space
                    space_id_par_list = space.GetParameters("ID_revit")
                    space_id = space_id_par_list[0]

                    # Assign space ID to "ID_revit" parameter
                    if space_id:
                        space_id.Set(space.Id.IntegerValue)

                    tr.Commit()

            except:
                pass

        def merge_two_dicts(d1, d2):
            '''
            Merges two dictionaries
            Returns a merged dictionary
            '''
            d_merged = d1.copy()
            d_merged.update(d2)
            return d_merged

        def write_to_excel(params_to_write):

            """
            Writing parameters to an Excel workbook
            """
            units = {UnitType.UT_Length: DisplayUnitType.DUT_METERS, 
                    UnitType.UT_Area: DisplayUnitType.DUT_SQUARE_METERS, 
                    UnitType.UT_HVAC_Airflow: DisplayUnitType.DUT_CUBIC_METERS_PER_HOUR}
            params = dict(izip(params_to_write, string.ascii_uppercase))
            rvt_filepath = doc.PathName
            xl_path = re.sub('.rvt$', '_Ex.xlsx', rvt_filepath)
            try:
                workBook = excel.Workbooks.Open(r'{}'.format(xl_path))
                workSheet = workBook.Worksheets(1)
            
            except:
                workBook = excel.Workbooks.Add()
                workSheet = workBook.Worksheets.Add()
                excel.ActiveWorkbook.SaveAs(r'{}'.format(xl_path))

            for par, col in params.items():
                workSheet.Cells[1, col] = par
            row = 2
            MEPSpaces = merge_two_dicts(New_MEPSpaces, Exist_MEPSpaces)
            for id, space in MEPSpaces.items():
                try:
                    workSheet.Cells[row, "A"] = id
                    for par in space.Parameters:
                        par_name = par.Definition.Name
                        if par_name in params.keys():
                            if par.StorageType == StorageType.String:
                                workSheet.Cells[row, params[par_name]] = par.AsString()
                            elif par.StorageType == StorageType.Integer:
                                workSheet.Cells[row, params[par_name]] = par.AsInteger()
                            else:
                                conv_val = UnitUtils.ConvertFromInternalUnits(par.AsDouble(), units.get(par.Definition.UnitType, DisplayUnitType.DUT_GENERAL))
                                workSheet.Cells[row, params[par_name]] = conv_val
                    row += 1
                    
                except:
                    pass

            # makes the Excel application visible to the user
            excel.Visible = True


        # open_dialog = FileOpenDialog("Revit Files: (*.rvt) | *.rvt")
        # open_dialog.Show()
        # model_path = open_dialog.GetSelectedModelPath()
        # linkedDoc = app.OpenDocumentFile(model_path, OpenOptions())
        # room_collector = FilteredElementCollector(linkedDoc)

        # Get AR RVT link and create a room collector instance
        # for linkedDoc in uiapp.Application.Documents:
        #	if linkedDoc.Title.Equals(link_name):
        #		room_collector = FilteredElementCollector(linkedDoc)

        # Create a Revit task dialog to communicate information to the user
        mainDialog = TaskDialog("Create spaces from rooms")
        mainDialog.MainInstruction = "Create spaces from rooms"
        mainDialog.MainContent = "Please pick a linked model to create spaces!"
        mainDialog.CommonButtons = TaskDialogCommonButtons.Ok
        mainDialog.DefaultButton = TaskDialogResult.Ok

        tResult = mainDialog.Show()

        # Create a selection
        sel = uidoc.Selection
        # Prompt the user to pick the required external link
        try:
            ref = sel.PickObject(ObjectType.Element, "Please pick a linked model instance")

            # Get AR RVT link
            rvt_link = doc.GetElement(ref.ElementId)
            linkedDoc = rvt_link.GetLinkDocument()

            # Create a room collector instance
            room_collector = FilteredElementCollector(linkedDoc)


            # Collect levels from the current document
            levels = FilteredElementCollector(doc).WhereElementIsNotElementType(
            ).OfCategory(BuiltInCategory.OST_Levels).ToElements()
            # For each level in the current model define its elevation and create level elevation:Level dictionary
            lvls_dict = {level.Elevation: level for level in levels}

            # Collect rooms from RVT link
            if room_collector and room_collector.GetElementCount() != 0:
                rooms_list = room_collector.WhereElementIsNotElementType(
                ).OfCategory(BuiltInCategory.OST_Rooms).ToElements()

            # Create a space collector instance
            _space_collector = FilteredElementCollector(
                doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_MEPSpaces)
            
            # Create spaces by rooms
            create_spaces_by_rooms(rooms_list)

        except:
            TaskDialog.Show("Error", "The operation was cancelled!")
        
        # Create a Revit task dialog to communicate information to the user
        mainDialog = TaskDialog("Write parameters to Excel")
        mainDialog.MainInstruction = "Write parameters to Excel"
        mainDialog.MainContent = "Do you want to write parameters to an Excell spreadsheet?"
        mainDialog.CommonButtons = TaskDialogCommonButtons.Yes|TaskDialogCommonButtons.No

        tResult = mainDialog.Show()
        if tResult == TaskDialogResult.Yes:
            # Write data to an excel file
            write_to_excel(excel_parameters)


    def Rooms_To_Spaces_v2(self):
        view = MainForm()
        model = Model(self)     
        controller = Controller(self, view=view, model=model)
        controller.run()


class Controller(object):
    def __init__(self, __revit__, view, model):
        self.app = __revit__.Application
        self.doc = __revit__.ActiveUIDocument.Document
        self._view = view
        self._model = model
        self.excel_parameters = []
        self.exl_file_dir_path = "C:\\"
        self.config_file_dir_path = os.path.dirname(os.path.realpath(__file__))
        self._open_xl_file_Dialog = WinForms.OpenFileDialog()
        self._open_config_file_Dialog = WinForms.OpenFileDialog()
        self._worker = BackgroundWorker()

        self._connectSignals()

    def _connectSignals(self):
        
        self._model.startProgress += self.startProgressBar
        self._model.ReportProgress += self.updateProgressBar
        self._model.endProgress += self.disableProgressBar
        self._worker.DoWork += lambda _, __: self._model.write_to_excel()
        self._worker.RunWorkerCompleted += self.run_worker_completed
        self._worker.WorkerSupportsCancellation = True
        self._view.Shown += self.On_MainForm_StartUp
        self._view.FormClosing += self.On_MainForm_Closing
        self._view._run_button.Click += self.run_button_Click
        self._view._browse_button.Click += self.browse_button_Click
        self._view._write_exl_button.Click += self.write_exl_button_Click
        self._view._del_spaces_button.Click += self.del_spaces_button_Click
        self._view._file_path_textBox.TextChanged += self.file_path_textBox_TextChanged
        self._view._shar_pars_check_all_btn.Click += self.shar_pars_check_all_btn_Click
        self._view._shar_pars_check_none_btn.Click += self.shar_pars_check_none_btn_Click
        self._view._builtin_pars_check_all_btn.Click += self.builtin_pars_check_all_btn_Click
        self._view._builtin_pars_check_none_btn.Click += self.builtin_pars_check_none_btn_Click
        self._view._space_id_comboBox.SelectedValueChanged += self.space_id_comboBox_SelectedValueChanged
        self._view._write_exl_checkBox.CheckedChanged += self.write_exl_checkBox_CheckChng
        self._view._load_stngs_button.Click += lambda _, __: self.load_settings((self._view._space_id_comboBox,
                                                                    self._view._file_path_textBox,
                                                                    self._view._shar_pars_checkedListBox,
                                                                    self._view._builtin_pars_checkedListBox))
        self._view._save_stngs_button.Click += lambda _, __: self.save_settings()

    def On_MainForm_StartUp(self, sender, args):
        self.load_project_parameters((self._view._space_id_comboBox, self._view._shar_pars_checkedListBox)) 
        self.load_builtin_parameters(self._view._builtin_pars_checkedListBox)

        if self._model.spaces_count > 0:
            self._view._del_spaces_button.Enabled = True

    def On_MainForm_Closing(self, sender, args):
        message = "Are you sure that you would like to close the form?"
        caption = "Form Closing"
        result = WinForms.MessageBox.Show(
            message, caption, WinForms.MessageBoxButtons.YesNo, WinForms.MessageBoxIcon.Question)
        # If the no button was pressed ...
        if result == WinForms.DialogResult.No:
            # cancel the closure of the form.
            args.Cancel = True
        
        elif result == WinForms.DialogResult.Yes:

            # Close Excel Application in Model class
            if self._model.excel is not None:
                self._model.excel.Quit()
                Marshal.FinalReleaseComObject(self._model.excel)

    def save_settings(self):

        try:
            filename = self._view._config_file_textBox.Text

            if self._view._config_file_textBox.Text == "":

                self._open_config_file_Dialog.InitialDirectory = self.config_file_dir_path
                self._open_config_file_Dialog.Filter = "JSON files (*.json)|*.json"
                self._open_config_file_Dialog.FilterIndex = 2
                self._open_config_file_Dialog.RestoreDirectory = True

                result = self._open_config_file_Dialog.ShowDialog(self._view)
                if result == WinForms.DialogResult.OK:
                    filename = self._open_config_file_Dialog.FileName
                    self._view._config_file_textBox.Text = filename
                else:
                    return

            config = {self._view._space_id_comboBox.Name: self._view._space_id_comboBox.SelectedItem,
                    self._view._file_path_textBox.Name: self._view._file_path_textBox.Text,
                    self._view._shar_pars_checkedListBox.Name: list(self._view._shar_pars_checkedListBox.CheckedItems),
                    self._view._builtin_pars_checkedListBox.Name: list(self._view._builtin_pars_checkedListBox.CheckedItems)}

            with io.open(filename,
                        'w', encoding='utf8') as file:
                json.dump(config, file, ensure_ascii=False, indent=4, sort_keys=True)

        except Exception as e:
            logger.error(e, exc_info=True)

    def load_settings(self, controls):

        self._open_config_file_Dialog.InitialDirectory = self.config_file_dir_path
        self._open_config_file_Dialog.Filter = "JSON files (*.json)|*.json"
        self._open_config_file_Dialog.FilterIndex = 2
        self._open_config_file_Dialog.RestoreDirectory = True

        result = self._open_config_file_Dialog.ShowDialog(self._view)
        if result == WinForms.DialogResult.OK:
            self._view._config_file_textBox.Clear()
            filename = self._open_config_file_Dialog.FileName
            self._view._config_file_textBox.Text = filename
        else:
            return

        try:
            with io.open(filename,
                        'r', encoding='utf8') as file:
                config = json.load(file)

                for c in controls:
                    if c.GetType() == clr.GetClrType(WinForms.ComboBox):
                        c.SelectedItem = config[c.Name]
                    elif c.GetType() == clr.GetClrType(WinForms.TextBox):
                        c.Text = config[c.Name]
                    elif c.GetType() == clr.GetClrType(WinForms.CheckedListBox):
                        for item in config[c.Name]:
                            index = c.FindStringExact(item)
                            if (index != c.NoMatches):
                                c.SetItemChecked(index, True)

        except Exception as e:
            logger.error(e, exc_info=True)

    def get_parameter_bindings(self):
        prj_defs_dict = {}
        binding_map = self.doc.ParameterBindings
        it = binding_map.ForwardIterator()
        it.Reset()
        while it.MoveNext():
            current_binding = it.Current
            if current_binding.Categories.Contains(Category.GetCategory(self.doc, BuiltInCategory.OST_MEPSpaces)):
                prj_defs_dict[it.Key.Name] = it.Key

        return prj_defs_dict

    def load_project_parameters(self, controls):
        internal_defs = self.get_parameter_bindings()
        for control in controls:
            control.Items.Clear()
            control.BeginUpdate()
            for par_name in sorted(internal_defs):
                if controls.index(control) == 0:
                    match = re.match(r'\w*(ROOM)?\w*(ID)', par_name, re.IGNORECASE)
                    if match:
                        control.Items.Add(par_name)
                else:
                    control.Items.Add(par_name)
            control.EndUpdate()
        controls[0].Items.Insert(0, "Please select a parameter...")
        controls[0].SelectedIndex = 0

    def shar_pars_check_all_btn_Click(self, sender, args):
        if self._view._shar_pars_checkedListBox.Items.Count > 0:
            i = 0
            while i < self._view._shar_pars_checkedListBox.Items.Count:
                self._view._shar_pars_checkedListBox.SetItemChecked(i, True)
                i += 1

    def shar_pars_check_none_btn_Click(self, sender, args):
        if self._view._shar_pars_checkedListBox.Items.Count > 0:
            i = 0
            while i < self._view._shar_pars_checkedListBox.Items.Count:
                self._view._shar_pars_checkedListBox.SetItemChecked(i, False)
                i += 1

    def builtin_pars_check_all_btn_Click(self, sender, args):
        if self._view._builtin_pars_checkedListBox.Items.Count > 0:
            i = 0
            while i < self._view._builtin_pars_checkedListBox.Items.Count:
                self._view._builtin_pars_checkedListBox.SetItemChecked(i, True)
                i += 1

    def builtin_pars_check_none_btn_Click(self, sender, args):
        if self._view._builtin_pars_checkedListBox.Items.Count > 0:
            i = 0
            while i < self._view._builtin_pars_checkedListBox.Items.Count:
                self._view._builtin_pars_checkedListBox.SetItemChecked(i, False)
                i += 1

    def get_builtin_parameters(self):
        builtin_pars = set()
        bicId = [ElementId(BuiltInCategory.OST_MEPSpaces)]
        catlist = List[ElementId](bicId)
        paramIds = ParameterFilterUtilities.GetFilterableParametersInCommon(self.doc, catlist)
        for bpar in BuiltInParameter.GetValues(clr.GetClrType(BuiltInParameter)):

            try:
                bparId = ElementId(bpar)
                if bparId in paramIds:
                    builtin_pars.add(LabelUtils.GetLabelFor(bpar))

            except Exception as e:
                logger.error(e, exc_info=True)
                pass

        return builtin_pars

    def load_builtin_parameters(self, control):
        builtin_pars = self.get_builtin_parameters()

        control.Items.Clear()
        control.BeginUpdate()
        for par_name in sorted(builtin_pars):
            control.Items.Add(par_name)
            control.EndUpdate()

    def space_id_comboBox_SelectedValueChanged(self, sender, args):
        if self._view._space_id_comboBox.SelectedIndex != 0:
            self._view._run_button.Enabled = True
            self._model.search_id = self._view._space_id_comboBox.SelectedItem

            if self._view._file_path_textBox.Text != "" and self._model.spaces_count > 0:
                self._view._write_exl_button.Enabled = True
                self._view._write_exl_checkBox.Enabled = True
                
        else:
            self._view._run_button.Enabled = False
            self._view._write_exl_button.Enabled = False
            self._view._write_exl_checkBox.Enabled = False

    def file_path_textBox_TextChanged(self, sender, args):
        if sender.Text != "" and self._view._space_id_comboBox.SelectedIndex != 0:
            self._view._write_exl_checkBox.Enabled = True
            if self._model.spaces_count > 0:
                self._view._write_exl_button.Enabled = True

    def write_exl_checkBox_CheckChng(self, sender, args):
        if sender.CheckState == WinForms.CheckState.Checked:
            self._model.xl_write_flag = True
        else:
            self._model.xl_write_flag = False

    def browse_button_Click(self, sender, args):
        self._open_xl_file_Dialog.InitialDirectory = self.exl_file_dir_path
        self._open_xl_file_Dialog.Filter = "CSV files (*.csv)|*.csv|Excel Files|*.xls;*.xlsx"
        self._open_xl_file_Dialog.FilterIndex = 2
        self._open_xl_file_Dialog.RestoreDirectory = True
        
        result = self._open_xl_file_Dialog.ShowDialog(self._view)
        if result == WinForms.DialogResult.OK:
            self._view._file_path_textBox.Clear()
            filename = self._open_xl_file_Dialog.FileName
            self._view._file_path_textBox.Text = filename

    def run_button_Click(self, sender, args):
        self.excel_parameters = sorted(list(self._view._shar_pars_checkedListBox.CheckedItems) + \
                                list(self._view._builtin_pars_checkedListBox.CheckedItems))
        if len(self.excel_parameters) == 0:
            WinForms.MessageBox.Show("At least one parameter must be chosen!", "Task cancelled!", 
            WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)
            return

        self.excel_parameters.insert(0, self._view._space_id_comboBox.SelectedItem)

        self._model.xl_file_path = self._view._file_path_textBox.Text
        self._model.excel_parameters = self.excel_parameters
        self._model.main()

    def write_exl_button_Click(self, sender, args):
        self.excel_parameters = sorted(list(self._view._shar_pars_checkedListBox.CheckedItems) + \
                                list(self._view._builtin_pars_checkedListBox.CheckedItems))
        if len(self.excel_parameters) == 0:
            WinForms.MessageBox.Show("At least one parameter must be chosen!", "Task cancelled!", 
            WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)
            return

        self.excel_parameters.insert(0, self._view._space_id_comboBox.SelectedItem)

        self._model.xl_file_path = self._view._file_path_textBox.Text
        self._model.excel_parameters = self.excel_parameters

        if not self._worker.IsBusy:
            self._worker.RunWorkerAsync()

    def del_spaces_button_Click(self, sender, args):
        self._model.delete_spaces()
        self._view._del_spaces_button.Enabled = False
        self._view._write_exl_button.Enabled = False

    def startProgressBar(self, *args):
        self._view._progressBar.Value = 0
        self._view._progressBar.Maximum = args[0]

    def updateProgressBar(self, *args):
        self._view._progressBar.Value = args[0]

    def disableProgressBar(self, *args):
        if self._view._progressBar.Maximum > self._view._progressBar.Value:
            self._view._progressBar.Value = self._view._progressBar.Maximum
            WinForms.MessageBox.Show("Task completed successfully!", "Success!", 
            WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)

            self._view._del_spaces_button.Enabled = True

            if self._view._file_path_textBox.Text != "":
                self._view._write_exl_button.Enabled = True

    def run_worker_completed(self, sender, args):
        
        if args.Cancelled:
            WinForms.MessageBox.Show("Task canceled!", "Canceled!", 
            WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)

        elif args.Error:
            WinForms.MessageBox.Show("Error in writing parameters!", "Error!", 
            WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error)
        
        else:
            self.disableProgressBar()

    def dispose(self):
        self._view.components.Dispose()
        WinForms.Form.Dispose(self._view)

    def run(self):
        '''
        Start our form object
        '''
        # Run the Application
        WinForms.Application.Run(self._view)


class Model(object):
    def __init__(self, __revit__, search_id=None, excel_parameters=None, xl_file_path=None, xl_write_flag=False):
        # region Get Document and Application
        self.doc = __revit__.ActiveUIDocument.Document
        self.uidoc = UIDocument(self.doc)
        self.app = __revit__.Application
        self.uiapp = UIApplication(self.app)
        self.excel = Excel.ApplicationClass()
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo("en-US")
        # endregion

        # region Initial parameters
        self._search_id = search_id
        self._excel_parameters = excel_parameters
        self._xl_file_path = xl_file_path
        self._xl_write_flag = xl_write_flag
        self.New_MEPSpaces = {}
        self.Exist_MEPSpaces = {}

        # endregion

        # region Custom Events
        self.startProgress = Event()
        self.ReportProgress = Event()
        self.endProgress = Event()

        # Create a space collector instance
        self._space_collector = FilteredElementCollector(
            self.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_MEPSpaces)

        # endregion

    # region Getters and Setters
    @property
    def search_id(self):
        return self._search_id

    @search_id.setter
    def search_id(self, value):
        self._search_id = value
    
    @property
    def xl_file_path(self):
        return self._xl_file_path

    @xl_file_path.setter
    def xl_file_path(self, value):
        self._xl_file_path = value
    
    @property
    def xl_write_flag(self):
        return self._xl_write_flag

    @xl_write_flag.setter
    def xl_write_flag(self, value):
        self._xl_write_flag = value
    
    @property
    def excel_parameters(self):
        return self._excel_parameters

    @excel_parameters.setter
    def excel_parameters(self, par_list):
        self._excel_parameters = par_list

    @property
    def spaces_count(self):
        self._spaces_count = self._space_collector.GetElementCount()
        return self._spaces_count
    # endregion Getters and Setters

    def main(self):

        # Create a selection
        sel = self.uidoc.Selection
        # Prompt the user to pick the required external link
        try:
            ref = sel.PickObject(ObjectType.Element, "Please pick a linked model instance")

            # Get AR RVT link
            rvt_link = self.doc.GetElement(ref.ElementId)
            linkedDoc = rvt_link.GetLinkDocument()

            # Create a room collector instance
            room_collector = FilteredElementCollector(linkedDoc)

            # Create a space collector instance
            self._space_collector = FilteredElementCollector(
                self.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_MEPSpaces)

            # Collect levels from the current document
            levels = FilteredElementCollector(self.doc).WhereElementIsNotElementType(
            ).OfCategory(BuiltInCategory.OST_Levels)

            # For each level in the current model define its elevation and create level elevation:Level dictionary
            self.lvls_dict = {level.Elevation: level for level in levels}
            
            # Collect rooms from RVT link
            if room_collector and room_collector.GetElementCount() != 0:
                self.rooms_list = room_collector.WhereElementIsNotElementType(
                ).OfCategory(BuiltInCategory.OST_Rooms)

            self.counter = 0
            self.startProgress.emit(room_collector.GetElementCount() + self.spaces_count + \
                (2*(room_collector.GetElementCount()*len(self._excel_parameters))))
            self.__main()
            self.endProgress.emit()

        except Exception as e:
            logger.error(e, exc_info=True)
            WinForms.MessageBox.Show("The operation was cancelled!", "Error!",
            WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)
            return

    def __main(self):
        try:
            # Create spaces by rooms
            self.create_spaces_by_rooms(self.rooms_list)
            
            # Write data to an excel file
            if self._xl_write_flag:
                self.__write_to_excel(self._excel_parameters)

        #except Exception as e:
        #    msg = traceback.format_exc(str(e))
        #    WinForms.MessageBox.Show(msg, "Error!",
        #    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)
        #    return

        except Exception as e:
            logger.error(e, exc_info=True)
            pass
        
    def create_spaces_by_rooms(self, rooms):
        # Initiate the transacton group
        with TransactionGroup(self.doc, "Batch create spaces/Transfer parameters") as tg:
            tg.Start()
            Room_UniqueIds = [room.UniqueId for room in rooms if room.Area > 0 and room.Location != None]
            self.counter += 1
            self.ReportProgress.emit(self.counter)
            # Define if there're spaces in a model
            if self._space_collector.GetElementCount() == 0:
                # If there are no spaces
                # Create a space
                for room in rooms:
                    if room.Area > 0 and room.UniqueId not in self.New_MEPSpaces:
                        self.space_creator(room, self.lvls_dict)

                        self.counter += 1
                        self.ReportProgress.emit(self.counter)

            # If there are spaces in the model
            else:
                self.space_check(Room_UniqueIds)
                # For each room in RVT link rooms take room UniqueId and check if there's space with the same Id among the existing MEP spaces
                for room in rooms:
                    # If there's such space get it
                    if room.UniqueId in self.Exist_MEPSpaces:
                        exst_space = self.Exist_MEPSpaces[room.UniqueId]
                        self.space_updater(room, exst_space)
                        
                        self.counter += 1
                        self.ReportProgress.emit(self.counter)
                        
                    # If there's no such space
                    else:
                        # Create a space
                        if room.Area > 0 and room.UniqueId not in self.New_MEPSpaces:
                            self.space_creator(room, self.lvls_dict)

                        self.counter += 1
                        self.ReportProgress.emit(self.counter)

            tg.Assimilate()
    
    def space_creator(self, room, lvls):
        ''' 
        Function creates new spaces 
        room: Revit Room, lvls: level elevations and levels dictionary
        '''
        try:
            # Get the room level
            point = room.Location.Point
            room_lvl = room.Level
            # Get a level from the lvls dictionary by room level elevation
            lvl = lvls.get(room_lvl.Elevation)
            # Create space by coordinates and level taken from room
            with Transaction(self.doc, "Batch create spaces") as tr:
                tr.Start()
                
                options = tr.GetFailureHandlingOptions()
                failureHandler = ErrorSwallower()
                options.SetFailuresPreprocessor(failureHandler)
                options.SetClearAfterRollback(True)
                tr.SetFailureHandlingOptions(options)
                space = self.doc.Create.NewSpace(lvl, UV(point.X, point.Y))
                
                # Get "REFERENCED_ROOM_UNIQUE_ID" parameter
                ref_id_par = space.GetParameters(self._search_id)[0]
                # Assign room UniqueID to "REFERENCED_ROOM_UNIQUE_ID" parameter
                if ref_id_par:
                    if ref_id_par.StorageType == StorageType.String:
                        ref_id_par.Set(room.UniqueId)
                        self.New_MEPSpaces[room.UniqueId] = self.New_MEPSpaces.get(room.UniqueId, space)
                        self.para_setter(room, space)
                    else:
                        ref_id_par.Set(room.Id.IntegerValue)
                        self.New_MEPSpaces[room.Id.IntegerValue] = self.New_MEPSpaces.get(room.Id.IntegerValue, space)
                        self.para_setter(room, space)
                tr.Commit()

        except Exception as e:
            logger.error(e, exc_info=True)
            pass 

        try:
            with Transaction(self.doc, "Set space Id") as tr:
                tr.Start()   
                # Get "ID_revit" parameter of a MEP space
                space_id = space.GetParameters("ID_revit")[0]
                # Assign space ID to "ID_revit" parameter
                if space_id:
                    space_id.Set(space.Id.IntegerValue)
                tr.Commit()

        except Exception as e:
            logger.error(e, exc_info=True)
            pass

    def space_updater(self, room, exst_space):
        ''' 
        Function updates existing spaces and moves them if necessary
        '''
        # And extract its coordinates
        try:
            exst_space_coords = exst_space.Location.Point
            # Get room coordinates
            room_coords = room.Location.Point
            # Compare two sets of coordinates
            # If they are almost the same
            if exst_space_coords.IsAlmostEqualTo(room_coords):
                with Transaction(self.doc, "Update parameters") as tr:
                    tr.Start()
                    options = tr.GetFailureHandlingOptions()
                    failureHandler = ErrorSwallower()
                    options.SetFailuresPreprocessor(failureHandler)
                    options.SetClearAfterRollback(True)
                    tr.SetFailureHandlingOptions(options)
                    # Transfer room parameters to a corresponding space
                    self.para_setter(room, exst_space)
                    tr.Commit()
            # Otherwise, move the existing space according to room coordinates
            else:
                move_vector = room.Location.Point - exst_space.Location.Point
                with Transaction(self.doc, "Move spaces") as tr:
                    tr.Start()
                    options = tr.GetFailureHandlingOptions()
                    failureHandler = ErrorSwallower()
                    options.SetFailuresPreprocessor(failureHandler)
                    options.SetClearAfterRollback(True)
                    tr.SetFailureHandlingOptions(options)
                    exst_space.Location.Move(move_vector)
                    # Transfer room parameters to a corresponding space
                    self.para_setter(room, exst_space)
                    tr.Commit()
        except System.MissingMemberException as e:
            logger.error(e, exc_info=True)
            pass
    
    def space_check(self, Room_UniqueIds):
        ''' 
        Function checks for not placed spaces, obsolete spaces and deletes them
        '''
        # Collect all existing MEP spaces in this document
        for space in self._space_collector.ToElements():

            # Get "REFERENCED_ROOM_UNIQUE_ID" parameter of each space
            id_par_list = space.GetParameters(self._search_id)
            ref_id_par_val = id_par_list[0].AsString()

            # Check if REFERENCED_ROOM_UNIQUE_ID is in room Room_UniqueIds
            # If it's not the case delete the corresponding space
            if space.Area == 0 or ref_id_par_val not in Room_UniqueIds:
                with Transaction(self.doc, "Delete spaces") as tr:
                    tr.Start()
                    
                    try:
                        self.doc.Delete(space.Id)
                    except Exception as e:
                        logger.error(e, exc_info=True)
                        pass

                    tr.Commit()

            else:
                # Otherwise cast it into Existing MEP spaces dictionary
                self.Exist_MEPSpaces[ref_id_par_val] = self.Exist_MEPSpaces.get(
                    ref_id_par_val, space)
            
            self.counter += 1
            self.ReportProgress.emit(self.counter)

    def para_setter(self, room, space):
        ''' 
        Function transers parameters from the room to newly created spaces
        room: Revit Room, space: MEPSpace
        '''
        # For each parameter in room parameters define if it's shared or builtin parameter
        for par in room.Parameters:
            par_name = par.Definition.Name
            if par.IsShared and par_name in self._excel_parameters:
                # If room parameter is shared get space from Spaces dictionary by UniqueId and extract corresponding space parameter from it by room parameter GUID
                # Depending on the room parameter storage type set its value to the space parameter
                if not par.IsReadOnly:
                    try:
                        space_par = space.get_Parameter(par.GUID)
                        if par.StorageType == StorageType.String and par.HasValue:
                            space_par.Set(par.AsString())
                        elif par.StorageType == StorageType.Integer and par.HasValue:
                            space_par.Set(par.AsInteger())
                        elif par.StorageType == StorageType.Double and par.HasValue:
                            space_par.Set(par.AsDouble())
                    except Exception as e:
                        logger.error(e, exc_info=True)
                        pass

                    self.counter += 1
                    self.ReportProgress.emit(self.counter)

            elif par.Definition.BuiltInParameter != BuiltInParameter.INVALID\
                and LabelUtils.GetLabelFor(par.Definition.BuiltInParameter) in self._excel_parameters:
                # If room parameter is builtin get space from Spaces dictionary by UniqueId and extract corresponding space parameter from it by builtin parameter
                # Depending on the room parameter storage type set its value to the space parameter
                if not par.IsReadOnly:
                    try:
                        space_par = space.get_Parameter(par.Definition)
                        if par.StorageType == StorageType.String and par.HasValue:
                            space_par.Set(par.AsString())
                        elif par.StorageType == StorageType.Integer and par.HasValue:
                            space_par.Set(par.AsInteger())
                        elif par.StorageType == StorageType.Double and par.HasValue:
                            space_par.Set(par.AsDouble())

                    except Exception as e:
                        logger.error(e, exc_info=True)
                        pass

                    self.counter += 1
                    self.ReportProgress.emit(self.counter)
    
    def delete_spaces(self):
        ''' 
        Function deletes spaces
        '''
        self.counter = 0

        space_collector = FilteredElementCollector(
            self.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_MEPSpaces)
        spaces = space_collector.ToElements()
        self.startProgress.emit(len(spaces))

        with Transaction(self.doc, "Delete spaces") as tr:
            tr.Start()

            for space in spaces:

                try:
                    self.doc.Delete(space.Id)
                    self.counter += 1
                    self.ReportProgress.emit(self.counter)

                except Exception as e:
                    logger.error(e, exc_info=True)
                    pass

            tr.Commit()
            
            self.New_MEPSpaces = {}
            self.Exist_MEPSpaces = {}
            self.endProgress.emit()

    def merge_two_dicts(self, d1, d2):
        '''
        Merges two dictionaries
        Returns a merged dictionary
        '''
        d_merged = d1.copy()
        d_merged.update(d2)
        return d_merged

    def write_to_excel(self):
        """
        Writing parameters to an Excel workbook
        """
        self.counter = 0
        self.startProgress.emit(self._space_collector.GetElementCount() + \
                ((self._space_collector.GetElementCount()*len(self._excel_parameters))))
            
        # Obtaining existing spaces
        self.__get_exist_spaces()
        # Write data to an excel file
        self.__write_to_excel(self._excel_parameters)

    def __get_exist_spaces(self):
        """
        Obtains existing spaces
        Fills the Exist_MEPSpaces dictionary: key: reference_room_id, value: MEP space
        """
        self.Exist_MEPSpaces = {}
        # Collect all existing MEP spaces in this document
        for space in self._space_collector:

            try:
                # Get "REFERENCED_ROOM_UNIQUE_ID" parameter of each space
                id_par_list = space.GetParameters(self._search_id)
                ref_id_par_val = id_par_list[0].AsString()

                # Cast space into Existing MEP spaces dictionary
                self.Exist_MEPSpaces[ref_id_par_val] = self.Exist_MEPSpaces.get(
                        ref_id_par_val, space)

                self.counter += 1
                self.ReportProgress.emit(self.counter)

            except Exception as e:
                logger.error(e, exc_info=True)
                pass
            
    def __write_to_excel(self, params_to_write):
        """
        Writing parameters to an Excel workbook
        (private method)
        """
        units = {UnitType.UT_Length: DisplayUnitType.DUT_METERS,
                UnitType.UT_Area: DisplayUnitType.DUT_SQUARE_METERS,
                UnitType.UT_Volume: DisplayUnitType.DUT_CUBIC_METERS,
                UnitType.UT_Number: DisplayUnitType.DUT_PERCENTAGE,
                UnitType.UT_HVAC_Heating_Load: DisplayUnitType.DUT_KILOWATTS,
                UnitType.UT_HVAC_Cooling_Load: DisplayUnitType.DUT_KILOWATTS,
                UnitType.UT_HVAC_Airflow_Density: DisplayUnitType.DUT_CUBIC_FEET_PER_MINUTE_SQUARE_FOOT,
                UnitType.UT_HVAC_Airflow: DisplayUnitType.DUT_CUBIC_METERS_PER_HOUR}
        
        columns = list(string.ascii_uppercase) + ["".join(i) for i in product(string.ascii_uppercase, repeat=2)]
        params = dict(izip(params_to_write, columns))

        try:
            workBook = self.excel.Workbooks.Open(r'{}'.format(self._xl_file_path))
            workSheet = workBook.Worksheets(1)
        
        except Exception as e:
            logger.error(e, exc_info=True)
            pass

        for par, col in params.items():
            workSheet.Cells[1, col] = par
        row = 2

        MEPSpaces = self.merge_two_dicts(self.New_MEPSpaces, self.Exist_MEPSpaces)

        for space_id, space in MEPSpaces.items():
            try:
                workSheet.Cells[row, "A"] = space_id
                for par in space.Parameters:
                    par_name = par.Definition.Name
                    if par_name in params.keys():
                        if par.StorageType == StorageType.String:
                            workSheet.Cells[row, params[par_name]] = par.AsString()
                        elif par.StorageType == StorageType.Integer:
                            workSheet.Cells[row, params[par_name]] = par.AsInteger()
                        else:
                            conv_val = UnitUtils.ConvertFromInternalUnits(par.AsDouble(), units.get(par.Definition.UnitType, DisplayUnitType.DUT_GENERAL))
                            workSheet.Cells[row, params[par_name]] = conv_val
                        self.counter += 1
                        self.ReportProgress.emit(self.counter)
                row += 1
                    
            except Exception as e:
                logger.error(e, exc_info=True)
                pass

        # makes the Excel application visible to the user
        self.excel.Visible = True


class Event(object):
    def __init__(self):
        self.handlers = []

    def __iadd__(self, handler):
        self.handlers.append(handler)
        return self

    def __isub__(self, handler):
        self.handlers.remove(handler)
        return self

    def emit(self, *args):
        for handler in self.handlers:
            handler(*args)


class ErrorSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        failureMessages = failuresAccessor.GetFailureMessages()

        for fma in failureMessages.GetEnumerator():
            
            try:
                failureSeverity = fma.GetSeverity()

                if failureSeverity == FailureSeverity.Warning:
                    failuresAccessor.DeleteWarning(fma)
                
                elif fma.HasResolutionOfType(FailureResolutionType.FixElements):
                    fma.SetCurrentResolutionType(FailureResolutionType.FixElements)
                    failuresAccessor.ResolveFailure(fma)
                    return FailureProcessingResult.ProceedWithCommit
                
                else:
                    return FailureProcessingResult.ProceedWithRollBack
            except Exception as e:
                logger.error(e, exc_info=True)
                pass
        
        return FailureProcessingResult.Continue


class DuplicateSpaceWarningSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        failureMessages = failuresAccessor.GetFailureMessages()

        for fma in failureMessages.GetEnumerator():
            failure_id = fma.GetFailureDefinitionId()
            
            try:
                failureSeverity = fma.GetSeverity()
                if failureSeverity == FailureSeverity.Warning and failure_id == BuiltInFailures.RoomFailures.RoomsInSameRegionSpaces:
                    failuresAccessor.DeleteWarning(fma)
            except Exception as e:
                logger.error(e, exc_info=True)
                pass

        return FailureProcessingResult.Continue
