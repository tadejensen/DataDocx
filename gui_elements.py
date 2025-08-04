import tkinter as tk
from tkinter import ttk, font
from tkinter import filedialog as fd
import tkinterDnD as dnd

import pandas as pd
import numpy as np

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from ipydex import IPS
import os
import sys
import traceback
import platform
import subprocess
import datetime

from lengthy_imports import *
import gui_functions as gui_f
import physicals as phys
import database_functions as dbf

from typing import Optional

idx = pd.IndexSlice
config = load_config()

# TODO: possibility to copy remarks from one turbine to another, like when defining which sections are in Fazit or when determining which points are in Prüfliste (copy to turbine, replace remark [text only/delete images?])
# TODO: add support for RBLi, RBLa, Drone (checklist multiplication)
# TODO: Gitbutton: check if origin is ahead and (in case it is) pull before pushing (maybe always push/pull, so that it's one sync button?)
# TODO: put all data into database and use .get() methods in the classes
# TODO: use symbols on Buttons
# TODO: sidebar for all extra functions (alle anlagen im Park anzeigen, daten synchronisieren, alle titel aus order anzeigen)
# TODO: button to undo remark deletion
# TODO: find and eliminate redundant function calls (e.g. 7 times load_turbines whatever you do lol)

# IDEA: make remarks / titles clickable, so that a menu pops out below. here you can select the actions (add, delete, modifiy, copy text, ...?, if modify: open new window containing only tk.Text and ok Button)
# IDEA: add alles i.O. Button that automatically puts the turbine's usual alles i.O. Text and marks the chapter as done, also consider copy [WEA ID] Button
# IDEA: reduce remarks clutter/complications on merge by splitting up the database into multiple files

# BUG: when in RemarkEditor, changing page of Remarkselector and then changing back to original page, then clicking okay, the page will not be refreshed
# BUG: when changing the width of an entry to the checklist (text length, adding flag), the size of the remarkeditor window can flicker. this happens when the width of the checklist_frm needs to change. this hangs up the program which then has to be killed
# BUG: after placing a pagebreak, a new paragraph is started. this can lead to an empty page, if the pagebreak is inserted where there would be a pagebreak anyway because of a full page (pagebreak sits on empty page)
# BUG: catch no weas selected selected on setup / no subfolders for weas


class Mainwindow(dnd.Tk):
    def __init__(self, project=None):
        super().__init__()
        combo = ttk.Combobox()
        combo.unbind_class('TCombobox', '<MouseWheel>')
        combo.unbind_class('TCombobox', '<ButtonPress-4>')
        combo.unbind_class('TCombobox', '<ButtonPress-5>')

        self.minsize(500, 700)
        self.project = project
        self.title('DataDocx - kein Projekt ausgewählt')
        self.iconbitmap(f'{mainpath}/images/icon.ico')

        self.lift()
        self.focus_force()

        self.allframe = ttk.Frame(self)
        self.allframe.pack(fill='both', expand=True)

        self.headline_frm = ttk.Frame(self.allframe)
        self.headline_frm.grid_columnconfigure(1, weight=100)
        self.parkname_frm = ttk.Frame(self.headline_frm)
        self.headlinetext = tk.StringVar(self)
        self.headlinetext.set('Kein Projekt ausgewählt.\nErstelle einen Projektordner mit einem Unterornder pro WEA.\nStarte DataDocx neu.')
        headline = MultilineLabel(self.parkname_frm, textvar=self.headlinetext)
        headline.grid(row=0, column=0, sticky='ew', padx=1, pady=1)
        self.parkname_frm.grid(row=0, column=0, padx=1, pady=1, sticky='w')
        self.headline_frm.grid(row=0, column=0, sticky='ew')
        self.show_park_var = tk.IntVar(self, value=0)

        self.workframe = ttk.Frame(self.allframe)
        self.workframe.grid(row=1, column=0, sticky='nswe')
        self.weawidget = None
        self.allframe.grid_rowconfigure(1, weight=100)
        self.allframe.grid_columnconfigure(0, weight=100)

        self.weaselscrl = ScrollFrame(self.headline_frm, orient='horizontal', def_height=42)
        self.weaselfrm = self.weaselscrl.viewPort
        if self.project is not None:
            self.weaselscrl.grid(row=0, column=1, sticky='ew', padx=1, pady=1)
        
        self.wea_selectors = {}
        self.wea = None

        self.init_databases()

        if self.project is not None:
            self.set_project(self.project)

        self.protocol('WM_DELETE_WINDOW', self.close)

    def close(self):
        self.destroy()
        sys.exit()

    def init_databases(self):
        if 'databases' not in os.listdir(mainpath):
            os.mkdir(f'{mainpath}/databases')
        dbf.init_databases()

        if 'report' not in os.listdir(f'{mainpath}/databases'):
            os.mkdir(f'{mainpath}/databases/report')
        dbf.init_config()

    def set_project(self, project):
        self.project = project
        self.weaselscrl.grid(row=0, column=1, sticky='ew', padx=1, pady=1)
        self.title(f'DataDocx: {self.project.name}')
        self.headlinetext.set(f'{self.project.name}')
        active_wea, _ = self.project.get('active_page')
        self.park_cbx = ttk.Checkbutton(self.parkname_frm, variable=self.show_park_var,
                                        text='Alle zeigen')
        self.park_cbx.grid(row=1, column=0, sticky='n', padx=1, pady=1)

        weas = self.project.windfarm.get_weas()
        all_rems = self.project.get_all_remarks()
        for i, wea in enumerate(weas):
            wea_id = wea.get('id')
            try: remarks = dbf.filter_specific_report(all_rems, wea)
            except KeyError: remarks = dbf.get_empty_remarks_df()
            self.weaselfrm.grid_columnconfigure(i, weight=1)
            self.place_wea_selector(wea_id, i, active=False, remarks=remarks)
        
        self.wea = self.project.windfarm.get_wea(active_wea if active_wea\
                                                 else weas[0].get('id'))
        try: remarks = dbf.filter_specific_report(all_rems, self.wea)
        except: remarks = dbf.get_empty_remarks_df()
        self.update_weapage(self.wea.get('id'),
                            remarks=remarks if not remarks.empty else None)
            

    def update_weapage(self, wea_id, remarks=None):
        '''updates/changes the mainwindow's current turbine
        given remarks are only considered if the wea has not changed'''
        wea_changed = self.wea.get('id') != wea_id
        if wea_changed:
            self.update_wea_selector(self.wea.get('id'), active=False)
            self.project.set_active_wea(wea_id)
            self.wea = self.project.windfarm.get_wea(wea_id)
        else:
            self.update_wea_selector(self.wea.get('id'), active=True,
                                     remarks=remarks)
        _, active_chapter=self.project.get('active_page')
        if not active_chapter and type(self.weawidget) == TurbineViewer:
            active_chapter = self.weawidget.get_active_section()
        gui_f.delete_children(self.workframe)
        if self.wea.is_setup:
            self.weawidget = TurbineViewer(self.workframe, self.project, wea_id,
                                            active_section=active_chapter,
                                            parent_mainwindow=self,
                                            remarks=remarks if not wea_changed else None)
        else:
            self.weawidget = WeaEditor(self.workframe, self.project, wea_id, self)
        
        self.weawidget.pack(side='top', padx=2, pady=2, expand=True, fill='both')
        if wea_changed:
            self.update_wea_selector(self.wea.id, active=True)
        self.focus()

    def update_wea_selector(self, wea_id, active=False, remarks=None):
        wea_sel, i = self.wea_selectors[wea_id]
        wea_sel.update(remarks=remarks)
        self.place_wea_selector(wea_id, i, active, remarks=remarks)

    def place_wea_selector(self, wea_id, i, active, remarks=None):
        wea_sel = Single_Wea_Selector(self, wea_id, active=active, remarks=remarks)
        wea_sel.btn.grid(row=0, column=i, sticky='ew', padx=3, pady=1)
        wea_sel.progbarfrm.grid(row=1, column=i, sticky='ew', padx=3, pady=1)
        self.wea_selectors[wea_id] = (wea_sel, i)

class Single_Wea_Selector():
    '''helper class to change mainwindow's current turbine'''
    def __init__(self, parent_mainwindow, wea_id, active=False, remarks=None) -> None:
        self.parent_mainwindow = parent_mainwindow
        self.wea_id = wea_id
        self.wea = self.parent_mainwindow.project.windfarm.get_wea(self.wea_id)
        self.active = active

        self.btn = ttk.Button(self.parent_mainwindow.weaselfrm, text=wea_id,
                              command=self.change_wea)
        self.progbarfrm = ttk.Frame(self.parent_mainwindow.weaselfrm)
        self.progbarfrm.grid_rowconfigure(0, weight=100)
        
        self.update(remarks=remarks)
            
    def change_wea(self, remarks=None):
        if self.parent_mainwindow.wea.get('id') == self.wea_id:
            remarks = self.parent_mainwindow.wea.report.get_remarks(ordered=False) # ordering done in remarkSelector
        self.parent_mainwindow.update_weapage(self.wea_id, remarks=remarks)

    def update(self, remarks=None):
        self.update_progbar(remarks=remarks)
    
    def get_chapter_tup(self, remarks=None):
        '''returns tuple like (done_chaps, all_chaps). if wea is not set up,
        all_chaps = 0.'''
        if not self.wea.is_setup:
            return (0, 0)

        report = self.wea.report
        all_chapters = report.get_chapters(remarks=remarks)
        done_chapters = report.get_done_chapters()

        return (len(done_chapters), len(all_chapters))

    def update_progbar(self, remarks=None):
        gui_f.delete_children(self.progbarfrm)
        done_chaps, all_chaps = self.get_chapter_tup(remarks=remarks)

        progbarheight = 4
        green = 'green2' if self.active else 'pale green'
        gray = 'black' if self.active else 'gray54'
        white = 'orchid2' if self.active else 'white smoke'

        if all_chaps == 0 or done_chaps == all_chaps:
            col = gray if all_chaps == 0 else green
            frm = tk.Frame(self.progbarfrm, background=col, height=progbarheight)
            self.progbarfrm.grid_columnconfigure(0, weight=1)
            frm.grid(row=0, column=0, sticky='nsew', padx=0, pady=0)
            return

        greenfrm = tk.Frame(self.progbarfrm, background=green, height=progbarheight)
        whitefrm = tk.Frame(self.progbarfrm, background=white, height=progbarheight)

        self.progbarfrm.grid_columnconfigure(0, weight=done_chaps)
        self.progbarfrm.grid_columnconfigure(1, weight=all_chaps-done_chaps)

        greenfrm.grid(row=0, column=0, sticky='nsew', padx=0, pady=0)
        whitefrm.grid(row=0, column=1, sticky='nsew', padx=0, pady=0)            

  
class Project_Setup_Window(tk.Toplevel):
    '''welche weas,
       welche art von prüfung(en),
       Jahr der prüfung, windpark, standort
    '''
    def __init__(self, parent, project: Optional[phys.Project]=None):
        super().__init__(parent, width=500, height=800)
        self.focus()
        self.bind('<Escape>', lambda e: self.destroy())
        self.bind('<Control-Return>', lambda e: self.save())
        self.title('Projekt konfigurieren')
        self.parent=parent
        self.project = project
        self.grid_columnconfigure(0, weight=100)
        self.grid_rowconfigure(0, weight=100)
        self.minsize(400, 700)

        allfrm = ttk.Frame(self)
        allfrm.grid(row=0, column=0, sticky='nsew')
        allfrm.grid_columnconfigure(0, weight=100)
        allfrm.grid_rowconfigure(0, weight=100)
        scrl = ScrollFrame(allfrm)
        scrl.grid(column=0, row=0, sticky='nswe', padx=1, pady=1)
        self.workfrm = scrl.viewPort
        self.workfrm.grid_columnconfigure(1, weight=100)

        i = 0
        # make list out of which used wea can be checked / unchecked
        ttk.Label(self.workfrm, text='WEAs auswählen')\
            .grid(row=i, column=0, sticky='w', padx=1, pady=1)
        i += 1
        allweas = gui_f.get_subfolder_names(os.getcwd())
        self.wea_boxes = Checkboxes_Group(self.workfrm, options=allweas,
                                          select_all_box=True,
                                          default_state=1)
        for wea_id in self.wea_boxes.keys():
            box, _ = self.wea_boxes[wea_id]
            box.grid(row=i, column=0, sticky='w', padx=1, pady=1, columnspan=3)
            i += 1
        ttk.Separator(self.workfrm)\
            .grid(row=i, column=0, columnspan=3, sticky='ew', padx=5, pady=1)
        i += 1


        # put label / input pair for every property in farm_info and turbine info
        projdata = list(project_properties.keys())

        self.projdata = {}

        immutable = ['Name', 'Jahr/ID'] if self.project else []
        proj_EF = EntryFrame(parent=self.workfrm,
                             properties_clearname=projdata,
                             prefilled_object=self.project if self.project\
                                else {'Auftragnehmer': config.default_contractor,
                                      'Verantwortl. Ingenieur': config.default_engineer},
                             clearname2argname=project_properties,
                             autoplace_start_row=i,
                             multiline=['Auftragnehmer', 'Unterauftragnehmer'],
                             immutable=immutable,
                             prefill_options={'Name': self.prefill_name,
                                              'Jahr/ID': self.prefill_yearid,
                                              'Auftragsdatum': self.prefill_date},
                             height=5)
        for clearname in projdata:
            strvar = proj_EF.get_textvars()[clearname]
            if clearname == 'Name':
                self.prefill_name_btn = proj_EF.get_entrylines()[clearname].get_button()
                strvar.trace_add('write', self.toggle_prefill_date_btn)
            if clearname == 'Jahr/ID':
                self.prefill_yearid_btn = proj_EF.get_entrylines()[clearname].get_button()
            if clearname == 'Auftragsdatum':
                self.date_prefill_btn = proj_EF.get_entrylines()[clearname].get_button()
                strvar.trace_add('write', self.toggle_prefill_yearid_btn)
            if clearname in immutable:
                strvar.trace_add('write', self.toggle_okay_button)
            self.projdata[clearname] = (strvar, project_properties[clearname])
        if self.project:
            for clearname in projdata[:2]:
                if btn := proj_EF.get_entrylines()[clearname].get_button():
                    btn.configure(state='disabled')
        i = proj_EF.get_next_free_row()

        ttk.Separator(self.workfrm)\
            .grid(row=i, column=0, columnspan=3, sticky='ew', padx=5, pady=1)
        i += 1
        ####### bis ichg gegangen bin war ich dabei, den part hier unten mit einem
        # entryframe zu ersetzen, die entryline zu löschen und die ganzen neuerungen zu testen
        self.inspectiondata = {}
        kw = 'kind'
        clearname = 'Prüfauftrag'

        if self.project: prefill = self.project.get_inspection_kind()
        else: prefill = ''
        entryline = Entryline(
                        self.workfrm, clearname, prefill,
                        immutable=True if self.project else False,
                        prefill_options=list(inspection_type_translator.keys()),
                        state='diabled' if self.project else 'readonly')
        entryline.get_label().grid(row=i, column=0, sticky='w', padx=1, pady=1)
        entryline.get_entry().grid(row=i, column=1, sticky='ew', padx=1, pady=1)
        inspkind_var = entryline.get_textvar()
        inspkind_var.trace_add('write', self.toggle_okay_button)
        
        self.inspectiondata[clearname] = (inspkind_var, inspection_properties[clearname])
        i += 1

        # add OK Button
        controls_frm = ttk.Frame(allfrm)
        controls_frm.grid(row=1, column=0, sticky='ew', padx=1, pady=1)
        
        self.okay_btn = ttk.Button(controls_frm,
                                   text='Aktualisieren' if self.project else 'Erstellen',
                                   command=self.save)
        self.okay_btn.grid(padx=1, pady=1, sticky='e', column=1, row=0)
        self.toggle_okay_button()
        self.toggle_prefill_date_btn()
        self.toggle_prefill_yearid_btn()
        self.toggle_prefill_name_btn()

    def toggle_okay_button(self, *_):
        for d in [self.projdata,
                  self.inspectiondata]:        
            for kw in d.keys():
                if kw not in ['name', 'year_id', 'kind']: continue
                strvar = d[kw][0]
                if not strvar.get():
                    self.okay_btn.configure(state='disabled')
                    return
        self.okay_btn.configure(state='normal')

    def toggle_prefill_name_btn(self):
        if self.project: self.prefill_name_btn.configure(state='disabled')
        else: self.prefill_name_btn.configure(state='normal')
    def prefill_name(self):
        idvar = self.projdata['Name'][0]
        basename = os.path.basename(os.path.normpath(os.getcwd()))
        idvar.set(basename[6:].strip())
    def prefill_date(self):
        datevar = self.projdata['Auftragsdatum'][0]
        datevar.set(self.get_date_from_foldername())
    def toggle_prefill_date_btn(self, *_):
        if self.get_date_from_foldername(): self.date_prefill_btn.config(state='normal')
        else: self.date_prefill_btn.config(state='disabled')
    def get_date_from_foldername(self):
        '''foldername must start with yymmdd'''
        basename = os.path.basename(os.path.normpath(os.getcwd()))
        datestr = basename[:6]
        try: date = pd.to_datetime(datestr, format='%y%m%d')
        except ValueError: return None
        return f'{str(date.day).zfill(2)}.{str(date.month).zfill(2)}.{date.year}'
    def prefill_yearid(self):
        strvar = self.projdata['Jahr/ID'][0]
        strvar.set(self.get_yearid())
    def get_yearid(self):
        date = self.projdata['Auftragsdatum'][0].get()
        if not date: date = self.get_date_from_foldername()
        if not date: return None
        try:
            year = str(pd.to_datetime(date, dayfirst=True).year)
            if year == 'NaT': return None
            return year
        except ValueError: return None
    def toggle_prefill_yearid_btn(self, *_):
        year = self.get_yearid()
        if not year or self.project: state = 'disabled'
        else: state = 'normal'
        self.prefill_yearid_btn.configure(state=state)
     
    def create_project(self):
        proj_kws = self.get_kws(which=['project'])

        windfarm = phys.Windfarm()
        self.project = phys.Project(windfarm=windfarm, **proj_kws)

        repo_kws = self.get_kws(which='report')
        
        wea_ids = self.wea_boxes.get_selected()
        for wea_id in wea_ids:
            new_wea = phys.Turbine(id=wea_id,
                                   report_kwargs=repo_kws,
                                   **{'oem': 'non_setup'})
            windfarm.add_wea(new_wea)

        self.parent.set_project(self.project)
        self.project.create_notes()
        self.project.save()
        self.destroy()

    def update_project(self):
        '''updates prject data according to the entries of the setup window
        NOTE: does not change turbines whose "is_setup" is set to True!
        same applies for these turbines' report and inspection
        '''
        proj_kws, _, repo_kws = self.get_kws()
        new_wea_ids = self.wea_boxes.get_selected()

        # if turbine is missing or not set up yet: create turbine anew
        # with currently given information
        for wea_id in new_wea_ids:
            curr_wea = self.project.windfarm.get_wea(wea_id)
            new_wea = phys.Turbine(id=wea_id,
                                   report_kwargs=repo_kws,
                                   **{'oem': 'non_setup'})
            if curr_wea is None:
                self.project.windfarm.add_wea(new_wea)

            elif not curr_wea.is_setup:
                self.project.windfarm.remove_wea(wea_id)
                self.project.windfarm.add_wea(new_wea)
            # else: do nothing, leave the wea, report and inspection untouched

        # if turbine checkbox has been removed: delete from project
        current_wea_ids = self.project.windfarm.get_wea_ids()
        for wea_id in current_wea_ids:
            if wea_id not in new_wea_ids:
                self.project.windfarm.remove_wea(wea_id)
        
        # change project data
        gui_f.change_attributes_from_dict(self.project, proj_kws)
        self.parent.set_project(self.project)
        self.project.save()
        self.destroy()

    def save(self):
        if self.project:
            self.update_project()
            return
        self.create_project()

    def get_kws(self,
                which: list=['project', 'inspection', 'report']):
        '''return the filled out form as dicts.
        which (list): any of ['project', 'inspection', 'report']
        return order: project_kws, inspection_kws, report_kws'''
        kws = []
        if isinstance(which, str):
            which = [which]
        if not which:
            which = ['project', 'inspection', 'report']
        if 'project' in which:
            proj_kws = {self.projdata[kw][1]: self.projdata[kw][0].get()\
                            for kw in self.projdata.keys()}
            kws.append(proj_kws)
            which.remove('project')
        if 'inspection' in which:
            insp_kws = {self.inspectiondata[kw][1]: self.inspectiondata[kw][0].get()\
                            for kw in self.inspectiondata.keys()}
            kws.append(insp_kws)
            which.remove('inspection')
        if 'report' in which:
            insp_kws = self.get_kws('inspection')
            repo_kws = {'parent_project': self.project,
                        'inspection_kwargs': {'has_happened': False, **insp_kws}}
            kws.append(repo_kws)
            which.remove('report')
        if which:
            raise ValueError(f'unexpected argument for which: {which}')
        if len(kws) == 1:
            return kws[0]
        return kws



class Authors_Frame(ttk.Frame):
    '''Frame containing a mask for entering rows like
    Rolle|Name|Ort|Datum|Bildname Unterschrift'''
    # TODO: make comboboxes as soon as reports have a database
    def __init__(self, master, report):
        super().__init__(master)
        self.grid_columnconfigure(0, weight=100)
        self.report = report

        self.enteringfrm = ttk.Frame(self)
        self.enteringfrm.grid(row=0, column=0, sticky='ew', padx=1, pady=1)
        for i in range(5):
            self.enteringfrm.grid_columnconfigure(i, weight=1)

        self.build_structure()
        self.add_newline_when_full()

    def build_structure(self):
        gui_f.delete_children(self.enteringfrm)
        curr_authors = self.get_curr_authors()

        self.authors_strvars = []
        count_to = len(curr_authors)
        if count_to < 1:
            count_to = 1

        for i, header in enumerate(['Aufgabe', 'Name', 'Ort', 'Datum', 'Bildname Unterschrift']):
            (ttk.Label(self.enteringfrm, text=header)
                    .grid(row=0, column=i, padx=1, pady=1, sticky='w'))

        for i in range(count_to):
            try:
                curr_role = curr_authors[i][0]
            except IndexError:
                curr_role = None
            try:
                curr_name = curr_authors[i][1]
            except IndexError:
                curr_name = None
            try:
                curr_location = curr_authors[i][2]
            except IndexError:
                curr_location = None
            try:
                curr_date = curr_authors[i][3]
            except IndexError:
                curr_date = None
            try:
                curr_sign = curr_authors[i][4]
            except IndexError:
                curr_sign = None

            role_entry, name_entry, location_entry, date_entry, sign_entry = self.create_new_line(
                curr_role, curr_name, curr_location, curr_date, curr_sign
                )
            role_entry.grid(row=i+1, column=0, padx=1, pady=1, sticky='ew')
            name_entry.grid(row=i+1, column=1, padx=1, pady=1, sticky='ew')
            location_entry.grid(row=i+1, column=2, padx=1, pady=1, sticky='ew')
            date_entry.grid(row=i+1, column=3, padx=1, pady=1, sticky='ew')
            sign_entry.grid(row=i+1, column=4, padx=1, pady=1, sticky='ew')

    def create_new_line(self, curr_role=None, curr_name=None, curr_location=None,
                        curr_date=None, curr_sign=None):
        rolevar = tk.StringVar(self, value=curr_role)
        role_entry = ttk.Entry(self.enteringfrm,
                               textvariable=rolevar)
        namevar = tk.StringVar(self, value=curr_name)
        name_entry = ttk.Entry(self.enteringfrm,
                               textvariable=namevar)
        locationvar = tk.StringVar(self, value=curr_location)
        location_entry = ttk.Entry(self.enteringfrm,
                               textvariable=locationvar)
        datevar = tk.StringVar(self, value=curr_date)
        date_entry = ttk.Entry(self.enteringfrm,
                               textvariable=datevar)
        signvar = tk.StringVar(self, value=curr_sign)
        sign_entry = ttk.Entry(self.enteringfrm,
                               textvariable=signvar)
        
        for var in [rolevar, namevar, locationvar, datevar, signvar]:
            var.trace_add('write', self.add_newline_when_full)
        self.authors_strvars.append((rolevar, namevar, locationvar, datevar, signvar))
        return role_entry, name_entry, location_entry, date_entry, sign_entry

    def add_newline_when_full(self, *args):
        for rolevar, namevar, locationvar, datevar, signvar in self.authors_strvars:
            if not (rolevar.get() and namevar.get()):
                return
        i = len(self.authors_strvars)+1
        role_entry, name_entry, location_entry, date_entry, sign_entry = self.create_new_line()
        role_entry.grid(row=i, column=0, padx=1, pady=1, sticky='ew')
        name_entry.grid(row=i, column=1, padx=1, pady=1, sticky='ew')
        location_entry.grid(row=i, column=2, padx=1, pady=1, sticky='ew')
        date_entry.grid(row=i, column=3, padx=1, pady=1, sticky='ew')
        sign_entry.grid(row=i, column=4, padx=1, pady=1, sticky='ew')

    def get_curr_authors(self):
        authors = []
        try:
            authors.extend(self.report.get_authors())
        except: pass
        if not authors:
            authors = [('Erstellt:', 'Dipl.-Ing Tade Marten Jensen', 'Sankt Annen', None, 'Tade_sign.jpg'),
                       ('Geprüft:', 'Dipl.-Ing (FH) Axel Jensen', 'Wyk auf Föhr',
                        'DATUM EINFÜGEN',
                        'Axel_sign.png')]
        return authors
    
    def get_authors(self):
        authors = []
        for rolevar, namevar, locationvar, datevar, signvar in self.authors_strvars:
            if not rolevar.get() and not namevar.get():
                continue
            
            authors.append((rolevar.get(), namevar.get(), locationvar.get(),
                               datevar.get(), signvar.get()))
        return authors

class Inspectors_Frame(ttk.Frame):
    '''Frame Containing mask for entering rows like Rolle|Name|Datum|Witterung'''
    def __init__(self, master, wea):
        super().__init__(master)
        self.grid_columnconfigure(0, weight=100)
        self.wea = wea
        self.inspectors_strvars = []
        self.all_inspectors_ever = dbf.get_all_inspectors_ever()

        self.enteringfrm = ttk.Frame(self)
        self.enteringfrm.grid(row=0, column=0, sticky='ew', padx=1, pady=1)
        for i in range(4):
            self.enteringfrm.grid_columnconfigure(i, weight=1)

        self.build_structure()
        self.add_newline_when_full()

    def build_structure(self):
        gui_f.delete_children(self.enteringfrm)
        curr_inspectors = self.get_curr_inspectors()
        self.inspectors_strvars = []
        count_to = len(curr_inspectors)
        if count_to < 2:
            count_to = 2

        for i, header in enumerate(['Aufgabe', 'Name', 'Datum', 'Witterung']):
            (ttk.Label(self.enteringfrm, text=header)
                    .grid(row=0, column=i, padx=1, pady=1, sticky='w'))

        for i in range(count_to):
            if len(curr_inspectors) == 0:
                curr_role, curr_name, curr_date, curr_weather = config.default_inspectors[i]
            else:
                try: curr_role = curr_inspectors[i][0]
                except IndexError: curr_role = None
                try: curr_name = curr_inspectors[i][1]
                except IndexError: curr_name = None
                try: curr_date = curr_inspectors[i][2]
                except IndexError: curr_date = None
                try: curr_weather = curr_inspectors[i][3]
                except IndexError: curr_weather = None
            role_cbx, name_cbx, date_entry, weather_entry = self.create_new_line(
                curr_role, curr_name, curr_date, curr_weather
                )
            role_cbx.grid(row=i+1, column=0, padx=1, pady=1, sticky='ew')
            name_cbx.grid(row=i+1, column=1, padx=1, pady=1, sticky='ew')
            date_entry.grid(row=i+1, column=2, padx=1, pady=1, sticky='ew')
            weather_entry.grid(row=i+1, column=3, padx=1, pady=1, sticky='ew')

    def add_newline_when_full(self, *args):
        for rolevar, namevar, _, _ in self.inspectors_strvars:
            if not (rolevar.get() and namevar.get()):
                return
        i = len(self.inspectors_strvars)+1
        role_cbx, name_cbx, date_entry, weather_entry = self.create_new_line()
        role_cbx.grid(row=i, column=0, padx=1, pady=1, sticky='ew')
        name_cbx.grid(row=i, column=1, padx=1, pady=1, sticky='ew')
        date_entry.grid(row=i, column=2, padx=1, pady=1, sticky='ew')
        weather_entry.grid(row=i, column=3, padx=1, pady=1, sticky='ew')


    def create_new_line(self, curr_role=None,
                        curr_name=None, curr_date=None, curr_weather=None):
        rolevar = tk.StringVar(self, value=curr_role)
        role_cbx = ttk.Combobox(self.enteringfrm,
                                values=list(self.all_inspectors_ever.keys()),
                                textvariable=rolevar)
        namevar = tk.StringVar(self, value=curr_name)
        name_cbx = Conditional_Combobox(self.enteringfrm,
                                        rolevar,
                                        value_dict=self.all_inspectors_ever,
                                        textvariable=namevar)
        datevar = tk.StringVar(self, value=curr_date)
        date_entry = ttk.Entry(self.enteringfrm,
                               textvariable=datevar)
        weathervar = tk.StringVar(self, value=curr_weather)
        weather_entry = ttk.Entry(self.enteringfrm,
                                  textvariable=weathervar)
        rolevar.trace_add('write', self.add_newline_when_full)
        namevar.trace_add('write', self.add_newline_when_full)
        for box in [role_cbx, name_cbx]:
            box.unbind_class('TCombobox', '<MouseWheel>')
            box.unbind_class('TCombobox', '<ButtonPress-4>')
            box.unbind_class('TCombobox', '<ButtonPress-5>')

        self.inspectors_strvars.append((rolevar, namevar, datevar, weathervar))
        return role_cbx, name_cbx, date_entry, weather_entry

    def get_curr_inspectors(self):
        try: return self.wea.report.inspection.get_inspectors()
        except: return []

    def get_inspectors(self):
        inspectors = []
        for rolevar, namevar, datevar, weathervar in self.inspectors_strvars:
            if not rolevar.get() and not namevar.get():
                continue
            inspectors.append((rolevar.get(), namevar.get(),
                               datevar.get(), weathervar.get()))
        return inspectors


class WeaEditor(ttk.Frame):
    '''
    Hier werden die WEAs genau spezifiziert
    Frame für jede WEA in das man die gewünschten Daten eingeben kann
    '''
    def __init__(self, parent, project, wea_id, parent_window=None):
        self.project = project
        self.wea_id = wea_id
        self.wea = self.project.windfarm.get_wea(self.wea_id)
        if self.wea.is_setup: self.wea.copy_db_data()
        self.report = self.wea.report
        self.inspection = self.report.inspection
        self.parent_window = parent_window
        super().__init__(parent)

        try: self.coords = str(gui_f.coordinates_from_images(self.wea_id))
        except AttributeError: self.coords = None

        self.grid_columnconfigure(0, weight=100)
        self.grid_rowconfigure(0, weight=100)
        
        scrlfrm = ScrollFrame(self, def_height=450)
        scrlfrm.grid(row=0, column=0, sticky='nsew', columnspan=2)

        self.workfrm = scrlfrm.viewPort
        self.workfrm.grid_columnconfigure(1, weight=100)

    	# EntryFrames for WEA, Inspection and report properties
        copywea_frm = ttk.Frame(self.workfrm)
        copywea_frm.grid(row=0, column=0, padx=1, pady=1, sticky='ew', columnspan=2)

        self.copywea_strvar = tk.StringVar(self, value='')
        available_weas = self.project.windfarm.get_setup_wea_ids()
        try: available_weas.remove(self.wea_id)
        except ValueError: pass
        if available_weas:
            copywea_combo = ttk.Combobox(copywea_frm,
                                        values=available_weas,
                                        textvariable=self.copywea_strvar,
                                        state='readonly')
            copywea_combo.pack(side='right', padx=1, pady=1)
            ttk.Label(copywea_frm, text='Daten kopieren von:   ')\
                .pack(side='right', padx=1, pady=1)
            
        self.avg_power_strvar = tk.StringVar(self, value='Betriebsstunden und Leistung ausfüllen.')
        
        immutable = ['Hersteller', 'Seriennummer'] if self.wea.is_setup else ['Seriennummer']
        oem = self.wea.get('oem')
        self.wea_EF = EntryFrame(self.workfrm, list(turbine_properties.keys()),
            prefilled_object=self.wea,
            clearname2argname=turbine_properties,
            immutable=immutable,
            multiline=['Betreiber'],
            prefill_options={'Hersteller': dbf.get_all('oem'),
                             'Modell': {'depend_on': 'Hersteller', 
                                     **dbf.get_oem_model_dict()},
                             'Untermodell': {'depend_on': 'Modell', 
                                     **dbf.get_oem_model_X_dict(oem, 'submodel')},
                             'Nennleistung (kW)': {'depend_on': 'Modell', 
                                     **dbf.get_oem_model_X_dict(oem, 'rated_power')},
                             'Nabenhöhe (m)': {'depend_on': 'Modell', 
                                     **dbf.get_oem_model_X_dict(oem, 'hub_height')},
                             'Turmtyp': dbf.get_all('tower_type', hybrid=True),
                             'Betreiber': lambda *_: Prefiller(self,
                                     self.wea_EF.get_textvars()['Betreiber'],
                                     dbf.get_all('owner')),
                             'Betriebsführer': dbf.get_all('operator'),
                             'Koordinaten (XX, YY)': self.prefill_coordinates})
        self.inspection_EF = EntryFrame(self.workfrm,
                             list(inspection_properties.keys()),
                             prefilled_object=self.inspection,
                             clearname2argname=inspection_properties,
                             immutable=['Prüfauftrag', 'Jahr/ID'],
                             prefill_options={'Prüfumfang': dbf.get_all_scopes_ever()})
        self.inspector_selector = Inspectors_Frame(self.workfrm, self.wea)
        self.report_EF = EntryFrame(self.workfrm,
                             list(report_properties.keys()),
                             prefilled_object=self.report,
                             prefill_options={
                                 'Nummer': self.prefill_inspectionnumber
                             },
                             clearname2argname=report_properties,
                             )
        self.author_selector = Authors_Frame(self.workfrm, self.report)
        self.build_structure()
        self.set_avg_power()
        self.add_bindings()

    def build_structure(self):
        def place_EF(EF: EntryFrame, row: int):
            '''place an entryframe and return the next free row'''
            for label, entry, btn in EF.get_labels_entries_buttons():
                stick='n' if isinstance(entry, StrVarText) else ''
                label.grid(row=row, column=0, pady=1, padx=1, sticky=f'{stick}w')
                entry.grid(row=row, column=1, pady=1, padx=5, sticky='ew')
                if btn is not None:
                    btn.grid(row=row, column=2, padx=1, pady=1, sticky=f'{stick}e')
                row += 1

                # specials....
                if btn is not None and label.cget('text') == 'Koordinaten (XX, YY)':
                    if not self.coords: btn.configure(state='disabled')
                if label.cget('text') == 'Ertrag (kWh)':
                    ttk.Label(self.workfrm, text='Durchschnittsleistung')\
                        .grid(row=row, column=0, pady=1, padx=1, sticky='w')
                    ttk.Label(self.workfrm, textvariable=self.avg_power_strvar)\
                        .grid(row=row, column=1, pady=1, padx=5, sticky='ew')
                    row += 1
            return row
        
        def place_Separator(row: int):
            ttk.Separator(self.workfrm, orient='horizontal')\
              .grid(row=row, column=0, columnspan=3, padx=5, pady=1, sticky='ew')
            return row + 1



        ttk.Label(self.workfrm, text='WEA Daten:').grid(row=1, column=0, pady=1,
                                                        padx=1, sticky='w')
        i = 2
        i = place_EF(self.wea_EF, i)
        i = place_Separator(i)      
        ttk.Label(self.workfrm, text='Inspektionsdaten:').grid(row=i, column=0,
                                                    pady=1, padx=1, sticky='w')
        i += 1
        i = place_EF(self.inspection_EF, i)
        self.inspector_selector.grid(row=i, column=0, columnspan=2, padx=1,
                                     pady=1, sticky='ew')
        i += 1
        i = place_Separator(i)
        ttk.Label(self.workfrm, text='Berichtsdaten:').grid(row=i, column=0,
                                                    pady=1, padx=1, sticky='w')
        i += 1
        i = place_EF(self.report_EF, i)
        self.author_selector.grid(row=i, column=0, columnspan=2, padx=5, pady=1, sticky='ew')
        i += 1

        self.savebtn = ttk.Button(self,
                                  text='OK',
                                  command=self.save_changes)
        self.savebtn.grid(row=1, column=1, sticky='e', padx=1, pady=1)

    def add_bindings(self):
        self.copywea_strvar.trace_add('write', self.copy_from_other_turbine)
        for clearname in ['Betriebsstunden', 'Ertrag (kWh)']:
            strvar = self.inspection_EF.get_textvars()[clearname]
            strvar.trace_add('write', self.set_avg_power)
        strvar = self.wea_EF.get_textvars()['Hersteller']
        strvar.trace_add('write', self.prefill_from_db)
        strvar.trace_add('write', self.update_conditional_combos)

    def prefill_from_db(self, *_):
        oem = self.wea_EF.get_textvars()['Hersteller'].get()
        self.wea.__setattr__('oem', oem)
        db = self.wea.get_db_entry().dropna()
        if db.empty: return
        for kw in db.index:
            if kw in ['oem', 'id']: continue
            txtvar = self.wea_EF.get_textvars()[clearname]
            if txtvar.get(): continue
            val = db[kw]
            clearname = gui_f.argname2clearname(kw, turbine_properties)
            txtvar.set(val)
    
    def copy_from_other_turbine(self, *_):
        turb_id = self.copywea_strvar.get().strip()
        other_wea = self.project.windfarm.get_wea(turb_id)
        for clearname, kw in list(turbine_properties.items())[:-5]:
            this_wea_strvar = self.wea_EF.get_textvars()[clearname]
            if (curr := this_wea_strvar.get().strip())\
                and curr != 'non_setup': continue
            other_wea_val = other_wea.get(kw)
            this_wea_strvar.set(other_wea_val)

    def update_conditional_combos(self, *_):
        oem = self.wea_EF.get_textvars()['Hersteller'].get().strip()
        for clearname, entryline in self.wea_EF.get_entrylines().items():
            entry = entryline.get_entry()
            if clearname not in ['Untermodell', 'Nennleistung (kW)', 'Nabenhöhe (m)']:
                continue
            if not isinstance(entry, Conditional_Combobox):
                raise TypeError(f'expected Conditional_Combobox, got {type(entry)} for {clearname}')
            entry.replace_value_dict(
                dbf.get_oem_model_X_dict(oem, turbine_properties[clearname]))
            
    
    def set_avg_power(self, *_):
        hours = self.inspection_EF.get_textvars()['Betriebsstunden'].get()
        output = self.inspection_EF.get_textvars()['Ertrag (kWh)'].get()
        if not hours or not output:
            self.avg_power_strvar.set('Betriebsstunden und Leistung ausfüllen.')
            return
        try:
            hours = float(hours)
            output = float(output)
        except ValueError:
            self.avg_power_strvar.set('Betriebsstunden und Leistung nicht als Zahlen erkannt.')
            return
        avg_power = np.round(output / hours, 0)
        self.avg_power_strvar.set(f'{int(avg_power)} kW')


    def save_changes(self, *_):
        # change/add infos for turbine, inspection and report
        turbine_infos = self.wea_EF.get_argname_entry_dict()
        inspection_infos = self.inspection_EF.get_argname_entry_dict()
        inspection_infos['inspectors_list'] = self.inspector_selector.get_inspectors()
        report_infos = self.report_EF.get_argname_entry_dict()
        report_infos['authors'] = self.author_selector.get_authors()
        self.init_farm_infos()
        for (obj, argdict) in [(self.wea, turbine_infos), 
                               (self.report, report_infos),
                               (self.inspection, inspection_infos)]:
            gui_f.change_attributes_from_dict(obj, argdict)
        self.report.set_authors(self.report.authors)
        if not self.wea.oem or not self.wea.id or not self.wea.model:
            ErrorWindow(self, 'Bitte mindestens OEM, Modell und ID angeben.',
                        self.focus)
            return


        try: self.wea.setup()
        except Exception as e:
            ErrorWindow(self, f'Fehler beim Einrichten der WEA: {traceback.format_exc()}',
                        self.focus)
            self.project.save()
            return
        self.project.save()
        # save project with new wea data, even if setup is incomplete
        if isinstance(self.parent_window, MetadataEditor):
            self.parent_window.on_close()
            self.parent_window.destroy()
        self.destroy()
        try:
            if self.parent_window is not None:
                self.parent_window.update_weapage(self.wea_id)
        except AttributeError: pass

    def init_farm_infos(self):
        windfarm = self.project.windfarm
        turb_info = self.wea_EF.get_argname_entry_dict()
        farm_info = {'name': turb_info['windfarm'],
                     'location': turb_info['location']}
        for attr, val in farm_info.items():
            if windfarm.__getattribute__(attr): continue
            windfarm.__setattr__(attr, val)

    def prefill_coordinates(self):
        txtvar = self.wea_EF.get_textvars()['Koordinaten (XX, YY)']
        txtvar.set(self.coords)

    def prefill_inspectionnumber(self):
        weadata = self.wea_EF.get_argname_entry_dict()
        inspectiondata = self.inspection_EF.get_argname_entry_dict()

        farm = weadata['windfarm'].replace(' ', '')
        number = weadata['farm_number']
        id = weadata['id']

        year = self.wea.report.parent_project.year_id
        kind = inspectiondata['kind']

        inspectionnumber = f'8.2-{farm}{number}-{id}-{year}-{kind}'

        inspectionnumber_txtvar = self.report_EF.get_textvars()['Nummer']
        inspectionnumber_txtvar.set(inspectionnumber)


class MetadataEditor(tk.Toplevel):
    '''window to change the infos of report, inspection, turbine, project
    immutable infos: project id, year/id, turbine oem/id, inspection type
    (reason: they are used as index values, making them mutable is an optional
    TODO for later)'''

    def __init__(self, parent, project, wea_id, on_close=lambda *_:None):
        self.on_close = on_close        
        super().__init__(parent)
        self.minsize(800, 500)
        self.title('Randdaten bearbeiten')
        self.grid_columnconfigure(0, weight=100)
        self.grid_rowconfigure(0, weight=100)
        self.project=project
        self.wea_id = wea_id
        setup_frm = WeaEditor(self, self.project, self.wea_id, parent_window=self)
        setup_frm.grid(sticky='nsew')
        self.bind('<Escape>', lambda event: self.destroy()) 
        self.bind('<Control-Return>', setup_frm.save_changes)
        setup_frm.focus()


class Flag_Combobox(ttk.Combobox):
    def __init__(self, master, textvariable: tk.StringVar, include_blank=False):
        self.textvariable = textvariable
        self.vals = ['V', 'P', 'I', 'E', '*', '0', '2', '3', '4', 'S', '-', 'PP', 'PPP', 'RAW']
        if include_blank: self.vals.insert(0, '')
        super().__init__(master, values=self.vals,
                         textvariable=self.textvariable,
                         width=4,
                         state='readonly')
        self.unbind_class('TCombobox', '<MouseWheel>')
        self.unbind_class('TCombobox', '<ButtonPress-4>')
        self.unbind_class('TCombobox', '<ButtonPress-5>')

        for char in self.vals[:self.vals.index('PP')]:
            if char == '':
                continue
            if char == '-':
                self.bind(f'<minus>', lambda *args, c=char: self.textvariable.set(c))
                continue
            if not char == char.lower():
                self.bind(f'<{char.lower()}>', lambda *args, c=char: self.textvariable.set(c))
            self.bind(f'<{char}>', lambda *args, c=char: self.textvariable.set(c))

    def set(self, value: str):
        if pd.isna(value):
            self.textvariable.set('')
            return
        val = str(value).strip()
        if val not in self.vals:
            raise ValueError(f'Flag is {val} but must be one of {self.vals}')
        self.textvariable.set(val)

    def get(self):
        val = self.textvariable.get().strip()
        if val not in self.vals:
            raise ValueError(f'Flag is {val} but must be one of {self.vals}')
        return val


class RemarkEditor(tk.Toplevel):
    '''window that:
        - can be used to insert remark into a turbines's report
        - shows remark address
        - is used to enter remark flag
        - shows all available remarks of a title
        - associates images with inserted remark

    --> used to create remarks for turbines and accessing checklist
    '''
    def __init__(self, master,
                 wea: phys.Turbine,
                 checklist_index: pd.MultiIndex,
                 remark_index: tuple=None,
                 prefill_wea_remarktext: str=None,
                 parent_selection_frame = None,
                 **Toplevel_kwargs):
        # checklist_index als MultiIndex, weil mehrere Einträge zum titel in der
        # checklist sein können --> Fenster bekommt niemals eine speziell variantnr
        # zugewiesen, höchstens den Text aus einem checklist Variante im prefill
        self.wea = wea
        self.remark_index = remark_index
        self.checklist_index = checklist_index
        self.windfarm = self.wea.report.parent_project.windfarm
        self.other_turbines = self.windfarm.get_setup_weas()
        self.other_turbines.remove(self.wea)
        self.parent_selection_frame = parent_selection_frame
        self.selected_images = []
        self.start_selected_images = []

        super().__init__(master, **Toplevel_kwargs)
        self.focus()
        self.minsize(1000, 620)
        self.title(f'{self.wea.id}: Bemerkung hinzufügen')
        self.grid_columnconfigure(0, weight=100)
        self.grid_rowconfigure(0, weight=100)
        self.allfrm = tk.Frame(self, background='light gray')
        self.allfrm.grid_columnconfigure(0, weight=100)
        self.allfrm.grid_rowconfigure(0, weight=100)
        self.allfrm.grid(sticky='nsew')
        self.workframe = ttk.Frame(self.allfrm)
        self.workframe.grid(sticky='nsew', padx=3, pady=3)
        address_row = 0
        entering_row = 2
        checklist_row = 4
        otherweas_row = 6
        images_row = 8
        buttons_row = 10
        self.workframe.grid_columnconfigure(0, weight=100)
        self.workframe.grid_rowconfigure(checklist_row, weight=100)
        self.workframe.grid_rowconfigure(otherweas_row, weight=100)

        # put address of current remark on top
        self.addressfrm = ttk.Frame(self.workframe)
        self.addressfrm.grid(row=address_row, column=0, sticky='ew', padx=3, pady=3)

        self.addressstrvar = tk.StringVar(self)
        self.titlestrvar = tk.StringVar(self)
        if self.checklist_index is not None:
            self.address = f'{self.checklist_index[0]}|{self.checklist_index[1]}'
        elif self.remark_index is not None:
            self.address = self.remark_index[-2]
        else:
            self.address = None
        section, title = dbf.get_sec_title_from_address(self.address)

        self.addressstrvar.set(section)
        self.titlestrvar.set(title)
        self.checklist_entries = pd.DataFrame()

        self.addressentry = ttk.Entry(self.addressfrm, textvariable=self.addressstrvar)
        self.addressentry.grid(row=0, column=0, sticky='ew', padx=3, pady=3)
        self.titleentry = ttk.Entry(self.addressfrm, textvariable=self.titlestrvar)
        self.titleentry.grid(row=0, column=1, sticky='ew', padx=3, pady=3)
        self.addressfrm.grid_columnconfigure(0, weight=2, uniform="foo")
        self.addressfrm.grid_columnconfigure(1, weight=1, uniform="foo")

        # place Frame for entering adding remark to wea
        self.wea_flagstrvar = tk.StringVar(self)
        self.wea_remarkstrvar = tk.StringVar(self)
        self.rem_posstrvar = tk.StringVar(self)
      
        # prefill from existing remark
        if self.remark_index is not None:
            remarks_db = self.wea.report.get_remarks(ordered=False)
            self.existing_remark = remarks_db.loc[self.remark_index[-2:]]
            self.wea_flagstrvar.set(self.existing_remark.flag)
            self.wea_remarkstrvar.set(gui_f.dbtext2displaytext(self.existing_remark.fulltext))
            pos_nr = self.existing_remark.position
            if not pd.isna(pos_nr) and pos_nr is not None and pos_nr != '':
                self.rem_posstrvar.set(int(pos_nr))
            # get existing remark images
            self.selected_images = self.existing_remark.image_names
            if isinstance(self.selected_images, str):
                if self.selected_images == '':
                    self.selected_images = []
                self.selected_images = eval(self.selected_images)
            elif isinstance(self.selected_images, list):
                pass
            else:
                self.selected_images = []
            self.start_selected_images = [im for im in self.selected_images] # copy that list

        elif prefill_wea_remarktext is not None:
            self.wea_remarkstrvar.set(prefill_wea_remarktext)

        ttk.Separator(self.workframe).grid(row=entering_row-1, column=0, 
                                           sticky='ew', padx=5, pady=2)
        
        # frame for entering the remark
        self.wea_enteringframe = ttk.Frame(self.workframe)
        self.wea_enteringframe.grid(row=entering_row, column=0, sticky='ew',
                                    padx=3, pady=1)
        
        pos_col = 0
        flag_col = 1
        marker_col = 2
        remtext_col = 3
        self.wea_enteringframe.grid_columnconfigure(remtext_col, weight=100)
       
        # order entry
        (ttk.Label(self.wea_enteringframe, text='Pos.:')
            .grid(row=0, column=pos_col, sticky='nw', pady=2, padx=2))
        self.posentry = ttk.Entry(self.wea_enteringframe,
                               textvariable=self.rem_posstrvar,
                               width=4)
        self.posentry.grid(row=1, column=pos_col, sticky='nw', padx=2, pady=2)
        
        # flag entry
        (ttk.Label(self.wea_enteringframe, text='Flag:')
            .grid(row=0, column=flag_col, sticky='nw', pady=2, padx=2))
        self.wea_flagentry = Flag_Combobox(self.wea_enteringframe,
                                           textvariable=self.wea_flagstrvar)
        self.wea_flagentry.grid(row=1, column=flag_col, padx=2, pady=2, sticky='nw')

        # marker for remark flag color
        self.markerfrm = tk.Frame(self.wea_enteringframe, background='snow', width=5)
        self.markerfrm.grid(row=1, column=marker_col, sticky='ns', padx=1, pady=1)
        self.change_marker_color()
       
        # remark entry as text for multiline support
        (ttk.Label(self.wea_enteringframe, text='Bemerkungstext:')
            .grid(row=0, column=remtext_col, sticky='nw', pady=2, padx=2))
        self.wea_remarktext = StrVarText(self.wea_enteringframe,
                                         height=5,
                                         wrap='word',
                                         textvariable=self.wea_remarkstrvar)
        self.wea_remarktext.configure(font=font.nametofont('TkDefaultFont'))
        self.wea_remarktext.grid(row=1, column=remtext_col, sticky='nsew', padx=2, pady=2)
        # self.wea_remarktext.insert('1.0', self.wea_remarkstrvar.get())
       
        # add to checklist button
        self.add_to_checklist_btn = ttk.Button(self.wea_enteringframe,
                                                text='Zu Vorlage hinzufügen',
                                                command=self.add_to_checklist)
        self.add_to_checklist_btn.grid(row=2, column=remtext_col,
                                       padx=3, pady=2,
                                       sticky='ne')
                
        # add checklist_entries for selection
        ttk.Separator(self.workframe).grid(row=checklist_row-1, column=0, 
                                           sticky='ew', padx=5, pady=2)
        self.update_checklist_frm()
        self.checklist_frm.grid(row=checklist_row, column=0, sticky='nsew', padx=3, pady=3)

        # text von anderen weas kopieren
        ttk.Separator(self.workframe).grid(row=otherweas_row-1, column=0, 
                                           sticky='ew', padx=5, pady=2)
        scrl = ScrollFrame(self.workframe, orient='vertical', def_height=200)
        scrl.grid(row=otherweas_row, column=0, sticky='nswe', pady=3, padx=3)
        self.otherweasfrm = scrl.viewPort
        self.otherweasfrm.grid_columnconfigure(3, weight=100)
        other_remarks = self.get_other_turbines_remarks(formatted=True)
        if len(self.other_turbines) == 0 or \
            self.address is None or \
            len(other_remarks) == 0:
            
            (ttk.Label(self.otherweasfrm,
                       text='Bemerkung nicht in anderen WEAs vorhanden.')
             .grid(column=0, row=0, sticky='w'))
        else:
            for i, (wea_str, flag, text) in enumerate(other_remarks):
                ttk.Label(self.otherweasfrm, text=wea_str)\
                    .grid(row=i, column=0, sticky='nw', padx=1, pady=1)
                ttk.Button(self.otherweasfrm, text='+',
                           command=lambda flag=flag, text=text:\
                                   self.copy_remark(flag, text),
                           width=3)\
                    .grid(row=i, column=1, sticky='nw', padx=1, pady=1)
                ttk.Label(self.otherweasfrm, text=flag)\
                    .grid(row=i, column=2, sticky='nw', padx=1, pady=1)
                MultilineLabel(self.otherweasfrm, textvar=tk.StringVar(self, value=text))\
                    .grid(row=i, column=3, sticky='ew', padx=1, pady=1)

        # final buttons: OK, Bilder hinzufügen, Checklist manipulieren
        ttk.Separator(self.workframe).grid(row=buttons_row-1, column=0, 
                                           sticky='ew', padx=5, pady=2)
        self.finalbuttonsfrm = ttk.Frame(self.workframe)
        self.finalbuttonsfrm.grid(row=buttons_row, column=0, pady=3, padx=3, sticky='ew')

        self.ok_btn = ttk.Button(self.finalbuttonsfrm,
                                 text='Speichern',
                                 command=self.add_remark_to_report)
        self.open_checklist_manipulation_window_btn = ttk.Button(
            self.finalbuttonsfrm, text='Vorlage anpassen',
            command=lambda: self.add_to_checklist(prefill=False)
        )

        self.ok_btn.pack(padx=3, pady=3, side='right')
        self.open_checklist_manipulation_window_btn.pack(padx=3, pady=3, side='right')
        ttk.Separator(self.workframe).grid(row=images_row-1, column=0, 
                                           sticky='ew', padx=5, pady=2)
        self.images_frm = ttk.Frame(self.workframe)
        self.images_frm.grid(row=images_row, column=0, pady=3, padx=3, sticky='ew')
        self.update_images_frm()
        self.wea_flagstrvar.trace_add('write', self.change_marker_color)
        self.bind('<Escape>', lambda _: self.destroy())
        self.bind('<Control-Return>', lambda _: self.add_remark_to_report())
        self.bind('<Control-b>', lambda _: self.open_image_selector())
        self.bind('<Control-B>', lambda _: self.open_image_selector())
        self.wea_remarktext.bind('<Control-Return>', lambda _: self.add_remark_to_report())
        self.wea_remarktext.focus()

        self.allfrm.register_drop_target('*')
        self.allfrm.bind('<<Drop>>', self.add_images_from_drop)
        self.allfrm.bind('<<DropLeave>>', self.unhighlight)
        self.allfrm.bind('<<DropEnter>>', self.highlight)

    def update_checklist_frm(self, *_):
        try:
            self.checklist_entries = (self.wea.report.
                                      get_checklist().loc[self.get_chapter(),
                                                          self.get_title()])
        except KeyError:
            self.checklist_entries = pd.DataFrame()
        except AssertionError:
            self.checklist_entries = pd.DataFrame()

        addcol = 0
        flagcol = 1
        textcol = 2
        try: gui_f.delete_children(self.checklist_frm, leave_out=self.searchbarfrm)
        except AttributeError:
            self.checklist_frm = ttk.Frame(self.workframe)
            self.checklist_frm.grid_columnconfigure(textcol, weight=100)
            self.searchtxtvar = tk.StringVar(self)
            self.searchtxtvar.trace_add('write', self.update_checklist_frm)
            self.searchbarfrm = ttk.Frame(self.checklist_frm)
            searchbar = ttk.Entry(self.searchbarfrm, textvariable=self.searchtxtvar)
            searchbar.pack(side='left', padx=1, pady=1, fill='x', expand=True)
            ttk.Separator(self.searchbarfrm, orient='vertical')\
                .pack(side='left', pady=1, padx=1, fill='y')
            searchall_btn = ttk.Button(self.searchbarfrm, text='Alles durchsuchen',
                                       command=self.open_Prefiller_searchall)
            searchall_btn.pack(side='left', pady=1, padx=1)            
            self.searchbarfrm.grid(row=0, column=flagcol, columnspan=2,
                              padx=1, pady=1, sticky='ew')

        searchterm = self.searchtxtvar.get()
        for i, ind in enumerate(self.checklist_entries.index, 1):
            cl_entry = self.checklist_entries.loc[ind]
            text = gui_f.dbtext2displaytext(cl_entry.fulltext)
            if searchterm.lower() not in text.lower():
                continue
            try: rec_flag = cl_entry.recommended_flag.upper()
            except AttributeError: rec_flag = ''
            btn = ttk.Button(self.checklist_frm, text='+', width=3,
                command=lambda flag=rec_flag, text=text: self.copy_remark(flag, text))
            flaglbl = ttk.Label(self.checklist_frm,
                                text=f'({rec_flag})' if rec_flag else '')
            textvar = tk.StringVar(self, value=text)
            textlbl = MultilineLabel(self.checklist_frm, textvar, wraplength=950)
            btn.grid(row=i, column=addcol, padx=3, pady=1, sticky='n')
            flaglbl.grid(row=i, column=flagcol, padx=3, pady=1, sticky='n')
            textlbl.grid(row=i, column=textcol, padx=1, pady=1, sticky='new')

        
    def add_remark_to_report(self):
        flag = self.wea_flagstrvar.get().upper().strip()
        remarktext = self.get_remarktext().strip()
        if flag != 'RAW':
            remarktext = remarktext.replace('\n', db_split_char)
        chap = self.addressstrvar.get()
        title = self.titlestrvar.get()
        pos_nr = self.rem_posstrvar.get().strip()
        selected_images = self.selected_images if self.selected_images else []
        # if nothing is in textbox, remarks will not be deleted (intentional)!
        if len(remarktext) == 0 and flag == '' and not self.selected_images:
            self.destroy()
            return
        if len(remarktext) == 0:
            ErrorWindow(self, 'Kein Bemerkungstext angegeben.',
                        on_end=lambda: self.wea_remarktext.focus())
            return 'break'
        if flag == '':
            ErrorWindow(self, 'Keine Kennzeichnung angegeben.',
                        on_end=lambda: self.wea_flagentry.focus())
            return 'break'

        if chap == 'XXX':
            ErrorWindow(self, 'Kapitel muss "Kapitel|Unterkapitel|..." sein, nicht XXX.',
                        on_end=lambda: self.addressentry.focus())
            return 'break'
        if title == 'XXX':
            ErrorWindow(self, 'Titel der Bemerkung sollte den Inhalt zusammenfassen, nicht XXX sein.',
                        on_end=lambda: self.titleentry.focus())
            return
        if chap == '':
            ErrorWindow(self, 'Kapitel muss "Kapitel|Unterkapitel|..." sein, nicht leer.',
                        on_end=lambda: self.addressentry.focus())
            return 'break'
        if title == '':
            ErrorWindow(self, 'Titel der Bemerkung sollte den Inhalt zusammenfassen, nicht leer sein.',
                        on_end=lambda: self.titleentry.focus())
            return 'break'
        if pos_nr:
            try: pos_nr = int(float(pos_nr))
            except:
                ErrorWindow(self, 'Positionsnummer muss ganzzahlig und > 0 sein '
                                 f'oder leer, ist aber {pos_nr}')
                self.posentry.focus()
                return 'break'
            if pos_nr < 1:
                ErrorWindow(self, 'Positionsnummer muss ganzzahlig und > 0 sein '
                                 f'oder leer, ist aber {pos_nr}')
                self.posentry.focus()
                return 'break'
        
        chapter_order = dbf.get_order('chapters')
        if chap not in chapter_order.keys():
            Orderer(self, 'chapter', chap)
            return 'break'
        if title not in chapter_order[chap]:
            Orderer(self, 'title', title,
                    address=chap)
            return 'break'
        
        if self.remark_index is not None:
            existing_timestamp = self.remark_index[-1]
            self.wea.report.remove_remark(self.address,
                                          existing_timestamp)
        else:
            existing_timestamp = None
        chap_is_new = chap not in self.wea.report.get_chapters()
        self.wea.report.add_remark(address=self.get_address(),
                                   flag=flag,
                                   text=remarktext,
                                   pos_nr=pos_nr,
                                   timestamp=existing_timestamp,
                                   image_names=selected_images)
        ims = self.start_selected_images
        ims.extend(selected_images)
        ims = list(set(ims))
        self.wea.report.update_compressed_images(check_only=ims)
        if chap_is_new:
            self.parent_selection_frame.parent_turbine_viewer.parent_mainwindow.update_weapage(self.wea.id)
            # destroys itself, because turbine viewer gets destroyed
            self.destroy()
            return 'break'
        try:
            self.parent_selection_frame.build_selection_body(
                        set_info=False if chap=='Prüfergebnis|Fazit' else True)
        except:
            self.parent_selection_frame.parent_turbine_viewer.change_chapter(chap)
        self.parent_selection_frame.parent_turbine_viewer.focus()
        self.destroy()
        return 'break'


    def get_address(self):
        return f'{self.get_chapter()}|{self.get_title()}'
    def get_chapter(self):
        return self.addressstrvar.get()
    def get_title(self):
        return self.titlestrvar.get()
    
    def open_Prefiller_searchall(self):
        all_values = dbf.get_all_db_values()
        pf = Prefiller(self, self.wea_remarkstrvar, all_values,
                       init_searchtext=self.searchtxtvar.get())
        pf.searchvar.set(self.searchtxtvar.get())

    def open_image_selector(self):
        '''opens file explorer whose selected images are
        appended to the remark's images'''
        filenames = fd.askopenfilenames(
            filetypes=[('Bilder', ('.png', '.jpg')), ('Alle', '.*')],
            initialdir=f'{os.getcwd()}/{self.wea.id}/0-Fertig',
            title=f'Bilder auswählen für {self.titlestrvar.get()}')
        filenames = [path[path.rfind('/')+1:] for path in filenames]
        self.focus()
        for imagename in list(filenames):
            self.add_image(imagename)

    def add_image(self, image):
        if image.split('.')[-1].lower() not in ['png', 'jpg', 'jpeg']:
            ErrorWindow(self, (f'Bild {image} wurde nicht hinzugefügt. '
                                'Format muss png oder jpg sein.'))
            return
        self.selected_images.append(image)
        self.update_images_frm()

    def add_images_from_drop(self, drop):
        for raw_im in drop.data.split('} {'):
            im = raw_im.split('/')[-1]
            if im.endswith('}'):
                im = im[:-1]
            self.add_image(im)
        self.unhighlight()



    def update_images_frm(self):
        gui_f.delete_children(self.images_frm)
        add_btn = ttk.Button(self.images_frm, text='+', width=3,
                             command=self.open_image_selector)
        add_btn.pack(side='right', padx=2, pady=2, anchor='s')

        if not self.selected_images:
            lbl = ttk.Label(self.images_frm,
                            text=('Keine Bilder angefügt. "+" Drücken oder '
                                  'Bilder aus 0-Fertig in dieses Fenster ziehen.'))
            lbl.pack(side='left', padx=3, pady=3)
            return
        
        width = 120
        for image in self.selected_images:
            frm = ttk.Frame(self.images_frm)
            frm.pack(side='left')
            image_path = f'{os.getcwd()}/{self.wea.id}/0-Fertig/{image}'
            try:
                tk_pic = gui_f.get_tkinter_pic(image_path, master=self,
                                               width_pxl=width if 'timeline' not in image else width*2)

                lbl_btn = Label_w_Button(frm, '',
                                         command=self.remove_image,
                                         command_args=[image],
                                         buttonkwargs={'text': '-',
                                                       'width': 3},
                                         image=tk_pic)
                lbl_btn.label.image=tk_pic
                lbl_btn.label.bind('<Button-1>', lambda e, path=image_path: gui_f.open_image(path))
            except FileNotFoundError:
                lbl_btn = Label_w_Button(frm,
                                         f'{image} nicht in 0-Fertig',
                                         command=self.remove_image,
                                         command_args=[image],
                                         buttonkwargs={'text': '-',
                                                       'width': 3})

            lbl_btn.btn.grid(row=1, column=0, pady=1, padx=1)
            lbl_btn.label.grid(row=0, column=0, pady=1, padx=1, sticky='ew')

    def remove_image(self, image):
        self.selected_images.remove(image)
        self.update_images_frm()

    def highlight(self, *_):
        self.allfrm.configure(background='orange')
    def unhighlight(self, *_):
        self.allfrm.configure(background='light gray')

    def get_remarktext(self):
        raw_text = self.wea_remarkstrvar.get()
        text = raw_text.strip()
        return text
    
    def add_to_checklist(self, prefill=True):
        prefill_tup = ('', '')
        self.searchtxtvar.set('')
        if prefill:
            prefill_tup = (self.wea_flagstrvar.get(), self.get_remarktext())

        title = self.titlestrvar.get()
        if title.strip().lower() in ['', 'xxx']:
            ErrorWindow(self, ('Bitte aussagekräftigen Bemerkungstitel einfügen,'
                               f' nicht "{title}".'),
                               self.titleentry.focus())
            return

        ChecklistEditor(
                self, self.addressstrvar.get(),
                self.titlestrvar.get(), self.wea,
                prefill=prefill_tup,
                parent_RemEditor=self
            )
        
    def get_other_turbines_remarks(self, formatted=False) -> list:
        '''returns dict of tuples like
        {text1: [(flag1, wea_id1), ..., (flagn, wea_idn)], text2: ...}
        '''
        def format_other_remarks(other_remarks) -> list:
            '''summarize other_remarks dict to tuples like
            ('wea_id1 + 3', flag, text)'''
            summary = []
            for text, wea_tup_list in other_remarks.items():
                flag_counter = {}
                flag_identifier = {}
                for flag, wea_id in wea_tup_list:
                    if flag in flag_counter.keys(): flag_counter[flag] += 1
                    else:
                        flag_counter[flag] = 1
                        flag_identifier[flag] = wea_id
                for flag, count in flag_counter.items():
                    counter = count-1
                    leading_wea = flag_identifier[flag]
                    wea_str = f'{leading_wea} + {counter}' if counter else leading_wea
                    summary.append((wea_str, flag, text))
            return summary

        other_remarks = {}
        for wea in self.other_turbines:
            report = wea.report
            remarks = report.get_remarks(self.get_address())
            if remarks is None:
                continue
            if remarks.empty:
                continue
            for ind in remarks.index:
                rem = remarks.loc[ind]
                text = gui_f.dbtext2displaytext(rem.fulltext)
                wea_tup = (rem.flag, wea.get('id'))
                if text in other_remarks.keys():
                    other_remarks[text].append(wea_tup)
                    continue
                other_remarks[text] = [wea_tup]
        
        if formatted: return (format_other_remarks(other_remarks))
        return other_remarks
    
    def copy_remark(self, flag, text):
        curr_text = self.get_remarktext()
        if not text.strip().lower() in curr_text.strip().lower():
            self.wea_remarktext.insert('end', text)
        if self.wea_flagstrvar.get() == '':
            self.wea_flagstrvar.set(flag)
        self.wea_remarktext.focus()

    def change_marker_color(self, *args):
        flag = self.wea_flagstrvar.get()
        if not flag: color = 'snow'
        else: color = dbf.get_flagcolor(flag)
        self.markerfrm.config(background=color)

    
class TimelineEditor(tk.Toplevel):
    '''Window to enter timeline data. Creates timeline image and inserts data 
    to inspection database on OK'''
    name_col = 0
    interval_col = 1
    dates_col = 2

    def __init__(self, master, wea: phys.Turbine,
                 tl_dict: Optional[dict]=None,
                 start_month: Optional[str]=None, end_month: Optional[str]=None,
                 **toplevel_kwargs):
        super().__init__(master, **toplevel_kwargs)
        self.title('Zeitstrahl erzeugen')
        self.minsize(800, 500)
        self.lines = []
        self.master = master
        self.wea = wea
        self.figure = None
        self.in_displaying = False             # flag for automatically updating preview, used only internally to avoid focussing out of an entry...

        self.allfrm = ttk.Frame(self)
        scrl = ScrollFrame(self.allfrm, orient='horizontal', def_width=800)
        self.workfrm = scrl.viewPort
        self.curr_row = 0
        
        self.btnsfrm = ttk.Frame(self.allfrm)
        self.previewfrm = ttk.Frame(self.allfrm)
        self.build_structure(tl_dict, start_month, end_month)

        self.allfrm.pack(fill='both', expand=True, padx=1, pady=1)
        scrl.pack(fill='both', expand=True, padx=1, pady=1)
        self.btnsfrm.pack(fill='x', padx=1, pady=1)
        self.previewfrm.pack(fill='x', padx=1, pady=1)

        self.bind('<Control-Return>', lambda *_: self.save())
        self.focus()


    def build_structure(self, tl_dict: Optional[dict]=None,
                        start_month: Optional[str]=None,
                        end_month: Optional[str]=None):
        
        # info
        ttk.Label(self.allfrm, text='Name der Arbeit: Absatz durch "|" erzeugen.')\
            .pack(padx=1, pady=1, fill='x', anchor='w')
        ttk.Label(self.allfrm, text='Intervall: in Monaten, nur ganze Zahlen oder leer')\
            .pack(padx=1, pady=1, fill='x', anchor='w')
        ttk.Label(self.allfrm, text='Daten: Format mm/yy, z.B. 03/24')\
            .pack(padx=1, pady=1, fill='x', anchor='w')
        
        ttk.Separator(self.allfrm).pack(fill='x', padx=1, pady=1)
        
        # entries for start/end entering
        startendfrm = ttk.Frame(self.workfrm)
        startendfrm.grid(row=self.curr_row, column=0, columnspan=3,
                         padx=1, pady=1, sticky='ew')
        self.startstrvar = tk.StringVar(self, value=start_month)
        self.endstrvar = tk.StringVar(self, value=end_month)
        ttk.Label(startendfrm, text='Start:    ')\
            .grid(row=0, column=0, padx=1, pady=1)
        ttk.Entry(startendfrm, textvariable=self.startstrvar)\
            .grid(row=0, column=1, padx=1, pady=1)
        ttk.Label(startendfrm, text='Ende:    ')\
            .grid(row=1, column=0, padx=1, pady=1)
        ttk.Entry(startendfrm, textvariable=self.endstrvar)\
            .grid(row=1, column=1, padx=1, pady=1)
        
        self.curr_row += 1
        ttk.Separator(self.workfrm).grid(row=self.curr_row, column=0, columnspan=3,
                                         padx=1, pady=1, sticky='ew')
        self.curr_row += 1
        
        # make header
        ttk.Label(self.workfrm, text='Titel')\
            .grid(row=self.curr_row, column=self.name_col, padx=1, pady=1, sticky='w')
        ttk.Label(self.workfrm, text='Intervall')\
            .grid(row=self.curr_row, column=self.interval_col, padx=1, pady=1, sticky='w')
        ttk.Label(self.workfrm, text='Daten')\
            .grid(row=self.curr_row, column=self.dates_col, padx=1, pady=1, sticky='w')
        self.workfrm.grid_columnconfigure(self.dates_col, weight=100)

        # lines for dates entering
        self.curr_row += 1
        if tl_dict:
            for title, dates in tl_dict.items():
                row_prefill = [title, dates]
                line = single_timeLine(self, self.curr_row, row_prefill)
                self.lines.append(line)
                self.curr_row += 1
            self.display_preview()
        self.lines.append(single_timeLine(self, self.curr_row))
        self.curr_row += 1

        ttk.Button(self.btnsfrm, text='OK', command=self.save)\
            .pack(side='right', padx=1, pady=1)
        ttk.Button(self.btnsfrm, text='Vorschau', command=self.display_preview)\
            .pack(side='right', padx=1, pady=1)

    def add_newline_if_full(self, *_):
        for line in self.lines:
            if not line.get_title(): return
        self.lines.append(single_timeLine(self, self.curr_row))
        self.curr_row += 1

    def get_timeline_dict(self):
        _, end = self.get_startend()
        end = dbf.monthyear2datetime(end)
        tl_dict = {}
        for line in self.lines:
            if not line.has_dates(): continue
            title = line.get_title(empty_allowed=False)
            dates = line.get_dateslist()
            for i, datestr in enumerate(dates):
                if i == 0 and isinstance(dates[0], int): continue 
                date = dbf.monthyear2datetime(datestr)
                if date > end:
                    dates.remove(datestr)
                    ErrorWindow(self, f'{title}: Datum {datestr} wird ignoriert (neuer als Enddatum).')
            tl_dict[title] = dates
        return tl_dict
    
    def get_startend(self):
        def get_single(what='start'):
            if what == 'start': strvar = self.startstrvar
            elif what == 'end': strvar = self.endstrvar
            else: raise AttributeError(f'what must be start or end, not {what}')

            if not (txt := strvar.get().strip()): return txt
            # test for right format
            dbf.monthyear2datetime(txt)
            return txt
        return get_single('start'), get_single('end')
    
    def save(self):
        '''saves timeline as image and in inspection's database'''
        try:
            start, end = self.get_startend()
        except Exception as e:
            ErrorWindow(self, (f'Fehler in Start/Ende Textboxen. Ist das Format mm/yy (z. B. 11/22)?\n'
                               f'Fehlertext: {e}'))
            return
        if not start or not end:
            ErrorWindow(self, 'Bitte Start und Ende angeben')
            return
        try: tl_dict = self.get_timeline_dict()
        except Exception as e:
            ErrorWindow(self, f'Zeitstrahlfehler: {e}')
            IPS()
            return
        
        insp = self.wea.report.inspection
        insp.timeline = f"({tl_dict}, '{start}', '{end}')"
        insp.insert_to_db()
        
        if len(tl_dict) == 0: self.destroy()
        self.save_timeline()
        self.master.focus()
        self.destroy()

    def get_figure(self):
        '''return pyplot figure showing the timeline'''
        if self.figure: plt.close(self.figure)
        return dbf.get_timeline_figure(self.get_timeline_dict(), *self.get_startend())

    def display_preview(self, *_):
        self.in_displaying = True
        curr_focus = self.focus_get()
        gui_f.delete_children(self.previewfrm)
        fig = self.get_figure()
        canvas = FigureCanvasTkAgg(fig, master=self.previewfrm)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True, padx=3, pady=3)
        self.update()
        try: curr_focus.focus()
        except AttributeError: pass
        self.in_displaying = False


    def save_timeline(self, name='timeline'):
        '''saves the timeline figure in the turbine's 0-Fertig folder'''
        fig = self.get_figure()
        path = f'{os.getcwd()}/{self.wea.get('id')}/0-Fertig/{name}.jpg'
        fig.savefig(path, dpi=300)



class single_timeLine():
    def __init__(self, master: TimelineEditor, row: int,
                 prefill: Optional[list]=None):
        self.master = master
        self.row = row

        try: title = prefill[0]
        except IndexError: title=None
        except TypeError: title=None
        self.titlevar = tk.StringVar(self.master, value=title)
        self.titlevar.trace_add('write', self.master.add_newline_if_full)
        
        try: interval = int(prefill[1][0])
        except ValueError: interval = None
        except IndexError: interval = None
        except TypeError: interval = None
        self.intervalvar = tk.StringVar(self.master, value=interval)

        self.datevars = []

        # entries
        self.titleentry = ttk.Entry(self.master.workfrm,
                                textvariable=self.titlevar)
        self.intervalentry = ttk.Entry(self.master.workfrm,
                                    textvariable=self.intervalvar, width=4)
        self.datesfrm = ttk.Frame(self.master.workfrm)
        if prefill:
            dates = prefill[1][1:] if interval else prefill[1]
            for i, date in enumerate(dates):
                self.add_single_datecol(i, prefill_date=date)

        self.titleentry.grid(row=self.row, column=self.master.name_col,
                             padx=1, pady=1, sticky='ew')
        self.intervalentry.grid(row=self.row, column=self.master.interval_col,
                                padx=1, pady=1, sticky='ew')
        self.datesfrm.grid(row=self.row, column=self.master.dates_col,
                           padx=1, pady=1, sticky='ew')

        self.add_datecol_if_full()

    def add_datecol_if_full(self, *_):
        i = 0
        for datevar in self.datevars:
            if not datevar.get(): return
            i += 1
        self.add_single_datecol(i)
    
    def add_single_datecol(self, col, prefill_date: Optional[str]=None):
        datestrvar = tk.StringVar(self.master, value=prefill_date)
        datestrvar.trace_add('write', self.add_datecol_if_full)
        dateentry = ttk.Entry(self.datesfrm, textvariable=datestrvar, width=10)
        dateentry.grid(row=0, column=col, padx=1, pady=1, sticky='ew')
        dateentry.bind('<Tab>', self.update_preview)
        self.datevars.append(datestrvar)

    def update_preview(self, *_):
        if self.master.in_displaying: return
        try: self.master.display_preview()
        except Exception as e: print(e)
        finally: self.master.in_displaying = False

    def get_title(self, empty_allowed=True):
        title = self.titlevar.get().strip()
        if empty_allowed: return title
        if not title:
            raise ValueError('Title can\'t be empty.')
        return title
    def get_dateslist(self):
        if not self.has_dates(): return []
        try: title = self.get_title()
        except ValueError as e:
            ErrorWindow(self.master,
                        'Alle Titel bei eingetragenen Daten müssen ausgefüllt sein.',
                        self.titleentry.focus())
            raise e
        interval = self.intervalvar.get().strip()
        if interval:
            try: interval = int(interval)
            except ValueError as e:
                ErrorWindow(self.master,
                    f'Intervall muss zahlwertig oder leer sein ({title}).',
                    self.intervalentry.focus())
                raise e
            ret_list = [interval]
        else: ret_list = []
        for datevar in self.datevars:
            date = datevar.get().strip()
            if not date: continue
            try: dbf.monthyear2datetime(date)
            except Exception as e:
                ErrorWindow(self.master,
                            (f'Fehler beim Konvertieren von {date} zu einem Datum. '
                             'Ist das Format mm/yy (z. B. 05/21)?\n'
                             f'Fehlernachricht: {e}'))
                raise e
            ret_list.append(date)
        return ret_list

    def has_dates(self):
        for datevar in self.datevars:
            if datevar.get().strip(): return True
        return False


class ErrorWindow(tk.Toplevel):
    '''simple window for showing some text'''
    def __init__(self, master, text, on_end=None) -> None:
        super().__init__(master)
        self.text=text
        self.on_end = on_end

        self.workfrm = ttk.Frame(self)
        self.errlbl = ttk.Label(self.workfrm, text=self.text)
        self.okbtn = ttk.Button(self.workfrm, text='OK',
                                command=self.end)

        self.workfrm.grid(sticky='nsew')
        self.errlbl.grid(row=0, column=0, sticky='ew', padx=3, pady=3)
        self.okbtn.grid(row=1, column=0, sticky='e', padx=3, pady=3)
        self.bind('<Escape>', self.end)
        self.bind('<Return>', self.end)
        self.okbtn.focus()


    def end(self, *args):
        if self.on_end is not None:
            self.on_end()
        self.destroy()

class ChecklistEditor(tk.Toplevel):
    '''
    '''
    colors_translator = {'keine': '', 'grün': 'green2', 'gelb': 'yellow',
                         'rot': 'red', 'blau': 'deep sky blue',
                         'lila': 'purple', 'orange': 'orange'}
    def __init__(self, master,
                 address: Optional[str]=None,
                 title: Optional[str]=None,
                 wea: Optional[phys.Turbine]=None,
                 prefill: Optional[tuple]=('', ''),
                 parent_RemEditor: Optional[RemarkEditor]=None,
                 ) -> None:
        '''prefill: (flag, text)'''
        super().__init__(master)
        self.minsize(800, 500)
        prettytitle = f'{address}|{title}'.replace('|', ' > ')
        self.title(f'Checklist bearbeiten ({prettytitle})')
        self.bind('<Escape>', lambda _: self.destroy())
        self.bind('<Control-Return>', lambda _: self.update_checklist())

        self.chapter = address
        self.title = title
        self.prefill = prefill
        self.wea = wea
        self.rems = {}
        self.RemEditor = parent_RemEditor
        self.delete_ogindex = []

        self.allfrm = ttk.Frame(self)
        self.allfrm.grid_columnconfigure(0, weight=100)
        self.allfrm.grid_rowconfigure(0, weight=100)

        self.colorstrvar = tk.StringVar(self)

        self.pos_col, self.recflag_col, self.defaultflag_col, self.text_col, self.delete_col = (0, 1, 2, 3, 4)
        self.head_frm = ttk.Frame(self.allfrm)
        self.newrem_frm = ttk.Frame(self.allfrm)
        self.newrem_frm.grid_columnconfigure(self.text_col, weight=100)
        scrl = ScrollFrame(self.allfrm)
        self.oldrem_frm = scrl.viewPort
        self.oldrem_frm.grid_columnconfigure(self.text_col, weight=100)
        self.tail_frm = ttk.Frame(self.allfrm)

        self.build_framing()
        self.build_new_entry()
        try:
            self.build_old_entries()
        except IndexError as e: ErrorWindow(self, e, self.destroy)
        
        self.allfrm.pack(fill='both', expand=True)
        self.head_frm.pack(side='top', fill='x', padx=1, pady=1)
        ttk.Separator(self.allfrm).pack(side='top', fill='x', expand=True, padx=1, pady=1)
        self.newrem_frm.pack(side='top', fill='x', padx=1, pady=1, expand=True)
        ttk.Separator(self.allfrm).pack(side='top', fill='x', expand=True, padx=1, pady=1)
        scrl.pack(side='top', fill='both', expand=True, padx=1, pady=1)
        ttk.Separator(self.allfrm).pack(side='top', fill='x', expand=True, padx=1, pady=1)
        self.tail_frm.pack(side='top', fill='x', padx=1, pady=1)
        self.focus()

    def build_framing(self):
        '''create entries for adderss/title and okay button'''
        gui_f.delete_children(self.head_frm)
        ttk.Combobox(self.head_frm, textvariable=self.colorstrvar,
            values=list(self.colors_translator.keys()),
            state='readonly')\
                .pack(side='right', padx=1, pady=1)
        ttk.Label(self.head_frm, text='Markerfarbe:     ')\
                .pack(side='right', padx=1, pady=1)
        self.set_markercolor()

        gui_f.delete_children(self.tail_frm)
        ok_btn = ttk.Button(self.tail_frm, text='OK',
                            command=self.update_checklist)
        ok_btn.pack(side='right', padx=1, pady=1)

    def set_markercolor(self):
        try:
            colors = self.get_checklist_slice()['bg_color'].dropna().unique()
        except KeyError: return
        if len(colors) == 0: return

        for c in colors:
            for i in self.colors_translator.items():
                if c.strip() == i[1]:
                    self.colorstrvar.set(i[0])
                    break
            else: 
                continue
        

    def build_new_entry(self):
        gui_f.delete_children(self.newrem_frm)

        ttk.Label(self.newrem_frm, text='Pos.')\
            .grid(row=0, column=self.pos_col, padx=2, pady=1)
        ttk.Label(self.newrem_frm, text='Flags')\
            .grid(row=0, column=self.recflag_col, padx=2, pady=1, columnspan=2)
        ttk.Label(self.newrem_frm, text='Empfohlen')\
            .grid(row=1, column=self.recflag_col, padx=2, pady=1)
        ttk.Label(self.newrem_frm, text='Voreingestellt')\
            .grid(row=1, column=self.defaultflag_col, padx=2, pady=1)
        ttk.Label(self.newrem_frm, text='Bemerkungstext')\
            .grid(row=0, column=self.text_col, padx=2, pady=1, sticky='w')
        
        single_rem_manipulator = ChecklistLineEditor(
            master=self.newrem_frm,
            init_text=self.prefill[1],
            init_recflag=self.prefill[0],
            init_blacklist=[],
            init_whitelist=[]
        )
        
        _ = self.grid_single_rem_manipulator(single_rem_manipulator, 2)
        if single_rem_manipulator.wl_selector.get_list():
            ErrorWindow(self, 'wl_selector of new entry has entries. why? please investigate in console')
            IPS()
        self.rems['new'] = single_rem_manipulator

    def build_old_entries(self):
        gui_f.grid_forget_children(self.oldrem_frm)

        cl = self.get_checklist_slice()
        i = 0
        if cl.index.duplicated().any():
            raise IndexError(f'Mindestens eine variantnr doppelt vorhanden. Bitte in CSV heilen.')
        for posnr in cl.index:
            if posnr in self.delete_ogindex:
                continue
            try:
                single_rem_manip = self.rems[posnr]
            except KeyError:
                rem = cl.loc[posnr]
                single_rem_manip = ChecklistLineEditor(
                    master=self.oldrem_frm,
                    init_text=rem.fulltext,
                    init_recflag=rem.recommended_flag,
                    init_defflag=rem.default_state,
                    init_pos=posnr,
                    init_blacklist=dbf.db_blwl2list(rem.blacklist),
                    init_whitelist=dbf.db_blwl2list(rem.whitelist),
                )
            del_btn = ttk.Button(self.oldrem_frm, text='-', width=3,
                                 command=lambda nr=posnr: self.delete_rem(nr))
            del_btn.grid(column=self.delete_col, row=i, sticky='nw', padx=1, pady=1)
            i = self.grid_single_rem_manipulator(single_rem_manip, init_i=i)
            self.rems[posnr] = single_rem_manip
            
            ttk.Separator(self.oldrem_frm)\
                .grid(column=0, columnspan=5, row=i, padx=1, pady=1, sticky='ew')
            i += 1

    def delete_rem(self, ogindex):
        self.delete_ogindex.append(ogindex)
        del self.rems[ogindex]
        self.build_old_entries()


    def grid_single_rem_manipulator(self, single_rem_manipulator, init_i):
        '''places single_rem_manipulator in starting at row init_i.
        binds <Control-Return> to update_checklist for remtext'''
        row = init_i
        srm = single_rem_manipulator
        srm.posentry.grid(column=self.pos_col, row=row, padx=1, pady=1, sticky='new')
        srm.rec_flagcbx.grid(column=self.recflag_col, row=row, padx=1, pady=1, sticky='new')
        srm.def_flagcbx.grid(column=self.defaultflag_col, row=row, padx=1, pady=1, sticky='new')
        srm.remtext.grid(column=self.text_col, row=row, sticky='ew', padx=1, pady=1)
        row+=1
        srm.blwl_frm.grid(column=self.text_col, row=row, sticky='ew', padx=1, pady=1)
        row+=1
        srm.remtext.bind('<Control-Return>', lambda *_: self.update_checklist())
        return row        

    def get_checklist_slice(self):
        cl = dbf.load_checklist()
        try:
            return cl.loc[self.chapter, self.title]
        except KeyError:
            return pd.DataFrame()

    def update_checklist(self):
        # --> just delete slice of checklist entirely and rebuild with current state of this window
        cl = dbf.load_checklist()
        new_cl = dbf.multi_level_drop(cl, ['chapter', 'title'],
                                      [self.chapter, self.title])
        rem_DF = self.get_remarks_DF()
        
        if not rem_DF.empty:
            new_cl = pd.concat([new_cl, rem_DF]).sort_index()

        dbf.save_db(new_cl, 'checklist')
        self.RemEditor.update_checklist_frm()
        
        if self.wea is not None:
            self.wea.report.update_checklist()

        self.destroy()

    def get_remarks_DF(self):
        rem_dfs = []
        unused_oginds = list(self.rems.keys())
        used_positions = []
        used_oginds = []

        for og_ind in unused_oginds:
            rem = self.rems[og_ind]
            posnr = rem.posstrvar.get()
            if not posnr:
                continue
            pos = int(posnr)
            if pos not in used_positions: used_positions.append(pos)
            else:
                ErrorWindow(self, f'Pos. {pos} mehrfach vergeben. Bitte heilen.')
                return
            rem_df = rem.get_remark_DF()
            if not rem_df.empty: rem_dfs.append(rem.get_remark_DF())
            used_oginds.append(og_ind)
        
        for used_ogind in used_oginds: unused_oginds.remove(used_ogind)

        try: pos = np.array(used_positions).max() + 1
        except ValueError: pos = 1
        for og_ind in unused_oginds:
            rem = self.rems[og_ind]
            rem.posstrvar.set(str(pos))
            rem_df = rem.get_remark_DF()
            if not rem_df.empty: rem_dfs.append(rem.get_remark_DF())

        if not rem_dfs: return pd.DataFrame()
        df = pd.concat(rem_dfs)
        if color:= self.colorstrvar.get():
            df['bg_color'] = [self.colors_translator[color]]*len(df)
        else: df['bg_color'] = ['']*len(df)
        df = self.handle_duplicate_text(df)
        df = df.sort_index().reset_index(drop=True)
        df.index = df.index + 1 
        df.index = pd.MultiIndex.from_product([[self.get_chapter()],
                                               [self.get_title()],
                                               df.index], names =['chapter', 'title', 'variantnr'])
        return df

    def handle_duplicate_text(self, df):
        '''if there are duplicated fulltext entries, deletes one of the entries,
        concats black-/whitelist and chooses the worst flag for default and
        recommended flag
        NOTE: drops old rows on index, so make sure there are no duplicate indices!'''
        text_duplicated = df.duplicated(subset='fulltext')
        if not text_duplicated.any(): return df
        def concat_blwl(list_of_blwl):
            concatted = ''
            for blwl in list_of_blwl:
                if not blwl: continue
                entries = blwl.split(',')
                for entry in entries:
                    entry = entry.strip()
                    print(entry)
                    if not entry: continue
                    if f',{entry},' in concatted or f', {entry},' in concatted: continue
                    concatted += f'{entry}, '
            return concatted
        def get_first_non_empty(value_list):
            for val in value_list:
                if val: return val
            return ''

        duplicated_text = df[text_duplicated]['fulltext'].unique()
        new_rows = []
        for dup_text in duplicated_text:
            dup_rows = df[df.fulltext==dup_text]
            ind = dup_rows.index[0]
            new_bl = concat_blwl(dup_rows.blacklist.values)
            new_wl = concat_blwl(dup_rows.whitelist.values)
            new_def_state = gui_f.get_worst_flag(dup_rows.default_state.values)
            new_rec_flag = gui_f.get_worst_flag(dup_rows.recommended_flag.values)
            new_bg_color = get_first_non_empty(dup_rows.bg_color.values)
            author = config.default_author
            ind = pd.Index([int(ind)], name='variantnr')
            new_df = pd.DataFrame([[dup_text, new_bl, new_wl, new_def_state,
                                    new_rec_flag, new_bg_color, author]],
                                  index=ind,
                                  columns=['fulltext', 'blacklist', 'whitelist',
                                           'default_state', 'recommended_flag',
                                           'bg_color', 'author'])
            new_rows.append(new_df)
            df = df.drop(dup_rows.index)
        df = pd.concat([df, *new_rows])
        return df

    def get_chapter(self):
        return self.chapter
    def get_title(self):
        return self.title
    def get_addresstup(self):
        return (self.get_chapter(), self.get_title())
    def get_prettyaddress(self):
        return f'{self.get_chapter()}{db_split_char}{self.get_title()}'\
                    .replace(db_split_char, ' > ')

class ChecklistLineEditor():
    def __init__(self, master,
                 init_text: Optional[str]=None,
                 init_recflag: Optional[str]=None,
                 init_defflag: Optional[str]=None,
                 init_pos: Optional[str]=None,
                 init_blacklist: list=[],
                 init_whitelist: list=[],
                 ):
        self.master = master
        
        self.posstrvar = tk.StringVar(self.master)
        recflagstrvar = tk.StringVar(self.master)
        defaultflagstrvar = tk.StringVar(self.master)

        self.posentry = ttk.Entry(self.master, width=4,
                                   textvariable=self.posstrvar)
        self.rec_flagcbx = Flag_Combobox(self.master, recflagstrvar,
                                         include_blank=True)
        self.def_flagcbx = Flag_Combobox(self.master, defaultflagstrvar,
                                         include_blank=True)                                         

        self.remtext = tk.Text(self.master, height=5, wrap='word')
        self.remtext.configure(font=font.nametofont('TkDefaultFont'))

        self.blwl_frm = ttk.Frame(self.master)
        self.blwl_frm.grid_columnconfigure(1, weight=100)

        ttk.Label(self.blwl_frm, text='Blacklist: ')\
            .grid(row=0, column=0, padx=1, pady=1, sticky='w')
        ttk.Label(self.blwl_frm, text='Whitelist: ')\
            .grid(row=2, column=0, padx=1, pady=1, sticky='w')
        
        self.bl_selector = ListSelector(self.blwl_frm, init_blacklist)
        self.wl_selector = ListSelector(self.blwl_frm, init_whitelist)
        self.bl_selector.grid(row=0, column=1, padx=1, pady=1, sticky='ew')
        ttk.Separator(self.blwl_frm).grid(row=1, column=1, sticky='ew', pady=1, padx=1)
        self.wl_selector.grid(row=2, column=1, padx=1, pady=1, sticky='ew')

        self.init_prefill(init_text, init_recflag, init_defflag, init_pos)

    def init_prefill(self,
                     init_text: Optional[str]=None,
                     init_recflag: Optional[str]=None,
                     init_defflag: Optional[str]=None,
                     init_pos: Optional[str]=None,):
        if init_text: self.remtext.insert('1.0', gui_f.dbtext2displaytext(init_text))
        if init_recflag: self.rec_flagcbx.set(init_recflag)
        if init_defflag: self.def_flagcbx.set(init_defflag)
        if init_pos: self.posstrvar.set(str(int(init_pos)))

    def get_remark_DF(self):
        pos = self.posstrvar.get()
        text = self.get_remtext()
        if not text: return pd.DataFrame()
        ind = None
        if pos:
            try:
                ind = pd.Index([int(pos)], name='variantnr')
            except ValueError:
                ErrorWindow(self.master, f'Pos. muss Zahl sein, ist aber {pos}')
                return 'break'
        df = pd.DataFrame([[self.get_remtext(),
                            self.bl_selector.get_list(format='csv_str'),
                            self.wl_selector.get_list(format='csv_str'),
                            self.def_flagcbx.get(),
                            self.rec_flagcbx.get(),
                            '',
                            config.default_author
                            ]],
                            index=ind if ind is not None else [99999],
                            columns=['fulltext', 'blacklist', 'whitelist',
                                     'default_state', 'recommended_flag',
                                     'bg_color', 'author'])
        return df

    def get_remtext(self):
        raw_text = self.remtext.get('1.0', 'end-1c')
        text = raw_text.strip().replace('\n', db_split_char)
        return text



class ListSelector(ttk.Frame):
    def __init__(self, master, initial_list=[], expanded=False):
        super().__init__(master)
        self.expanded = False
        self.curr_selected = initial_list
        self.viewfrm = ttk.Frame(self)
        self.selectionfrm = ttk.Frame(self)
        
        self.oemvar = tk.StringVar(self)
        self.modelvar = tk.StringVar(self)
        self.towertypevar = tk.StringVar(self)
        for strvar in [self.oemvar, self.modelvar, self.towertypevar]:
            strvar.trace_add('write', lambda *e, var=strvar: self.add(var))

        self.create_comboboxes()
        self.update()

        self.viewfrm.pack(side='top', expand=True, fill='x', padx=1, pady=1)
        if expanded:
            self.expand_comboboxes()


    def get_list(self, format='normal'):
        '''format (str), normal or csv_str --> returns list as python list or as
        a string thats ready to put in csv and be interpreted by
        blacklist/whitelist parser'''
        if format == 'normal': return self.curr_selected
        elif format == 'csv_str':
            if not self.curr_selected: return ''
            return f'{', '.join(self.curr_selected)},'
        else: raise ValueError(f'format must be normal or csv_str, but is {format}')
        
    def create_comboboxes(self):
        oembox = ttk.Combobox(self.selectionfrm, textvariable=self.oemvar,
                              state='readonly', values=dbf.get_all('oem'))
        modelbox = ttk.Combobox(self.selectionfrm, textvariable=self.modelvar,
                                state='readonly', values=dbf.get_all('oemmodel'))
        tt_box = ttk.Combobox(self.selectionfrm, textvariable=self.towertypevar,
                              state='readonly', values=dbf.get_all('tower_type'))
        ttk.Label(self.selectionfrm, text='  + OEM ')\
            .pack(side='left', padx=1, pady=1)
        oembox.pack(side='left', padx=1, pady=1, fill='x', expand=True)
        ttk.Label(self.selectionfrm, text='  + Modell ')\
            .pack(side='left', padx=1, pady=1)
        modelbox.pack(side='left', padx=1, pady=1, fill='x', expand=True)
        ttk.Label(self.selectionfrm, text='  + Turmtyp ')\
            .pack(side='left', padx=1, pady=1)
        tt_box.pack(side='left', padx=1, pady=1, fill='x', expand=True)        

    def update(self):
        gui_f.delete_children(self.viewfrm)
        for entry in self.get_list():
            frm = ttk.Frame(self.viewfrm)
            frm.pack(side='left', padx=3, pady=1)
            ttk.Label(frm, text=entry).pack(side='top', pady=1, padx=1)
            ttk.Button(frm, text='-', width=3,
                       command=lambda ent=entry: self.delete(ent))\
                .pack(side='top', anchor='n')
        if not self.expanded:
            ttk.Button(self.viewfrm, text='+', width=3, 
                       command=self.expand_comboboxes)\
                .pack(side='right', padx=3, pady=1)

    def delete(self, entry):
        self.curr_selected.remove(entry)
        self.update()

    def add(self, strvar):
        selected_option = strvar.get()
        if selected_option not in self.get_list():
            self.curr_selected.append(selected_option)
        self.update()

    def expand_comboboxes(self):
        self.selectionfrm.pack(side='top', padx=1, pady=1, fill='x', expand=True)
        self.expanded = True
        self.update()

class Base_Data_Entering_Window(tk.Toplevel):
    '''Window for entering data into a mask
    mask stems from given dataframe
    data will be saved into given dataframe
    
    can be used to select wea oem, model, inspection year, id
    and generate data entering tmeplate from that,
    alternatively generate mask from a turbine present in the given dataframe
    
    (important) attributes: wea, db, data_dict, curr_oem_txtvar,
        curr_model_txtvar, curr_id_txtvar, curr_insp_year_txtvar,
    functions: toggle_copywea_selection, get_turbine_specifics

    '''
    def __init__(self, master,
                 db: pd.DataFrame | pd.Series,
                 build_database_entering_function,
                 insert_to_db_function,
                 wea: Optional[phys.Turbine]=None,
                 title: Optional[str]='Dateneingabe',
                 entering_height=400,
                 entering_width=None,
                 **kwargs) -> None:
        '''build_database_entering and ..._from_curr_turbine
        are the functions that are being called if the values of the 
        copy_wea_var (checkbox state) change or if the entries for the
        current turbine change'''
        self.db = db
        self.wea = wea

        # functions for builing the data entering mechanic
        self.build_database_entering = build_database_entering_function

        # function for database adding
        self.insert_to_db = insert_to_db_function
        
        self.data_dict = {}

        index_to_drop = list(self.db.index.names)
        index_to_drop.remove('model')
        index_to_drop.remove('oem')
        # available models in the database that shall be added to
        ind_available_models_in_db = (self.db
                                      .groupby(['oem', 'model'])
                                      .head(1)
                                      .sort_index()
                                      .index
                                      .droplevel(index_to_drop))
        self.available_models_in_db = dbf.multiindex2dict(
            ind_available_models_in_db
        )
        # available models in the "turbines" database
        self.available_models_in_wea_db = dbf.get_oem_model_dict()


        # setup basic container
        super().__init__(master, **kwargs)
        self.focus()
        self.minsize(800, 500)
        self.title(title)
        self.allfrm = ttk.Frame(self)
        self.allfrm.pack(fill='both', expand=True, padx=1.5, pady=1.5)
        self.allfrm.grid_columnconfigure(0, weight=1)

        # current wea is fixed, as window can only ever be accessed from a report
        self.curr_wea_frm = ttk.Frame(self.allfrm)
        self.curr_wea_frm.pack(fill='x', expand=False, padx=1, pady=1)

        self.infostrvar = tk.StringVar(self)
        infolbl = ttk.Label(self.curr_wea_frm, textvariable=self.infostrvar)
        self.set_infotxt()
        infolbl.pack(padx=1, pady=1)

        # Auswahl "Einträge übernehmen von"
        self.copy_wea_frm = ttk.Frame(self.allfrm)
        self.copy_wea_var = tk.IntVar(self, )
        self.copy_wea_checkbox = ttk.Checkbutton(self.copy_wea_frm,
                text='Eingabemuster übernehmen von anderer WEA',
                variable=self.copy_wea_var,
                command=self.toggle_copywea_selection)
        self.copy_wea_checkbox.grid(row=0, column=0, padx=1.5, pady=1.5, sticky='w')
        self.copy_wea_frm.pack(fill='x', padx=1, pady=1)
        self.copywea_selection_frm = ttk.Frame(self.copy_wea_frm)

        ttk.Separator(self.allfrm).pack(fill='x', padx=1, pady=1)

        # Eingabefelder
        border_of_data_entering_scrl = ttk.Frame(self.allfrm)
        border_of_data_entering_scrl.pack(fill='both', expand=True, padx=1, pady=1)
        data_entering_scrl = ScrollFrame(border_of_data_entering_scrl,
                                         def_height=entering_height,
                                         def_width=entering_width)
        data_entering_scrl.pack(fill='both', expand=True, padx=1, pady=1)
        self.data_entering_frm = data_entering_scrl.viewPort

        # OK, Abbrechen Button
        self.buttonsfrm = ttk.Frame(self.allfrm)
        self.buttonsfrm.pack(fill='x', padx=1, pady=1)

        self.okbtn = ttk.Button(self.buttonsfrm, text='OK',
                                command=self.insert_to_db)
        self.okbtn.pack(side='right', padx=5, pady=5)

    def set_infotxt(self):
        try: insp_year_prefill = self.wea.report.inspection.get_year()
        except: insp_year_prefill = None

        infostr = f'{self.wea.get('oem')} {self.wea.get('model')}: {self.wea.get('id')}'
        if insp_year_prefill:
            infostr += f' ({insp_year_prefill})'

        self.infostrvar.set(infostr)


    def toggle_copywea_selection(self):
        if self.copy_wea_var.get():
            self.copywea_selection_frm = ttk.Frame(self.copy_wea_frm)
            self.copywea_selection_frm.grid(row=1, column=0, padx=1.5, pady=1.5, sticky='ew')

            self.copywea_oem_txtvar = tk.StringVar(self, )
            self.copywea_model_txtvar = tk.StringVar(self, )

            self.copywea_oem_cbx = ttk.Combobox(self.copywea_selection_frm,
                                textvariable=self.copywea_oem_txtvar,
                                values=list(self.available_models_in_db.keys()),
                                state='readonly')
            self.copywea_model_cbx = Conditional_Combobox(self.copywea_selection_frm,
                                parent_strvar=self.copywea_oem_txtvar,
                                textvariable=self.copywea_model_txtvar,
                                value_dict=self.available_models_in_db,
                                state='readonly')
            self.copywea_model_txtvar.trace_add('write', self.build_database_entering)

            
            ttk.Label(self.copywea_selection_frm, text='Hersteller').grid(
                row=0, column=0, sticky='w', padx=5, pady=1.5)
            ttk.Label(self.copywea_selection_frm, text='Modell').grid(
                row=0, column=1, sticky='w', padx=5, pady=1.5)

            self.copywea_oem_cbx.grid(row=1, column=0, sticky='w', padx=5, pady=1.5)
            self.copywea_model_cbx.grid(row=1, column=1, sticky='w', padx=5, pady=1.5)

        else:
            self.build_database_entering()
            self.copywea_selection_frm.destroy()

    def get_turbine_specifics(self):
        '''return current wea's oem, model, id and inspection year'''
        oem = self.wea.get('oem')
        model = self.wea.get('model')
        turb_id = self.wea.get('id')
        insp_year = self.wea.report.inspection.get_year()
        for attr in [oem, model, turb_id]:
            if attr == None or pd.isna(attr) or attr == '':
                raise AttributeError('Turbine lacks oem, model or turbine_id.'
                                     f'Given: {oem}, {model}, {turb_id}')
        return oem, model, turb_id, insp_year

class TemperaturesEditor(Base_Data_Entering_Window):
    '''Window to enter temperature values 
    that are then added to a temperature database'''
    def __init__(self, master,
                 wea: Optional[phys.Turbine]=None,
                 on_close = lambda *_: None) -> None:
        
        self.temperatures_db = dbf.load_temperatures()
        self.wea = wea
        self.data_dict = {}
        self.on_close = on_close
        super().__init__(master, db=self.temperatures_db,
                         build_database_entering_function=self.build_database_entering,
                         insert_to_db_function=self.insert_to_db,
                         wea=self.wea,
                         entering_width=350,
                         title='Temperaturen')
        self.build_database_entering()
        self.prefill_from_db()
        

    def build_database_entering(self, *args, retain_textful_rows=True):
        self.data_dict = {}
        gui_f.delete_children(self.data_entering_frm)
        self.rowcounter = 0
        self.oem = self.copywea_oem_txtvar.get() if self.copy_wea_var.get() \
            else self.wea.get('oem')
        self.model = self.copywea_model_txtvar.get() if self.copy_wea_var.get() \
            else self.wea.get('model')
        
        if not self.oem or not self.model:
            raise AttributeError(f'WEA needs a defined oem and model, but has {self.oem} and {self.model}')

        temp_names = []
                
        try:
            temp_names.extend(self.temperatures_db
                              .loc[idx[self.oem, self.model]]
                              .index
                              .unique('name')
                              .to_list())
        except KeyError: temp_names.extend([])

        order = dbf.get_order('temperatures')
        ordered_names = []

        for name in order:
            if name not in temp_names: continue
            ordered_names.append(name)
            temp_names.remove(name)
        ordered_names.extend(temp_names)

        for name in ordered_names:
            lbl = ttk.Label(self.data_entering_frm, text=name)
            lbl.grid(row=self.rowcounter, column=0, padx=1.5, pady=1.5, sticky='w')
            var = tk.StringVar(self, )
            entry = ttk.Entry(self.data_entering_frm, textvariable=var)
            entry.grid(row=self.rowcounter, column=1, padx=1.5, pady=1.5, sticky='we')
            self.data_dict[name] = var
            self.rowcounter += 1

        self.data_entering_frm.grid_columnconfigure(1, weight=1)

        self.place_row_addbutton()
        self.regenerate_button = ttk.Button(self.data_entering_frm,
                                            text='Maske neu generieren',
                                            command=self.build_database_entering)
        self.regenerate_button.grid(row=self.rowcounter, column=0, sticky='w', 
                                padx=1.5, pady=1.5)

    def place_row_addbutton(self):
        '''places addbutton on self.data_entering_frm
        for placing entries for name/value at correct posotion
        according to self.rowcounter'''
        self.row_addbutton = ttk.Button(self.data_entering_frm, text='+',
                                        width=3, command=self.add_row)
        self.row_addbutton.grid(row=self.rowcounter, column=1, sticky='e', 
                                padx=1.5, pady=1.5) 

    def add_row(self):
        '''adds a row to the Dateneingabe frame
        the key in self.tempeature_txtvars will be a tk.StringVar, as will be
        the value in the dict'''
        # check if there are non-filled out added rows
        # newlines are structured {..., __newline__1: (name_StringVar, value_StringVar), }
        for key in self.data_dict.keys():
            if '__newline__' in key:
                val = self.data_dict[key]
                if not val[0].get():
                    ErrorWindow(self, 'Es gibt schon eine leere angefügte Zeile')
                    return
                
        self.row_addbutton.destroy()
        self.regenerate_button.grid_forget()
        
        namevar = tk.StringVar(self, )
        valvar = tk.StringVar(self, )

        nameentry = ttk.Entry(self.data_entering_frm, textvariable=namevar)
        nameentry.grid(row=self.rowcounter, column=0, padx=1.5, pady=1.5, sticky='w')

        valentry = ttk.Entry(self.data_entering_frm, textvariable=valvar)
        valentry.grid(row=self.rowcounter, column=1, padx=1.5, pady=1.5, sticky='we')
        
        self.data_dict[f'__newline__{self.rowcounter}'] = (namevar, valvar)
        self.rowcounter += 1

        self.place_row_addbutton()
        self.regenerate_button.grid(row=self.rowcounter, column=0, sticky='w', 
                                    padx=1.5, pady=1.5)
        
    def prefill_from_db(self):
        oem, model, turb_id, year = self.get_turbine_specifics()
        try: temps = self.temperatures_db.loc[idx[oem, model, :, turb_id, year]]
        except KeyError:
            try: temps = self.temperatures_db.loc[idx[oem, model, :, turb_id, int(year)]]
            except: return

        for name in temps.index:
            var = self.data_dict[name]
            if len(temps.loc[name]) > 1:
                ErrorWindow(self, f'Mehrere Einträge für Temperatur {name} '
                            'für diese WEA und Jahr/JahrInspektionsnummer',
                            on_end=self.focus)
                continue
            var.set(temps.loc[name].iloc[0])


    def collect_entries(self):
        '''builds a dict whose keys are the names of temperature values
        and whose values are the corresponding values
        
        truncates names, raises ValueError if value cant be converted to float
        
        return: dict like {'temp_name1': val1, 'temp_name2': val2, ...}'''
        entry_dict = {}
        for key in self.data_dict.keys():
            # handle manually added lines {..., __newline__1: (nameStrVar, valStrVar), ...}
            if '__newline__' in key:
                val = self.data_dict[key]
                key = val[0].get()
                if not key: # if the name entry is empty...
                    continue
                val = val[1].get().strip()
            else: val = self.data_dict[key].get().strip()
            # skip rows with empty fields for val
            if not val:
                continue

            key = key.strip()
            try:
                val = gui_f.str2float(val)
            except ValueError:
                raise ValueError(f'Eintrag "{key}" kann nicht als Zahl gedeutet werden: {val}')
            entry_dict[key] = val

        return entry_dict

    def insert_to_db(self):
        try:
            # get relevant values for index
            oem, model, turb_id, insp_year = self.get_turbine_specifics()
            # get entered temperatures
            entries = self.collect_entries()
        except ValueError as e:
            ErrorWindow(self, f'{e}')
            return

        new_temps = dbf.multi_level_drop(self.temperatures_db,
                        ['oem', 'model', 'turbine_id', 'insp_year'],
                        [oem, model, turb_id, insp_year])
        if len(entries) == 0:
            # delete all entries for this turbine and year
            dbf.save_db(new_temps, 'temperatures')
            self.destroy()
            return
        temp_order = dbf.get_order('temperatures')
        for name in entries.keys():
            if name not in temp_order:
                Orderer(self, 'temperature', name)
                return
        # create DataFrame
        rows = [pd.DataFrame([entry[1]],
                             index=pd.MultiIndex.from_tuples(
                                 [(oem, model, entry[0], turb_id, insp_year)],
                                 names=['oem', 'model', 'name', 'turbine_id', 'insp_year']
                             ),
                             columns=['value']) for entry in entries.items()]

        temps = pd.concat(rows)
        
        concat_temps = pd.concat([new_temps, temps])

        # checksum mit der länge der dataframes
        if len(concat_temps) == (len(new_temps) + len(temps)):
            dbf.save_db(concat_temps, 'temperatures')
        else:
            ErrorWindow(self, 'pd.concat hat irgendwelche Zeilen rausgeworfen. Temperaturen werden nicht in Datenbank gespeichert.')
            return
        # finally close window
        self.on_close()
        self.destroy()
        return

class Temperature_Comparison_Window(tk.Toplevel):
    def __init__(self, master, wea=None):
        super().__init__(master, height=600, width=1000)
        self.title('Temperaturen vergleichen')
        self.bind('<Escape>', lambda event: self.destroy()) 
        
        self.wea = wea
        self.data = dbf.load_temperatures()
        
        self.allfrm = ttk.Frame(self, width=1000, height=600)
        self.controlsfrm = ttk.Frame(self.allfrm)
        self.image_frm = ttk.Frame(self.allfrm)

        self.allfrm.pack(fill='both', expand=True)
        self.controlsfrm.pack(side='top', fill='x')
        ttk.Separator(self.allfrm).pack(side='top', fill='x', padx=5, pady=1)
        self.image_frm.pack(side='top', fill='both', expand=True)

        self.focus()
        self.oemvar = tk.StringVar(self, value=self.wea.get('oem') if self.wea else '')
        self.modelvar = tk.StringVar(self, value=self.wea.get('model') if self.wea else '')
        self.highlight_strvar = tk.StringVar(self, value=self.wea.get('id') if self.wea else '')

        self.build_structure()
        self.generate_image()

    def build_structure(self):
        available_models_ind = (self.data
                                .index
                                .droplevel(['name',
                                            'turbine_id',
                                            'insp_year'])
                                .unique())
        available_models = dbf.multiindex2dict(available_models_ind)
        oembox = ttk.Combobox(self.controlsfrm,
                              values=self.data.index.unique('oem').to_list(),
                              textvariable=self.oemvar, state='readonly')
        modelbox = Conditional_Combobox(self.controlsfrm, parent_strvar=self.oemvar,
                                        value_dict=available_models,
                                        textvariable=self.modelvar)
        identry = ttk.Entry(self.controlsfrm, textvariable=self.highlight_strvar)
        generate_button = ttk.Button(self.controlsfrm, text='Bild generieren',
                                     command=self.generate_image)
        
        oembox.pack(side='left', padx=1, pady=1)
        modelbox.pack(side='left', padx=1, pady=1)
        identry.pack(side='left', padx=1, pady=1)
        generate_button.pack(side='left', padx=1, pady=1)

        
    def generate_image(self):
        gui_f.delete_children(self.image_frm)
        wea_type = self.oemvar.get()
        model_startswith = self.modelvar.get()
        highlight_raw = self.highlight_strvar.get()
        if highlight_raw:
            hightlight_ids = [str(turb_id) for turb_id in\
                                    eval(f'{highlight_raw},')]
        else: hightlight_ids=[]
        
        fig, ax = plt.subplots(layout='constrained')
        try:
            ax = dbf.compare_temperatures(wea_type=model_startswith, ax=ax,
                                          highlight_ids=hightlight_ids,
                                          all_subtypes=True, oem=wea_type)
        except KeyError as e:
            ErrorWindow(self, 'Ausgewählte Konfiguration nicht in '
                              f'Temperaturdatenbank ({e.args[0]})', self.focus)
            return
        canvas = FigureCanvasTkAgg(fig, master=self.image_frm)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True, padx=3, pady=3)


class ComponentsEditor(Base_Data_Entering_Window):
    '''Window for entering the components of a turbine,
    stores the components in a database'''
    def __init__(self, master,
                 wea: Optional[phys.Turbine]=None,
                 on_close=lambda *_: None) -> None:

        self.components_db = dbf.load_parts()
        self.wea = wea
        self.component_list = []
        self.data_dict = {}
        self.on_close = on_close

        super().__init__(master, db=self.components_db,
                         build_database_entering_function=self.build_database_entering,
                         insert_to_db_function=self.insert_to_db,
                         wea=self.wea,
                         entering_width=1000,
                         title='Komponenten')
        self.build_database_entering(prefill=self.wea.get_parts())


    def build_database_entering(self, *args, prefill: Optional[dict]={},
                                retain_selected_rows=True):
        '''
        builds mask for entering component infos.
        prefill: dict with shape {partname: (oem, model, sn), ...}, is inserted
            in the corresponding rows
        retain_selected_rows: bool, if True forces to keep the currently selected
            rows hwen regenerating the mask, eg when template WEA changes.
            The row's values are replaced by the values of prefill, if prefill
            has values for the retained part's name
        '''
        # retain selected rows if retain_selected_rows = True
        textful_rows = self.collect_rows()
        self.data_dict = {}
        wea_oem, wea_model, _, _ = self.get_turbine_specifics()
        
        for key in textful_rows.keys():
            if retain_selected_rows and key not in prefill.keys():
                prefill[key] = textful_rows[key] 

        gui_f.delete_children(self.data_entering_frm)
        self.rowcounter = 0

        # build_oem/model is the variable name, because it can be either the
        # model/oem to copy from or the current wea (the wea that is actually
        # worked with)
        self.build_oem = self.copywea_oem_txtvar.get() if self.copy_wea_var.get() \
            else wea_oem
        self.build_model = self.copywea_model_txtvar.get() if self.copy_wea_var.get() \
            else wea_model
        # wea that is actually worked with defines dropdown options, even if
        # the data entering template is generated from another wea
        
        # if no wea is selected
        if not self.build_oem or not self.build_model:
            self.regenerate_button = ttk.Button(self.data_entering_frm,
                                    text='Maske neu generieren',
                                    command=self.build_database_entering)
            self.regenerate_button.grid(row=self.rowcounter, column=0, sticky='w', 
                                        padx=1.5, pady=1.5)
            self.place_row_addbutton()
            return
        
        # infer components from component database according to build_oem and build_model
        try:
            self.component_list = (self
                                   .components_db
                                   .loc[idx[self.build_oem, self.build_model]]
                                   .index
                                   .unique('component'))
            self.component_list = list(self.component_list)
        except KeyError:
            self.component_list = []

        # add prefill component names to component_list
        for component_name in prefill.keys():
            if component_name not in self.component_list:
                self.component_list.append(component_name)

        self.sort_component_list()
        
        for component_name in self.component_list:
            # add a row containing a checkbox, label, manufacturer and sn field
            # use only known components for current turbine type, not copy
            known_components = dbf.get_component_model_dict_for_turbine_type(
                    wea_oem, wea_model, component_name,
                )
            try:
                _prefill = prefill[component_name]
            except KeyError:
                _prefill = None
            component_selector = self.Component_Selection_Row(
                self.data_entering_frm,
                wea=self.wea,
                component_name=component_name,
                available_components_dict=known_components,
                prefill=_prefill,
            )
            self.place_row(component_selector)
            
            self.data_dict[component_name] = component_selector
            self.rowcounter += 1


        self.data_entering_frm.grid_columnconfigure([2, 3], weight=2)
        self.data_entering_frm.grid_columnconfigure(5, weight=3)
        self.data_entering_frm.grid_columnconfigure(4, minsize=10)
        self.data_entering_frm.grid_columnconfigure(5, minsize=10)
        self.place_row_addbutton()

    def place_row(self, row):
        row.lbl.grid(row=self.rowcounter, column=1,
                     padx=1.5, pady=1.5, sticky='w')
        row.manufacturer_combobox.grid(row=self.rowcounter, column=2,
                                       padx=5, pady=1.5, sticky='ew')
        row.model_combobox.grid(row=self.rowcounter, column=3,
                                padx=5, pady=1.5, sticky='ew')
        row.prefill_sn_btn.grid(row=self.rowcounter, column=4,
                                padx=4, pady=1.5, sticky='e')
        row.sn_entry.grid(row=self.rowcounter, column=5,
                          padx=5, pady=1.5, sticky='ew')
        row.delete_btn.grid(row=self.rowcounter, column=6,
                            sticky='e', padx=5, pady=1)

    def place_row_addbutton(self):
        '''places addbutton on self.data_entering_frm
        for placing entries for name/values at correct posotion
        according to self.rowcounter'''
        self.row_addbutton = ttk.Button(self.data_entering_frm, text='+',
                                        width=3, command=self.add_row)
        self.row_addbutton.grid(row=self.rowcounter, column=5, sticky='e', 
                                padx=1.5, pady=1.5)

    def add_row(self):
        '''adds a row to the Dateneingabe frame
        the key in self.tempeature_txtvars will be a tk.StringVar, as will be
        the value in the dict'''
        for key in self.data_dict.keys():
            val = self.data_dict[key]
            if val.is_empty_and_freely_selectable():
                ErrorWindow(self, 'Es gibt schon eine leere angefügte Zeile')
                return
                            
        self.row_addbutton.destroy()

        component_selector = self.Component_Selection_Row(
            self.data_entering_frm,
            wea=self.wea,
            component_name=None,
            available_components_dict={},
        )

        self.place_row(component_selector)
        
        self.data_dict[f'__newline__{self.rowcounter}'] = component_selector

        self.rowcounter += 1
        self.place_row_addbutton()

    def sort_component_list(self):
        '''sort values of self.component_list according to '''
        sorted_list = []
        component_order = dbf.get_order('components')
        # order component list according to component order
        for component in component_order:
            if component in self.component_list:
                sorted_list.append(component)
        
        # add components that are not in the component list
        for component in self.component_list:
            if component not in component_order:
                ErrorWindow(self, f'{component} ist nicht auf der "components_order" Liste in databases directory. Wird ans Ende der aktuellen Auswahl gesetzt.')
                sorted_list.append(component)
        self.component_list = sorted_list

    def _build_database_entering_from_curr_turbine(self, *args):
        '''rebuild database entering menus in case and copy_wea_var is 0
            - if there are selected rows, the rows are taken over to the
            newly generated entering mask, and prefilled
            - the remaining rows are inferred from turbine type'''
        # do nothing if wea is supposed to be copied
        # if self.copy_wea_var.get():
        #     return
        
        # take over already filled out lines
        prefill = self.collect_rows()
        self.build_database_entering(prefill=prefill)
    
    def collect_rows(self) -> dict:
        '''
        returns a dict with form {name1: (manufacturer, model, sn), name2: ...}
        in which the part infos of all rows that have the checkbox selected are
        contained.

        gives Error Message if name field is not filled out and skips that row.
        '''
        parts_dict = {}
        for _, component_row in self.data_dict.items():
            if component_row.is_empty(): continue
            name = component_row.get_name()
            component_info = component_row.get_part_info()
            parts_dict[name] = component_info
        return parts_dict

    def insert_to_db(self):
        try:
            # get relevant values for index
            oem, model, turb_id, insp_year = self.get_turbine_specifics()
            # get entered temperatures
            parts_dict = self.collect_rows()
        except ValueError as e:
            ErrorWindow(self, f'{e}')
            return
        # close without doing anything if no data is entered
        if len(parts_dict) == 0:
            self.destroy()
            return
        
        parts_order = dbf.get_order('components')
        for name in parts_dict.keys():
            if name not in parts_order:
                Orderer(self, 'component', name)
                return

        
        new_parts_df = (pd.DataFrame.from_dict(parts_dict,
                                           orient='index',
                                           columns=['component_oem',
                                                    'component_model',
                                                    'sn'])
                    .replace({'nan': np.nan})
                    .dropna(how='all'))
        ind = pd.MultiIndex.from_product([[oem],
                                          [model],
                                          new_parts_df.index,
                                          [turb_id],
                                          [insp_year]],
                                         names=['oem',
                                                'model',
                                                'component',
                                                'turbine_id',
                                                'insp_year'])
        new_parts_df.index = ind
        len_new = len(new_parts_df)

        parts_df = dbf.multi_level_drop(self.components_db, ['oem', 'model', 'turbine_id'],
                                        [oem, model, turb_id])
        len_old_without = len(parts_df)

        try:
            self.components_db = pd.concat([parts_df, new_parts_df],
                                           verify_integrity=True)
        except ValueError as err:      # raised when there are duplicate index values
            ErrorWindow(self, f'Unerwarteter Fehler: {err}')
            return
        
        # checksum mit der länge der dataframes
        if len(self.components_db) == (len_old_without + len_new):
            dbf.save_db(self.components_db, 'components')
        else:
            ErrorWindow(self, 'pd.concat hat irgendwelche Zeilen rausgeworfen. Komponenten werden nicht in Datenbank gespeichert.')
            return
        self.on_close()
        self.destroy()
        return


    class Component_Selection_Row():
        def __init__(self, master, wea,
                     component_name, available_components_dict,
                     prefill: Optional[tuple]=None) -> None:
            '''
            component_name: str or None, if given name will be label, else Entry
            prefill: (manufacturer, model, sn)
            '''
            self.wea = wea
            self.component_name = component_name
            self.available_components_dict = available_components_dict
            self.prefill = prefill
            self.prefill_sn_btn = ttk.Button(master, text='->', state='disabled',
                                            command=self.prefill_sn, width=3)
            self.delete_btn = ttk.Button(master, text='-', width=3,
                                         command=self.clear)

            self.component_name_var = tk.StringVar(master, value=self.component_name)
            self.manufacturer_var = tk.StringVar(master, )
            self.manufacturer_var.trace_add('write', self.toggle_prefill_sn_button)
            self.manufacturer_var.trace_add('write', self.toggle_delete_button)
            self.model_var = tk.StringVar(master, )
            self.model_var.trace_add('write', self.toggle_prefill_sn_button)
            self.model_var.trace_add('write', self.toggle_delete_button)
            self.sn_var = tk.StringVar(master, )
            self.sn_var.trace_add('write', self.toggle_delete_button)

            self.handle_init_prefill()
            
            if self.component_name:
                self.lbl = ttk.Label(master, textvariable=self.component_name_var)
            else:
                self.lbl = ttk.Entry(master, textvariable=self.component_name_var)

            self.manufacturer_combobox = ttk.Combobox(master,
                        textvariable=self.manufacturer_var,
                        values=list(self.available_components_dict.keys()))
            self.model_combobox = Conditional_Combobox(
                                    master,
                                    parent_strvar=self.manufacturer_var,
                                    textvariable=self.model_var,
                                    value_dict=self.available_components_dict,
                                    autoupdate_value=True,
                                )
            for box in [self.model_combobox, self.manufacturer_combobox]:
                box.unbind_class('TCombobox', '<MouseWheel>')
                box.unbind_class('TCombobox', '<ButtonPress-4>')
                box.unbind_class('TCombobox', '<ButtonPress-5>')
            self.sn_entry = ttk.Entry(master, textvariable=self.sn_var)
            self.toggle_prefill_sn_button()
            self.toggle_delete_button()

        # TODO: pretty todo; prefill SN for Parts like Enercon Rotor and Stator
        # that have neither oem nor model
        def handle_init_prefill(self, *args):
            if not self.prefill:
                return
            _prefill = []
            for value in self.prefill:
                if pd.isna(value):
                    _prefill.append('')
                    continue
                elif type(value)==str:
                    if value.lower() == 'nan' or value.lower() == 'none':
                        _prefill.append('')
                        continue
                _prefill.append(value)
            self.manufacturer_var.set(_prefill[0])
            self.model_var.set(_prefill[1])
            self.sn_var.set(_prefill[2])

        def toggle_prefill_sn_button(self, *_):
            if self.manufacturer_var.get() and self.model_var.get():
                oem = self.manufacturer_var.get()
                model = self.model_var.get()
                recent_sn_for_curr_part = dbf.get_sample_sn_from_part_model(
                    oem, model
                )
                if recent_sn_for_curr_part:
                    self.prefill_sn_btn.configure({'state': 'normal'})
                    return
            self.prefill_sn_btn.configure({'state': 'disabled'})

        def toggle_delete_button(self, *_):
            if self.is_empty(): self.delete_btn.configure(state='disabled')
            else: self.delete_btn.configure(state='normal')


        # TODO: pretty todo; prefill SN for Parts like Enercon Rotor and Stator
        # that have neither oem nor model
        def prefill_sn(self, *args):
            # prefill with last used SN for the current part
            oem = self.manufacturer_var.get()
            model = self.model_var.get()
            if not oem or not model: return

            recent_sn_for_curr_part = dbf.get_sample_sn_from_part_model(
                oem, model,
                self.wea.get('oem'), self.wea.get('model'), self.wea.get('id')
            )
            self.sn_var.set(recent_sn_for_curr_part)
                    
        def clear(self):
            self.model_var.set('')
            self.sn_var.set('')
            self.manufacturer_var.set('')

        def is_empty_and_freely_selectable(self, *_):
            '''return True for a completely empty row, else return False'''
            # check if component name is filled in
            if self.component_name_var.get():
                return False
            # check if any values (manufacturer, model, sn) are filled out
            if not self.is_empty(): return False
            return True

        def is_empty(self):
            if self.manufacturer_var.get() \
                or self.model_var.get() \
                or self.sn_var.get():
                return False
            return True

        def get_name(self):
            return self.component_name_var.get().strip()
        
        def get_manufacturer(self):
            manufacturer = self.manufacturer_var.get().strip()
            if pd.isna(manufacturer) or manufacturer == '':
                manufacturer = None
            return manufacturer
        
        def get_model(self):
            model = self.model_var.get().strip()
            if pd.isna(model) or model == '':
                model = None
            return model
        
        def get_sn(self):
            sn = self.sn_var.get().strip()
            try:
                if '"' in sn:
                    IPS()
            except: pass
            if pd.isna(sn) or sn == '':
                sn = None
            return sn

        def get_part_info(self) -> list:
            return [self.get_manufacturer(),
                    self.get_model(),
                    self.get_sn()]

class Extras_Selection_Frame(ttk.Frame):
    class Single_Extra_Row():
        def __init__(self, master, extra: str, wea: Optional[phys.Turbine]=None,
                     parent_turbine_viewer=None):
            '''extra: can be components, temperatures'''
            if extra not in ['components', 'temperatures', 'timeline']:
                raise ValueError("Extra must be in ['components', 'temperatures', 'timeline']")
            self.extra = extra
            self.wea = wea
            self.master = master
            self.parent_turbine_viewer = parent_turbine_viewer
            if self.extra == 'components':
                self.btntxt_prefix = 'Hauptkomponenten'
                entering_window_function = lambda: ComponentsEditor(master, self.wea, on_close=self.update)
            elif self.extra == 'temperatures':
                self.btntxt_prefix = 'Temperaturen'
                entering_window_function = lambda: TemperaturesEditor(master, self.wea, on_close=self.update)
            elif self.extra == 'timeline':
                self.btntxt_prefix = 'Zeitstrahl'
                entering_window_function = self.open_TLEditor

            self.has_data_var = tk.IntVar(master, value=0)
            self.btntxt = 'FEHLER'
            self.btn = ttk.Button(master, command=entering_window_function)
            self.has_data_checkbox = ttk.Checkbutton(master, state='disabled',
                                                     variable=self.has_data_var)
            self.update()

        def open_TLEditor(self):
            tl_data = self.wea.report.get_timeline_data()
            if not tl_data: tl_data = (None, None, None)
            TimelineEditor(self.parent_turbine_viewer, self.wea, *tl_data)

        def set_btntxt(self):
            if self.has_data_var.get(): suffix = 'ändern'
            else: suffix = 'eingeben'
            self.btntxt = f'{self.btntxt_prefix} {suffix}'
            self.btn.config(text=self.btntxt)

        def set_has_data(self):
            if self.extra == 'components':
                self.has_data_var.set(int(dbf.check_if_parts_in_db(self.wea)))
            elif self.extra == 'temperatures':
                self.has_data_var.set(
                    int(dbf.check_if_current_temperatures_in_db(self.wea,
                                                                self.wea.report.inspection))
                    )
            elif self.extra == 'timeline':
                self.has_data_var.set(1 if self.wea.report.get_timeline_data()\
                                      else 0)

        def update(self):
            self.set_has_data()
            self.set_btntxt()

        def __repr__(self):
            return repr(f'{self.extra} window selection row, state: {self.has_data_var.get()}')

    def __init__(self, master, wea, parent_turbine_viewer,
                 remarks: Optional[pd.DataFrame]=None):
        # TODO: copy to other WEAs Button where appropriate
        super().__init__(master)
        self.wea = wea
        self.parent_turbine_viewer = parent_turbine_viewer
        extras_frm = ttk.Frame(self)
        ttk.Label(extras_frm, text='Daten vorh.').grid(row=0, column=0,
                                                       padx=1, pady=1)

        self.extras = {}
        for i, extra in enumerate(['components', 'temperatures', 'timeline'], 1):
            extra_enterer = self.Single_Extra_Row(extras_frm, extra, self.wea,
                                                  self.parent_turbine_viewer)
            self.extras[extra] = extra_enterer
            extra_enterer.btn.grid(row=i, column=1, padx=1, pady=1, sticky='ew')
            extra_enterer.has_data_checkbox.grid(row=i, column=0,
                                                 padx=5, pady=1, sticky='n')
        compare_temps_btn = ttk.Button(extras_frm, text='Temperaturen vergleichen',
                                       command=lambda: Temperature_Comparison_Window(
                                                        self, self.wea
                                       ))
        compare_temps_btn.grid(row=2, column=2, padx=1, pady=1)

        manipulate_btn = ttk.Button(self, text='Metadaten ändern',
                                    command=lambda: MetadataEditor(
                                        self, self.wea.report.parent_project,
                                        self.wea.get('id'), on_close=self.update
                                    ))
        manipulate_btn.pack(side='top', padx=1, pady=1, fill='x')

        products_frm = ttk.Frame(self)
        products_frm.grid_columnconfigure(2, weight=100)
        self.compile_all_turbines_var = tk.IntVar(self, value=0)
        self.overview_style_var = tk.StringVar(self, value='DataDocx')
        all_turbines_cbx = ttk.Checkbutton(products_frm,
                                           text='Alle Anlagen',
                                           variable=self.compile_all_turbines_var)
        compile_btn = ttk.Button(products_frm, text='Prüfbericht erstellen', 
                                 command=self.compile_report)
        self.show_doc_btn = ttk.Button(products_frm, text='Öffnen',
                                       command=self.open_doc)
        reporad = ttk.Radiobutton(products_frm, text='DataDocx', value='DataDocx',
                                  variable=self.overview_style_var,
                                  command=self.update)
        inspectrad = ttk.Radiobutton(products_frm, text='Inspect', value='Inspect',
                                     variable=self.overview_style_var,
                                     command=self.update)
        overview_btn = ttk.Button(products_frm, text='Mängelliste erstellen',
                                  command=self.make_overview)
        self.show_overview_btn = ttk.Button(products_frm, text='Öffnen',
                                            command=self.open_overview)
        
        all_turbines_cbx.grid(row=0, column=0, sticky='w',
                              padx=1, pady=1, columnspan=2)
        compile_btn.grid(row=0, column=2, sticky='ew', padx=1, pady=1)
        self.show_doc_btn.grid(row=0, column=3, sticky='e', padx=1, pady=1)

        reporad.grid(row=1, column=0, padx=1, pady=1, sticky='w')
        inspectrad.grid(row=1, column=1, padx=1, pady=1, sticky='w')
        overview_btn.grid(row=1, column=2, padx=1, pady=1, sticky='ew')
        self.show_overview_btn.grid(row=1, column=3, padx=1, pady=1, sticky='e')

        scrl = ScrollFrame(self)
        self.todos_frm = scrl.viewPort
        self.todos_frm.grid_columnconfigure(0, weight=0)
        self.todos_frm.grid_columnconfigure(1, weight=100)

        extras_frm.pack(fill='x', padx=1, pady=1)
        ttk.Separator(self).pack(fill='x', padx=2, pady=2)
        products_frm.pack(fill='x', padx=1, pady=1)
        ttk.Separator(self).pack(fill='x', padx=2, pady=2)
        scrl.pack(fill='both', expand=True, padx=1, pady=2)

        self.parent_turbine_viewer.reset_info()
        self.update(remarks=remarks)
        
    def update(self, *_, remarks=None):
        if remarks is None: remarks=self.wea.report.get_remarks(ordered=False)
        for extra in self.extras.keys():
            self.extras[extra].update()
        self.toggle_show_btns()
        self.update_todos(remarks=remarks)

    def toggle_show_btns(self):
        path = os.getcwd()
        wea_id = self.wea.get('id')
        doc_name = f'{self.wea.report.get('id')}.docx'
        if doc_name in os.listdir(f'{path}/{wea_id}'):
            self.show_doc_btn.configure(state='normal')
        else: self.show_doc_btn.configure(state='disabled')
        
        excelname = self.get_excelname()
        if excelname in os.listdir(path):
            self.show_overview_btn.configure(state='normal')
        else: self.show_overview_btn.configure(state='disabled')


    def compile_report(self):
        weas = []
        if self.compile_all_turbines_var.get():
            weas = self.wea.report.parent_project.windfarm.get_setup_weas()
        else: weas = [self.wea]

        for wea in weas:
            try:
                wea.report.to_word()
                print(f'created report for {wea.get('id')}')
            except PermissionError:
                ErrorWindow(self, 'Dokument kann nicht neu generiert werden, da das Dokument gerade geöffnet ist.')
            except Exception as e:
                ErrorWindow(self, traceback.format_exc())
            finally:
                self.update()
    def open_doc(self):
        doc_name = self.wea.report.get('id')
        doc_path = f'{os.getcwd()}/{self.wea.get('id')}/{doc_name}.docx'
        os.startfile(doc_path)

    def make_overview(self):
        style = self.overview_style_var.get()
        try:
            self.wea.report.parent_project.create_overview(style=style)
        except PermissionError: ErrorWindow(self,
                                            'Bitte zunächst Überblick in Excel '
                                            'schließen.')
        self.update()
    def get_excelname(self):
        overview_style = self.overview_style_var.get()
        if overview_style == 'DataDocx': 
            return f'overview {self.wea.report.parent_project.get('name')}.xlsx'
        elif overview_style == 'Inspect':
            return f'overview {self.wea.report.parent_project.get('name')}_Insp.xlsx'
        else: raise ValueError(f'unknown value of overview style: {overview_style}')
    def open_overview(self):
        excelname = self.get_excelname()
        excelpath = f'{os.getcwd()}/{excelname}'
        os.startfile(excelpath)

    def update_todos(self, remarks=None):
        gui_f.delete_children(self.todos_frm)
        if remarks is None: remarks = self.wea.report.get_remarks(ordered=False)
        todos = []

        missing_report_data = []
        # TODO: get missing report metadata
        open_todos = self.wea.report.get_todo_count(remarks=remarks)
        missing_refs = self.wea.report.get_missing_refs(remarks=remarks)
        missing_imgs_from_report = self.wea.report.get_missing_images(where='report', remarks=remarks)
        missing_imgs = self.wea.report.get_missing_images(where='folder', remarks=remarks)
        imgs_multiply_used = self.wea.report.get_multiply_used_imgs(remarks=remarks)

        if open_todos: todos.append(f'Offene Todos ("???" in Bemerkung): {open_todos}')

        for img in missing_imgs:
            todos.append(f'Bild in Bericht aber nicht in 0-Fertig: {img}')
        
        if not self.wea.report.completely_signed():
            todos.append('Nicht alle Unterschriften sind komplett')

        missing_chapters_raw = self.wea.report.get_missing_chapters(remarks=remarks)
        chapter_counter = len(missing_chapters_raw)
        if chapter_counter > 3:
            missing_chapters = [f'{chapter_counter} Abschnitte nicht als fertig markiert']
        else:
            missing_chapters = []
            for chap in missing_chapters_raw:
                missing_chapters.append(f'Abschnitt nicht fertig: {chap}')
            
        todos.extend(missing_chapters)
        todos.extend(missing_refs)

        todos.extend(self.wea.report.inspection.get_missing_data())

        for extra in self.extras.keys():
            has_data = self.extras[extra].has_data_var.get()
            if not has_data:
                todos.append(f'Keine {extra} eingegeben')

        if missing_imgs_from_report:
            ttk.Label(self.todos_frm, text='Bilder in 0-Fertig, aber nicht in Bericht:')\
                .grid(row=0, column=1, padx=5, pady=0, sticky='ew')
            ImagesFrame(self.todos_frm, self.wea.id,
                       [img for img in missing_imgs_from_report])\
                .grid(row=1, column=1, padx=5, pady=0, sticky='ew')
            
        if imgs_multiply_used:
            ttk.Label(self.todos_frm, text='Mehrfach verwendete Bilder:')\
                .grid(row=2, column=1, padx=5, pady=0, sticky='ew')
            ImagesFrame(self.todos_frm, self.wea.id, imgs_multiply_used)\
                .grid(row=3, column=1, padx=5, pady=0, sticky='ew')
            
        for clearname in self.wea.get_missing_attributes(clearname=True):
            todos.append(f'WEA: {clearname} nicht angegeben')
        for clearname in self.wea.report.get_missing_attributes(clearname=True):
            todos.append(f'Bericht: {clearname} nicht angegeben')
        for clearname in self.wea.report.parent_project.get_missing_attributes(clearname=True):
            todos.append(f'Projekt: {clearname} nicht angegeben')
        
        for i, todo in enumerate(todos, 4):
            ttk.Label(self.todos_frm, text='-').grid(row=i, column=0,
                                            padx=5, pady=0, sticky='n')
            var = tk.StringVar(self, value=todo)
            MultilineLabel(self.todos_frm, textvar=var).grid(row=i, column=1,
                                            padx=5, pady=0, sticky='ew')




class Single_Remark_Displayer():
    '''button|flag|text
    class that contains a Multiline Label with text on the right
    and the corresponding flag (can be None) to the left
    and button to the left (can be None).'''
    def __init__(self, master, text: str,
                 remark_index=None, checklist_index=None,
                 flag: Optional[str]=None,
                 parent_remarkEditor=None,
                 images: Optional[list]=None,
                 commands: Optional[list]=['add'], textkwargs={},
                 **kwargs):
        '''
        master: Frame to put Single_remark into
        text: str to display
        address: address in remarks_db (relevant for delete and edit)
        flag: Flag to display, can be None
        commands: list, containing 'add', 'remove' or 'edit'
        '''
        self.master = master
        self.text = gui_f.dbtext2displaytext(text)
        self.flag = flag if flag is not None else ''
        self.remark_display_unit = master
        self.images = images
        self.remark_index = remark_index
        self.checklist_index = checklist_index
        self.parent_remarkEditor = parent_remarkEditor
        self.wea = self.parent_remarkEditor.wea

        # initialize StringVars
        self.flagstrvar = tk.StringVar(self.master, )
        self.flagstrvar.set(self.flag)
        self.textstrvar = tk.StringVar(self.master, )
        self.textstrvar.set(self.text)

        # display Text
        self.flaglbl = ttk.Label(self.master,
                                 textvariable=self.flagstrvar,
                                 anchor='center')
        self.textlbl = MultilineLabel(self.master, textvar=self.textstrvar)

        # display buttons
        self.addbutton = None
        self.editbutton = None
        self.deletebutton = None

        if 'add' in commands:
            self.place_addbutton()

        if 'edit' in commands:
            self.place_editbutton()

        if 'remove' in commands:
            self.place_removebutton()

    def add_remark(self):
        self.parent_remarkEditor.reset_searchtext()
        prefill_kwargs = {'remark_index': self.remark_index,
                          'checklist_index': self.checklist_index}
        RemarkEditor(self.parent_remarkEditor.master,
                    wea=self.wea,
                    parent_selection_frame=self.parent_remarkEditor,
                    **prefill_kwargs)

    def edit_remark(self):
        self.parent_remarkEditor.reset_searchtext()
        prefill_kwargs = {'remark_index': self.remark_index,
                          'checklist_index': self.checklist_index}
        RemarkEditor(self.parent_remarkEditor.master, 
                    wea=self.wea,
                    parent_selection_frame=self.parent_remarkEditor,
                    **prefill_kwargs)

    def remove_remark(self):
        self.parent_remarkEditor.reset_searchtext()
        self.wea.report.remove_remark(address=self.remark_index[-2],
                                      timestamp=self.remark_index[-1])
        self.wea.report.update_compressed_images(check_only=self.images)
        self.parent_remarkEditor.build_selection_body()
    
    def place_addbutton(self):
        self.addbutton = ttk.Button(self.master,
                                    command=self.add_remark,
                                    text='+',
                                    width=3)

    def delete_addbutton(self):
        if self.addbutton is not None:
            self.addbutton.grid_forget()
            del self.addbutton
            self.addbutton = None

    def place_editbutton(self):
        self.editbutton = ttk.Button(self.master,
                                     command=self.edit_remark,
                                     text='...',
                                     width=3)

    def delete_editbutton(self):
        if self.editbutton is not None:
            self.editbutton.grid_forget()
            del self.editbutton
            self.editbutton = None

    def place_removebutton(self):
        self.removebutton = ttk.Button(self.master,
                                       command=self.remove_remark,
                                       text='-',
                                       width=3)

    def delete_removebutton(self):
        if self.removebutton is not None:
            self.removebutton.grid_forget()
            del self.removebutton
            self.removebutton = None


class RemarkSelector(ttk.Frame):
    '''frame where the selectable remarks are shown
    contains a searchbar which searches title and fulltext of all checklist
    elements in the current chapter/section'''
    def __init__(self, master, address, wea, parent_turbine_viewer=None,
                 remarks=None) -> None:
        super().__init__(master)

        self.address = address
        self.wea = wea
        self.master = master
        self.chapter_done = tk.IntVar(self.master,
                                      value=wea.report.get_chapter_done(address))
        self.chapter_done.trace_add('write', self.toggle_chapter_done)
        self.show_all = tk.IntVar(self, value=0)
        self.show_all.trace_add('write', self.build_selection_body)
        self.parent_turbine_viewer = parent_turbine_viewer
        
        # add searchbar
        searchbar_frm = ttk.Frame(self)
        searchbar_frm.pack(side='top', fill='x', padx=2, pady=2)
        searchbar_frm.grid_columnconfigure(1, weight=2)
        searchbarlbl = ttk.Label(searchbar_frm, text='Suche:       ')
        searchbarlbl.grid(row=1, column=0, sticky='w', padx=2, pady=1)
        self.searchtextvar = tk.StringVar(self, )
        self.searchtextvar.trace_add('write', lambda *_: self.build_selection_body()\
                                     if not self.searchtextvar.get() else None)
        searchbar = ttk.Entry(searchbar_frm,
                              textvariable=self.searchtextvar)
        searchbar.bind('<Return>', self.build_selection_body)
        searchbar.grid(row=1, column=1, sticky='ew', padx=2, pady=1)
        
        self.scrl = ScrollFrame(self)
        self.scrl.pack(side='top', fill='both', expand=True)
        self.remarks_selection_frm = ttk.Frame(self.scrl.viewPort)
        self.remarks_selection_frm.pack(side='top', anchor='n', fill='both', expand=True)

        self.build_selection_body(remarks=remarks)

        # pack "done" checkbox
        controlsfrm = ttk.Frame(self)
        controlsfrm.pack(fill='x', side='bottom')

        is_done_checkbutton = ttk.Checkbutton(controlsfrm,
                                              variable=self.chapter_done,
                                              text='Abschnitt fertig')
        is_done_checkbutton.grid(column=0, row=0, sticky='e', padx=2, pady=2)
        self.parent_turbine_viewer.parent_mainwindow.bind('<Control-Return>',
                                        lambda e: is_done_checkbutton.invoke())
        show_all_checkbutton = ttk.Checkbutton(controlsfrm,
                                               variable=self.show_all,
                                               text='Alle bekannten Titel')
        show_all_checkbutton.grid(column=1, row=0, padx=2, pady=2, sticky='e')

    def build_selection_body(self, *_, remarks=None, set_info=True):
        '''places the remark display units
        given remarks only used if chapter is done'''
        
        gui_f.delete_children(self.remarks_selection_frm)
        if self.chapter_done.get():
            self.build_done_view(remarks=remarks)
            return
        self.poscol_placed = False
            
        # cases: 'park', 'selected_only'
        marker_col=0
        addbtn_col = 1
        editbtn_col = 2
        deletebtn_col = 3
        flag_col = 4
        fulltext_col = 5
        manualpos_col = 7
        park_remcounter_col = 8

        def place_poscol():
            ttk.Label(self.remarks_selection_frm, text='Pos.')\
                .grid(row=0, column=manualpos_col, padx=1, pady=1)
            self.poscol_placed = True

        checklist = self.wea.report.get_checklist()

        
        self.parent_turbine_viewer.parent_mainwindow.park_cbx.configure(state='normal')
        self.remarks_selection_frm.grid_columnconfigure(fulltext_col, weight=100)
        for col in [marker_col, addbtn_col, editbtn_col, deletebtn_col,
                    flag_col, manualpos_col, park_remcounter_col]:
            self.remarks_selection_frm.grid_columnconfigure(col, weight=0)

        i = 0
        ttk.Label(self.remarks_selection_frm, text='Park')\
            .grid(row=i, column=park_remcounter_col, padx=1, pady=1)
        i += 1

        self.build_remarks_dict()
        curr_flags = self.get_flags_from_remarks_dict(self.wea.id)
        self.update_remindicator(flags=curr_flags)

        for title in self.remarks_dict.keys():
            curr_remarks = self.remarks_dict[title]
            art = RemarkArt(self.remarks_selection_frm,
                            self.wea.report.parent_project, curr_remarks,
                            self.parent_turbine_viewer.parent_mainwindow.update_weapage)
            search = self.searchtextvar.get()
            
            search_in_cl = False

            checklist_ind = (self.address, title[1:-1] if title.startswith('(')\
                                and title.endswith(')') else title)
            try:
                cl_entry = checklist.loc[checklist_ind]
                for ind in cl_entry.index:
                    text = cl_entry.loc[ind].fulltext
                    if search.lower() in text.lower():
                        search_in_cl = True
                try: col = cl_entry['bg_color'].dropna().unique()[0]
                except IndexError: raise KeyError() # go to pass below...
                tk.Frame(self.remarks_selection_frm, background=col, width=5)\
                            .grid(row=i, column=marker_col, padx=1, pady=4, sticky='ns')
            except KeyError: pass

            search_in_rems = False
            for wea_id in curr_remarks.keys():
                wea_remarks = curr_remarks[wea_id]
                for _, fulltext, _, _, _ in wea_remarks:
                    if search.lower() in fulltext.lower():
                        search_in_rems = True

            if search.lower() not in title.lower()\
                and not search_in_rems\
                    and not search_in_cl:
                continue

            title_srd = Single_Remark_Displayer(
                self.remarks_selection_frm, text=title,
                parent_remarkEditor=self, checklist_index=checklist_ind,
                commands=['add'])
            title_srd.addbutton.grid(row=i, column=addbtn_col,
                                     padx=1, pady=1, sticky='nw')
            title_srd.textlbl.grid(row=i, column=fulltext_col,
                                   padx=1, pady=1, sticky='new')

            art.grid(row=i, column=park_remcounter_col, padx=1, pady=2)            
            i += 1
            
            for curr_wea_rem in curr_remarks[self.wea.id]:
                flag, fulltext, images, pos, index = curr_wea_rem

                sr = Single_Remark_Displayer(self.remarks_selection_frm,
                                             text=fulltext, flag=flag,
                                             remark_index=index,
                                             parent_remarkEditor=self,
                                             images=images,
                                             checklist_index=checklist_ind,
                                             commands=['edit', 'remove'])
                sr.editbutton.grid(row=i, column=editbtn_col, padx=1, pady=1, sticky='nw')
                sr.removebutton.grid(row=i, column=deletebtn_col, padx=1, pady=1, sticky='nw')
                sr.flaglbl.grid(row=i, column=flag_col, padx=4, pady=1, sticky='nw')
                sr.textlbl.grid(row=i, column=fulltext_col,
                                padx=1, pady=1, sticky='new')
                
                if pos:
                    place_poscol()
                    ttk.Label(self.remarks_selection_frm, text=pos)\
                        .grid(row=i, column=manualpos_col, padx=1, pady=1, sticky='n')
                i += 1
                if isinstance(images, str):
                    images = eval(images)
                if not images: continue
                images_frm = self.get_image_frame(self.wea.id, images)
                images_frm.grid(row=i, column=fulltext_col, sticky='ew',
                                padx=1, pady=1)
                i += 1
            
            if not self.parent_turbine_viewer.parent_mainwindow.show_park_var.get():
                sep = ttk.Separator(self.remarks_selection_frm, orient='horizontal')
                sep.grid(row=i, column=0, columnspan=9, padx=1, pady=1, sticky='ew')
                i += 1
                continue
            
            old_wea_id = None
            for wea_id in curr_remarks.keys():
                if wea_id == self.wea.id:
                    continue
                for curr_wea_rem in curr_remarks[wea_id]:
                    flag, fulltext, images, _, index = curr_wea_rem
                    if not wea_id == old_wea_id:
                        ttk.Label(self.remarks_selection_frm, text=wea_id)\
                            .grid(row=i, column=addbtn_col, columnspan=3,
                                    padx=1, pady=1, sticky='nw')
                    ttk.Label(self.remarks_selection_frm, text=flag)\
                        .grid(row=i, column=flag_col, padx=4, pady=1, 
                              sticky='nw')
                    fulltext_var = tk.StringVar(self, value=gui_f.dbtext2displaytext(fulltext))
                    MultilineLabel(self.remarks_selection_frm,
                                    textvar=fulltext_var)\
                        .grid(row=i, column=fulltext_col, padx=1, pady=1,
                              sticky='new')
                    old_wea_id = wea_id
                    i += 1
                    if isinstance(images, str):
                        images = eval(images)
                    if not images: continue
                    images_frm = self.get_image_frame(wea_id, images, size='small')
                    images_frm.grid(row=i, column=fulltext_col, sticky='ew',
                                    padx=1, pady=1)
                    i += 1

            sep = ttk.Separator(self.remarks_selection_frm, orient='horizontal')
            sep.grid(row=i, column=0, columnspan=9, padx=1, pady=1, sticky='ew')
            i += 1
                # add add_remark button
        add_remark_btn = ttk.Button(self.remarks_selection_frm,
                                    text='+', width=3,
                                    command=lambda: RemarkEditor(
                                        self.master, self.wea, (self.address, 'XXX'),
                                        parent_selection_frame=self))
        add_remark_btn.grid(column=addbtn_col, row=i, padx=1, pady=1, sticky='nw')
        if set_info: self.set_infotxt()

    def set_infotxt(self, remarks=None):
        self.parent_turbine_viewer.reset_info()
        if self.address == 'Prüfergebnis|Fazit':
            if remarks is None: remarks = self.wea.report.get_remarks(ordered=False)
            flag_count = remarks.groupby('flag').count()['fulltext']
            swarmplots = self.get_PIE_swarmplots()
            for i, flag in enumerate(list('PEI')):
                try: count = flag_count[flag]
                except KeyError: count = 0

                counter_lbl = ttk.Label(self.parent_turbine_viewer.infofrm,
                                        text=f'{'    ' if i>0 else ''} {flag}: {count}')
                counter_lbl.pack(side='left', padx=1, pady=1)

                fig = swarmplots[flag]
                canvas = FigureCanvasTkAgg(fig, master=self.parent_turbine_viewer.infofrm)
                canvas.draw()
                canvas.get_tk_widget().pack(side='left', padx=3, pady=1)


    def get_PIE_swarmplots(self):
        all_remarks = dbf.load_remarks()
        swarmplots = {}

        oem = self.wea.get('oem')
        turb_id = self.wea.get('id')
        model = self.wea.get('model')
        insp_type = self.wea.report.inspection.get('kind')

        for flag in list('PEI'):
            swarmplots[flag] = dbf.get_flag_stripplot(flag, oem, model, insp_type,
                                                     turb_id, all_remarks)
            
        return swarmplots

    def toggle_chapter_done(self, *_):
        remarks = self.wea.report.get_remarks(ordered=False)
        if self.chapter_done.get():
            self.wea.report.mark_chapter_done(self.address)
        else:
            self.wea.report.mark_chapter_undone(self.address)
        self.build_selection_body(remarks=remarks)
        self.parent_turbine_viewer.progbar.update(self.address)
        self.parent_turbine_viewer.parent_mainwindow\
            .update_wea_selector(self.wea.id, active=True, remarks=remarks)

    def build_done_view(self, remarks=None):
        self.parent_turbine_viewer.parent_mainwindow.park_cbx.configure(state='disabled')
        self.set_infotxt(remarks=remarks)
        self.scrl.gotoTop()

        if remarks is None: remarks = self.wea.report.get_remarks(ordered=False)
        remarks = dbf.get_remarks_of_chapter(remarks=remarks, chapter=self.address)
        
        flag_col = 0
        fulltext_col = 1

        self.remarks_selection_frm.grid_columnconfigure(fulltext_col, weight=100)

        if remarks is None:
            self.update_remindicator(rems=[])
            return
        try:
            remarks = self.wea.report.order_remarks(remarks)
        except Exception as e:
            ErrorWindow(self, e, lambda: self.chapter_done.set(0))
        self.update_remindicator(rems=remarks)
        
        i = 0
        for ind in remarks.index:
            rem = remarks.loc[ind]
            text = gui_f.dbtext2displaytext(rem.fulltext)
            images = rem.image_names
            if isinstance(images, str):
                images = eval(images)

            flaglbl = ttk.Label(self.remarks_selection_frm,
                                text=rem.flag,
                                width=5)
            textlbl = MultilineLabel(self.remarks_selection_frm,
                                     textvar=tk.StringVar(
                                         self,
                                         value=text))

            flaglbl.grid(row=i, column=flag_col, padx=5, pady=1, sticky='new')
            textlbl.grid(row=i, column=fulltext_col, padx=1, pady=1, sticky='ew')
            i += 1
            if not images: continue
            images_frm = self.get_image_frame(self.wea.id, images)
            images_frm.grid(row=i, column=fulltext_col, padx=1, pady=1, sticky='ew')
            i += 1

    def update_remindicator(self, rems=None, flags=None):
        if rems is None and flags is None:
            ValueError('Either rems or flags must be given.')
        if flags is not None and rems is not None:
            ValueError('Give only rems or flags.')
        if flags is None:
            if len(rems) == 0: flags = [] # catches rems is list/tuple
            else: flags = rems.flag.to_list()
        self.parent_turbine_viewer.progbar.update_remindicator(self.address, flags)

    def get_flags_from_remarks_dict(self, wea_id):
        flags = []
        for _, title_remarks in self.remarks_dict.items():
            wea_remarks = title_remarks[wea_id]
            for remark in wea_remarks:
                flag = remark[0]
                flags.append(flag)

        return flags

    def reset_searchtext(self):
        if self.searchtextvar.get():
            self.searchtextvar.set('')

    def build_remarks_dict(self):
        wea_ids = self.wea.report.parent_project.windfarm.get_setup_wea_ids()
        sec_len = len(self.address.split(db_split_char))

        all_park_rems = self.wea.report.parent_project.get_all_remarks()

        correct_chap = []

        try:
            titles = (self.wea.report.get_checklist()
                      .loc[[self.address]]
                      .index.unique('title').to_list())
        except KeyError:
            titles = []

        if not all_park_rems.empty:
            correct_superchap = (all_park_rems
                                .index
                                .get_level_values('address')
                                .str
                                .startswith(f'{self.address}{db_split_char}'))
            park_correct_superchap = all_park_rems.loc[correct_superchap]

            for address in park_correct_superchap.index.get_level_values('address'):
                address_as_list = address.split(db_split_char)
                title = address_as_list[-1]
                if len(address_as_list[:-1]) == sec_len:
                    correct_chap.append(True)
                    if title not in titles: titles.append(title)
                    continue
                correct_chap.append(False)            
            
            park_rems = park_correct_superchap.loc[correct_chap]

        # sort titles
        titles_order = dbf.get_order('chapters')[self.address]
        ordered_titles = []
        # order titles according to ordered titles list in database folder
        all_titles = self.show_all.get()
        for title in titles_order:
            if title in titles:
                ordered_titles.append(title)
                titles.remove(title)
                continue
            if all_titles: ordered_titles.append(f'({title})')
        # order remaining (so far unordered) titles
        if titles:
            title = titles[0]
            Orderer(self, 'title', title, self.address, on_close=self.build_selection_body)
        else: titles = ordered_titles

        remarks_dict = {}
        for title in titles:
            address = f'{self.address}|{title}'
            curr_title_dict = {}
            
            for wea_id in wea_ids:
                curr_wea_title_remarks = []
                if all_park_rems.empty:
                    curr_title_dict[wea_id] = curr_wea_title_remarks
                    remarks_dict[title] = curr_title_dict
                    continue

                try: 
                    curr_remarks = park_rems.loc[idx[:, wea_id, :, :, address], :]
                    for ind in curr_remarks.index:
                        rem = curr_remarks.loc[ind]
                        curr_wea_title_remarks.append((rem.flag,
                                                       rem.fulltext,
                                                       rem.image_names,
                                                       dbf.rempos2str(rem.position),
                                                       ind))
                except KeyError:
                    pass
                curr_title_dict[wea_id] = curr_wea_title_remarks
            remarks_dict[title] = curr_title_dict
        
        self.remarks_dict = remarks_dict
    
    def get_image_frame(self, wea_id, images, size='normal') -> ttk.Frame:
        return ImagesFrame(self.remarks_selection_frm, wea_id,
                          images, image_size=size)
        
class ImagesFrame(ttk.Frame):
    def __init__(self, master, wea_id, images, image_size='normal'): # height_scrl is a dirty fix, better would be to let the image height determine the height_scrl automatically
        '''image_size: str, normal or small'''
        if image_size == 'normal':
            width = 160
        elif image_size == 'small':
            width = 120
        else: raise ValueError(f'image_size must be "normal" or "small", but is {image_size}')
        super().__init__(master)
        if not images:
            return
        cwd = os.getcwd()

        self.grid_columnconfigure(0, weight=100)

        if len(images) > 3:
            scrl = ScrollFrame(self, orient='horizontal', use_mousewheel=False)
            scrl.grid(row=0, column=0, sticky='nsew')
            frm = scrl.viewPort
        else:
            frm = self

        heights = []
        for i, image in enumerate(images):
            image_path = f'{cwd}/{wea_id}/0-Fertig/{image}'
            try:
                pic = gui_f.get_tkinter_pic(image_path, master=self,
                                            width_pxl=width if 'timeline' not in image else width*2)
                pic_lbl = ttk.Label(frm, image=pic)
                pic_lbl.image = pic
                pic_lbl.bind('<Button-1>', lambda e, path=image_path: gui_f.open_image(path))
                heights.append(pic.height())
            except FileNotFoundError:
                pic_lbl = tk.Label(frm, text=image, bg='red')
            pic_lbl.pack(side='left', padx=3, pady=3)

        if len(images) > 3:
            scrl.canvas.configure(height=max(heights)+5)


class TurbineViewer(ttk.Frame):
    '''das ist die klasse in der die weas komplett angezeigt werden, also
    Kapitel, bemerkungen, bilder
    '''
    def __init__(self, parent, project: phys.Project, wea_id: str,
                active_section=None, parent_mainwindow=None, remarks=None,
                *args, **frame_kwargs) -> None:
        super().__init__(parent, **frame_kwargs)
        self.parent = parent
        self.parent_mainwindow = parent_mainwindow
        self.project = project
        self.wea = self.project.windfarm.get_wea(wea_id)
        self.set_chapters(remarks=remarks)
        self.active_section = active_section
        if not self.active_section or self.active_section not in self.chapters:
            self.active_section = self.wea.report.get_first_undone_chapter()
        self.remark_sel_frame = None
        self.extras_frame = None

        infofrm = ttk.Frame(self)
        infofrm.grid_columnconfigure(1, weight=100)
        infofrm.grid(column=0, row=0, sticky='ew')
        self.chapterstrvar = tk.StringVar(self)
        self.renamevar = tk.StringVar(self)

        chaptercbx = ttk.Combobox(infofrm, textvariable=self.chapterstrvar,
                values=[chap.replace('|', ' > ') for chap in self.chapters],
                state='readonly')
        renamelbl = ttk.Label(infofrm, textvariable=self.renamevar)
        self.chapter_rename_btn = ttk.Button(infofrm, text='Umbenennen',
                                        command=self.open_chapter_renamer)
        
        chaptercbx.grid(row=0, column=1, padx=1, pady=1, sticky='ew')
        renamelbl.grid(row=0, column=2, padx=1, pady=1, sticky='w')
        self.chapter_rename_btn.grid(row=0, column=3, padx=1, pady=1, sticky='w')
        self.set_chapterstrvar()

        self.reset_info()
        self.remark_sel_frame_frm = ttk.Frame(self,)
        self.remark_sel_frame_frm.grid(row=2, column=0, sticky='nesw',
                                       padx=1, pady=1)
        self.grid_rowconfigure(2, weight=2)
        self.grid_columnconfigure(0, weight=100)

        controls_frm = ttk.Frame(self)
        controls_frm.grid(row=3, column=0, sticky='ew')
        self.prevbtn = ttk.Button(controls_frm, text='<-',
                                  command=self.prev_chapter, width=3)
        self.nextbtn = ttk.Button(controls_frm, text='->',
                                  command=self.next_chapter, width=3)
        progbar_frm = ttk.Frame(controls_frm)
        self.progbar = Chapterbar(progbar_frm, self.wea.report,
                                  self.change_chapter, self.active_section,
                                  remarks=remarks)
        self.progbar.pack(fill='both', expand=True, padx=5, pady=3)

        self.prevbtn.grid(row=0, column=0, sticky='w', padx=2, pady=2)
        self.nextbtn.grid(row=0, column=1, sticky='w', padx=2, pady=2)

        controls_frm.grid_columnconfigure(2, weight=100)
        progbar_frm.grid(row=0, column=2, sticky='nsew', padx=3, pady=2)


        self.build_structure(remarks=remarks)
        self.parent_mainwindow.bind('<Control-Left>', self.prev_chapter)
        self.parent_mainwindow.bind('<Control-Right>', self.next_chapter)
        # rebuild if changed through combobox
        self.chapterstrvar.trace_add('write',
                                     lambda *args: self.change_chapter())


    def build_structure(self, remarks=None):
        self.toggle_rename_btn()
        self.toggle_arrowbuttons(remarks=remarks)
        self.set_chapterstrvar()
        self.set_rename_info()
        self.reset_info()
        if self.remark_sel_frame is not None:
            self.remark_sel_frame.destroy()
        if self.extras_frame is not None:
            self.extras_frame.destroy()
        if self.active_section == 'Extras':
            self.parent_mainwindow.park_cbx.configure(state='disabled')
            self.extras_frame = Extras_Selection_Frame(self.remark_sel_frame_frm,
                                                       self.wea,
                                                       parent_turbine_viewer=self,
                                                       remarks=remarks)
            self.extras_frame.pack(fill='both',
                                   padx=2, pady=2,
                                   expand=True, side='top')
            return
        self.parent_mainwindow.park_cbx.configure(state='normal')
        curr_traces = self.parent_mainwindow.show_park_var.trace_info()
        for mode, trace in curr_traces:
            self.parent_mainwindow.show_park_var.trace_remove(mode[0], trace)

        self.remark_sel_frame = RemarkSelector(self.remark_sel_frame_frm,
                                               self.active_section,
                                               self.wea,
                                               parent_turbine_viewer=self,
                                               remarks=remarks)
        self.parent_mainwindow.show_park_var.trace_add('write',
                                       self.build_selfrm_from_show_park_var)
        self.remark_sel_frame.pack(fill='both',
                                    padx=2, pady=2,
                                    expand=True, side='top')

    def reset_info(self):
        try:
            gui_f.delete_children(self.infofrm)
            self.infofrm.destroy()
        except AttributeError: pass
        self.infofrm = ttk.Frame(self)
        self.infofrm.grid(row=1, column=0, sticky='ew', padx=1, pady=1)
        
    def build_selfrm_from_show_park_var(self, *args):
        self.remark_sel_frame.scrl.gotoTop()
        self.remark_sel_frame.build_selection_body()

    def toggle_arrowbuttons(self, remarks=None):
        if self.active_section == self.wea.report.get_chapters(remarks=remarks)[-1]:
            self.nextbtn.configure(state='disabled')
        else:
            self.nextbtn.configure(state='normal')

        if self.active_section == 'Extras':
            self.prevbtn.configure(state='disabled')
        else:
            self.prevbtn.configure(state='normal')

    def set_chapters(self, remarks=None):
        self.chapters = ['Extras']
        self.chapters.extend(self.wea.report.get_chapters(remarks=remarks))  # list of chapters

    def change_wea(self, wea_id):
        self.wea = self.project.windfarm.get_wea(wea_id)
        remarks = self.wea.report.get_remarks(ordered=False) # ordering is done in remarkSelector
        self.set_chapters(remarks=remarks)
        self.build_structure(remarks=remarks)

    def change_chapter(self, address=None, remarks=None):
        if address is None:
            address = self.chapterstrvar.get().replace(' > ', '|')
        if address == self.active_section: return # don'rebuild if changes through button/combobox
        self.active_section = address
        if remarks is None: remarks=self.wea.report.get_remarks(ordered=False) # ordering is done in remarkSelector
        try:
            self.progbar.update(new_chap=self.active_section)
        except KeyError: # when chapter that is not in every turbine is selected
            pass
        self.project.set_active_chapter(self.active_section)
        self.build_structure(remarks=remarks)

    def set_chapterstrvar(self):
        txt = f'{self.active_section.replace('|', ' > ')}'
        self.chapterstrvar.set(txt)

    def get_chapter_by_diff(self, diff: int):
        '''return address of the chapter that is diff places away from
        current address
        return first/last chapter if diff is too far'''
        curr = self.chapters.index(self.active_section)
        new = curr + diff
        if new < 0:
            new = 0
        elif new > len(self.chapters) - 1:
            new = len(self.chapters) - 1
        return self.chapters[new]
    
    def get_active_section(self):
        return self.active_section
    
    def next_chapter(self, *args):
        chap = self.get_chapter_by_diff(1)
        self.change_chapter(chap)

    def prev_chapter(self, *args):
        chap = self.get_chapter_by_diff(-1)
        self.change_chapter(chap)

    def open_chapter_renamer(self):
        ChapterRenamer(self, self.get_active_section(), self.wea.report,
                        self.parent_mainwindow)
        
    def toggle_rename_btn(self):
        if self.get_active_section().startswith('Prüfbemerkungen'):
            self.chapter_rename_btn.configure(state='normal')
            return
        self.chapter_rename_btn.configure(state='disabled')

    def set_rename_info(self):
        chap = self.get_active_section()
        renames = self.wea.report.get('chapter_renames')
        if not renames: 
            self.renamevar.set('')
            return

        if chap in renames.keys():
            newtitle = renames[chap]
            self.renamevar.set(f'-> "{newtitle}"')
        else: self.renamevar.set('')


class ChapterRenamer(tk.Toplevel):
    def __init__(self, master, curr_address, report, parent_mainwindow, **toplevel_kwargs):
        super().__init__(master, **toplevel_kwargs)
        self.title('Abschnitt umbenennen')
        self.bind('<Control-Return>', lambda event: self.rename()) 

        self.workfrm = ttk.Frame(self)
        self.workfrm.grid_columnconfigure(0, weight=100)

        self.parent_mainwindow = parent_mainwindow
        self.curr_address = curr_address
        self.report = report
        self.project = self.report.parent_project
        self.allreports_var = tk.IntVar(self, value=0)
        self.newname_var = tk.StringVar(self)

        self.build_structure()

        self.workfrm.pack(fill='both', expand=True)

    def build_structure(self):
        chap = self.curr_address.split(db_split_char)[-1]
        titlelbl = ttk.Label(self.workfrm, text=f'"{chap}" umbenennen in:')
        rename_entry = ttk.Entry(self.workfrm, textvariable=self.newname_var)

        all_reports_cbx = ttk.Checkbutton(self.workfrm, text='Bei allen Berichten anwenden',
                                          variable=self.allreports_var)
        ok_btn = ttk.Button(self.workfrm, text='OK', command=self.rename)

        titlelbl.grid(row=0, column=0, columnspan=2, padx=1, pady=2)
        rename_entry.grid(row=1, column=0, columnspan=2, padx=1, pady=1, sticky='ew')
        all_reports_cbx.grid(row=2, column=0, padx=1, pady=1, sticky='w')
        ok_btn.grid(row=2, column=1, padx=1, pady=1, sticky='e')

    
    def rename(self):
        newtitle = self.newname_var.get()
        if not newtitle:
            self.destroy()
            return 'break'
        if db_split_char in newtitle:
            ErrorWindow(self, f'"{db_split_char}" darf nicht im Kapitelnamen enthalten sein.')
            return 'break'

        if self.allreports_var.get():
            self.project.rename_chapters_in_all_reports(self.curr_address, newtitle)
        else: self.report.rename_chapter(self.curr_address, newtitle)
        self.parent_mainwindow.update_weapage(self.report.parent_wea.id)
        self.destroy()








    
class Orderer(tk.Toplevel):
    '''class to order unknown entries into an ordered list'''
    def __init__(self, parent, what: str, to_order: str, address=None,
                 on_close=lambda: None):
        '''what: str, "component", "temperature", "chapter", "title"
        address: str, only relevant if what is "title", gives chapter
        on_close: function that is executed when okay button is invoked'''
        super().__init__(parent)
        self.what = what
        self.to_order = to_order
        self.address = address
        self.on_close = on_close
        self.minsize(500, 500)


        self.set_ordered_list()

        self.workfrm = ttk.Frame(self)
        self.workfrm.pack(fill='both', expand=True)
        self.workfrm.grid_columnconfigure(0, weight=100)
        self.workfrm.grid_rowconfigure(2, weight=100)

        ttk.Label(self.workfrm, text=f'{self.to_order} einsortieren hinter:')\
            .grid(row=0, column=0, sticky='w', padx=1, pady=1)

        self.searchtxtvar = tk.StringVar(self)
        self.searchtxtvar.trace_add('write', self.build_structure)

        ttk.Entry(self.workfrm, textvariable=self.searchtxtvar)\
            .grid(row=1, column=0, sticky='ew', padx=1, pady=1)

        self.scrl = ScrollFrame(self.workfrm)
        self.scrl.grid(row=2, column=0, sticky='nsew', padx=1, pady=1)
        self.selectionfrm = self.scrl.viewPort

        self.optionsvar = tk.StringVar(self, value='Ganz oben')

        ok_btn = ttk.Button(self.workfrm, text='Speichern',
                            command=self.save)
        ok_btn.grid(row=3, column=0, sticky='e', padx=1, pady=1)

        self.build_structure()

        self.bind('<Escape>', self.close)
        self.bind('<Control-Return>', self.save)


    def build_structure(self, *_): 
        gui_f.delete_children(self.selectionfrm)
        self.scrl.gotoTop()
        self.selectionfrm.grid_columnconfigure(0, weight=100)

        search = self.searchtxtvar.get().lower()
        shown_options = [val for val in self.ordered_list \
                                if search.lower() in val.lower()]

        for i, option in enumerate(shown_options):
            rad = MultilineRadiobutton(self.selectionfrm, text=option,
                                       variable=self.optionsvar,
                                       value=option)
            rad.grid(row=i, column=0, sticky='ew')


    def set_ordered_list(self):
        if self.what in ['component', 'temperature', 'chapter']:
            self.ordered_list = dbf.get_order(f'{self.what}s')
            if self.what == 'chapter': self.ordered_list = list(self.ordered_list.keys())
        elif self.what == 'title':
            chapters = dbf.get_order('chapters')
            if not self.address:
                raise KeyError(f'When ordering a title, an address must be given. '
                               f'Got "{self.address}" instead')
            self.ordered_list = chapters[self.address]
        else:
            raise ValueError(f'What is "{self.what}" but must be one of '
                             '["component", "temperature", "chapter", "title"]')
        self.ordered_list.insert(0, 'Ganz oben')
        if self.to_order in self.ordered_list:
            ErrorWindow(f'{self.to_order} is already ordered at index {self.ordered_list.index(self.to_order)}')
            self.destroy()
            
    def set_new_ordered_container(self):
        option = self.optionsvar.get()
        position = self.ordered_list.index(option)
        self.ordered_list = self.ordered_list[1:]
        if self.what == 'chapter' or self.what == 'title':
            chapter_dict = dbf.get_order('chapters')
            if self.what == 'chapter':
                new_item = (self.to_order, [])
                pos = position
            elif self.what == 'title':
                pos = list(chapter_dict.keys()).index(self.address)
                del chapter_dict[self.address]
                self.ordered_list.insert(position, self.to_order)
                new_item = (self.address, self.ordered_list)
            items = list(chapter_dict.items())
            items.insert(pos, new_item)
            self.ordered_list = dict(items)
            return
        
        self.ordered_list.insert(position, self.to_order)

    
    def save(self, *args):
        self.set_new_ordered_container()
        if self.what == 'title': file = f'{mainpath}/databases/order_chapters.txt'
        else: file = f'{mainpath}/databases/order_{self.what}s.txt'

        with open(file, 'w', encoding='utf-8') as f:
            write_str = str(self.ordered_list)\
                                .replace("', ", "',\n")
            if self.what == 'chapter' or self.what == 'title':
                write_str = write_str.replace("',\n", "',\n    ")\
                                     .replace("'],", "',\n    ],\n")\
                                     .replace(": ['", ": [\n    '")\
                                     .replace("], '", "],\n'")\
                                     .replace('[]', '[\n    ]')\
                                     .replace("\n '", "\n'")

            f.write(write_str)
        self.close()

    def close(self, *args):
        self.on_close()
        self.destroy()

class Prefiller(tk.Toplevel):
    '''Window to prefill a textvariable from selectable options'''
    options_per_page = 50
    def __init__(self, master, strvar, options, init_searchtext='', **kwargs):
        '''strvar: tk.StringVar, the variable to prefill
        options: list, the options to choose from'''        
        super().__init__(master, **kwargs)
        self.bind('<Control-Return>', lambda *_: self.set_option())
        self.bind('<Escape>', lambda *_: self.destroy())
        self.strvar = strvar
        self.searchvar = tk.StringVar(self, value=init_searchtext)
        self.searchvar.trace_add('write', lambda *_: self.reset_on_empty())
        self.options = options
        self.selected_option = tk.StringVar(self)

        self.title('Option auswählen')
        self.minsize(800, 500)

        allfrm = ttk.Frame(self)
        allfrm.pack(fill='both', expand=True)

        searchentry = ttk.Entry(allfrm, textvariable=self.searchvar)
        searchentry.bind('<Return>', lambda *_: self.search_options())
        self.scrl = ScrollFrame(allfrm)
        self.option_frame = self.scrl.viewPort

        ttk.Label(allfrm, text='Suche: !!XXX!! -> XXX muss enthalten sein, '
                               'YYY/ZZZ -> YYY oder ZZZ sind enthalten.')\
            .pack(anchor='w', padx=1, pady=1)
        searchentry.pack(fill='x', padx=2, pady=2)
        ttk.Separator(allfrm).pack(fill='x', padx=1, pady=1)
        self.scrl.pack(fill='both', expand=True, padx=2, pady=2)
        self.active_page = 1
        self.active_options = []

        ttk.Separator(allfrm).pack(fill='x', padx=1, pady=1)
        btnsfrm = ttk.Frame(allfrm)
        btnsfrm.pack(fill='x', padx=1, pady=1)
        
        self.nextbtn = ttk.Button(btnsfrm, text='->',
                                  command=self.nextpage, width=3)
        self.prevbtn = ttk.Button(btnsfrm, text='<-',
                                  command=self.prevpage, width=3)
        self.pagestrvar = tk.StringVar(self)
        
        self.ok_button = ttk.Button(btnsfrm, text='OK', command=self.set_option)

        self.prevbtn.pack(side='left', padx=1, pady=1)
        self.nextbtn.pack(side='left', padx=1, pady=1)
        ttk.Label(btnsfrm, textvariable=self.pagestrvar)\
            .pack(side='left', padx=10, pady=1)
        self.ok_button.pack(side='right', pady=2, padx=2)

        self.set_active_options()
        self.set_maxpages()
        self.build_selection_page()

    def build_selection_page(self):
        self.scrl.gotoTop()

        gui_f.delete_children(self.option_frame)
        shown_options = self.active_options[(self.active_page-1)*self.options_per_page:\
                                                self.active_page*self.options_per_page]

        for option in shown_options:
            self.put_option(option)

        self.toggle_changer_btns()
        self.pagestrvar.set(f'{self.active_page}/{self.maxpages}')

    def put_option(self, option):
        radio_btn = MultilineRadiobutton(self.option_frame,
                                         text=gui_f.dbtext2displaytext(option),
                                         variable=self.selected_option,
                                         value=option)
        radio_btn.pack(anchor='w', fill='x', padx=1, pady=1)


    def search_options(self):
        self.set_active_options()
        self.set_maxpages()
        self.active_page = 1
        self.build_selection_page()
    def prevpage(self):
        self.active_page -= 1
        self.build_selection_page()
    def nextpage(self):
        self.active_page += 1
        self.build_selection_page()
    def toggle_changer_btns(self):
        if self.active_page == 1: self.prevbtn.config(state='disabled')
        else: self.prevbtn.config(state='normal')
        if self.active_page == self.maxpages: self.nextbtn.config(state='disabled')
        else: self.nextbtn.config(state='normal')


    def set_active_options(self) -> list:
        active_options = []
        searchtxt = self.searchvar.get().lower()
        if selected_option := self.selected_option.get():
            active_options.append(selected_option)
        if '!!' not in searchtxt and '/' not in searchtxt:
            active_options.extend([opt for opt in self.options if searchtxt in opt.lower()])
            self.active_options = active_options
        elif searchtxt.count('!!')%2 == 1:
            ErrorWindow(self, f'!! muss geradzahlig oft vorkommen.', self.focus)
            self.active_options = active_options
        else:
            mandatories, or_optionals = gui_f.extract_between_substrings(searchtxt, '!!', '/')
            active_options = [v for v in self.options]
            for mandatory in mandatories:
                active_options = [val for val in active_options if mandatory.lower() in val.lower()]
            if or_optionals:
                active_options = [option for option in active_options\
                                    if any(sub.lower() in option.lower() for sub in or_optionals)]
            self.active_options = active_options
        l = len(self.active_options)
        self.title(f'Option auswählen ({l} verfügbar)')


    def set_maxpages(self):
        l = len(self.active_options)
        self.maxpages = int(np.ceil(l/self.options_per_page))

    def reset_on_empty(self):
        if not self.searchvar.get():
            self.search_options()

    def set_option(self):
        selected_option = self.selected_option.get()
        if '+++' in selected_option:
            selected_option = selected_option[selected_option.rfind('+++')+4:].strip()
        if (cutoff:=selected_option.rfind('(Verwendet in:')) > -1:
            selected_option = selected_option[:cutoff].strip()
        self.strvar.set(selected_option)
        self.destroy()

class Chapterbar(ttk.Frame):
    def __init__(self, master, report, changer_func, active_sec=None,
                 remarks=None, **kwargs):
        super().__init__(master, **kwargs)
        self.master = master
        self.report = report
        self.wea = self.report.parent_wea

        self.changer_func = changer_func

        self.all_chapters = ['Extras']
        self.all_chapters.extend(report.get_chapters(remarks=remarks))
        self.done_chapters = report.get_done_chapters()
        self.active_sec = active_sec

        self.bar_elements = {}
        self.rem_indicators = {}
        self.build_structure(remarks=remarks)

    def build_structure(self, remarks=None):
        if remarks is None: remarks = self.report.get_remarks(ordered=False)
        self.grid_rowconfigure(0, weight=100)
        old_chap = ''
        old_subchap = ''
        col = 0
        for chapter in self.all_chapters:
            new_chap = chapter.split(db_split_char)[0]
            try: new_subchap = chapter.split(db_split_char)[1]
            except IndexError: new_subchap = ''

            if (new_chap != old_chap) and \
                    (new_chap=='Prüfbemerkungen' or old_chap=='Prüfbemerkungen'):
                self.place_chap_separator(col)
                self.grid_columnconfigure(col, weight=0)
                col += 1
            elif (new_subchap != old_subchap) and \
                    (new_chap=='Prüfbemerkungen' or old_chap=='Prüfbemerkungen'):
                self.place_subchap_separator(col)
                self.grid_columnconfigure(col, weight=0)
                col += 1
            old_chap = new_chap
            old_subchap = new_subchap

            color = self.get_chapter_color(chapter)
            frm = tk.Frame(self, background=color)
            self.grid_columnconfigure(col, weight=1)
            frm.grid(row=0, column=col, sticky='nsew', padx=0, pady=0)
            frm.bind('<Button-1>', lambda e, chap=chapter: self.changer_func(chap))
            self.bar_elements[chapter] = frm

            chapterrems = dbf.get_remarks_of_chapter(remarks, chapter)
            if chapterrems.empty:
                col += 1
                continue

            flags = chapterrems.flag.to_list()
            self.update_remindicator(chapter, flags)
            col += 1

    def place_chap_separator(self, col):
        tk.Frame(self, background='gray54', width=1)\
            .grid(row=0, column=col, sticky='ns', padx=0, pady=0)

    def place_subchap_separator(self, col):
        tk.Frame(self, background='gray54', width=1)\
            .grid(row=0, column=col, sticky='ns', padx=0, pady=4)

    def get_chapter_color(self, chapter):
        if chapter == 'Extras':
            if chapter == self.active_sec: return 'purple'
            else: return 'deep sky blue'
        if chapter in self.done_chapters:
            if chapter == self.active_sec: return 'green2'
            else: return 'pale green'
        if chapter == self.active_sec: return 'orchid2'
        else: return 'white smoke'

    def update(self, new_chap):
        old_chap = self.active_sec
        old_frm = self.bar_elements[self.active_sec]
        self.active_sec = new_chap
        new_frm = self.bar_elements[self.active_sec]
        old_frm.config(background=self.get_chapter_color(old_chap))
        new_frm.config(background=self.get_chapter_color(self.active_sec))

    def update_remindicator(self, chapter, flags):
        col = dbf.get_flagcolor(flags, only_PIEV=True)
        column = self.bar_elements[chapter].grid_info()['column']

        try:
            frm = self.rem_indicators[chapter]
        except KeyError:
            frm = tk.Frame(self, background=col, height=5)
        
        frm.grid_forget()
        if col is None:
            return
        frm.configure(background=col)
        frm.grid(row=1, column=column, sticky='nsew', pady=0, padx=0)
        frm.bind('<Button-1>', lambda _, chap=chapter: self.changer_func(chap))
        self.rem_indicators[chapter] = frm




class RemarkArt(ttk.Frame):
    def __init__(self, master, project, park_remarks_slice,
                 changer_func, **kwargs):
        '''Small Image showing the presence ob the current remark in the
        project's wea.
        Green / Orange / Red / Light Blue --> V / E / P(PP) / I
        White / Grey --> Remark not in WEA's report / WEA not setup
        Purple --> other flag
        '''
        super().__init__(master, **kwargs)
        self.project = project
        self.rems = park_remarks_slice
        self.wea_ids = project.windfarm.get_wea_ids()
        self.changer_func = changer_func

        self.build_structure()

    def build_structure(self):
        self.elements = {}
        colordict = self.build_colordict()
        for i, wea_id in enumerate(colordict.keys()):
            col = colordict[wea_id]
            frm = tk.Frame(self, background=col, width=8, height=8)
            frm.grid(column=i%3, row=int(i/3), padx=0, pady=0)
            frm.bind('<Button-1>', lambda e, id=wea_id: self.changer_func(id))
            self.elements[wea_id] = frm

    def build_colordict(self) -> dict:
        colors = {}
        for wea_id in self.wea_ids:
            wea = self.project.windfarm.get_wea(wea_id)
            if not wea.is_setup:
                colors[wea_id] = 'gray54'
                continue
            wea_rems = self.rems[wea_id]
            if not wea_rems:
                colors[wea_id] = 'snow'
                continue
            flags = [rem[0] for rem in wea_rems]
            colors[wea_id] = dbf.get_flagcolor(flags)
        return colors



class GitButton(ttk.Button):
    def __init__(self, master, project, *args, mode='pull', **kwargs):
        '''mode: pull to update local database, push to push local data to main'''
        self.mode = mode
        self.project = project
        if self.mode == 'pull':
            command = self.pull_from_main
            btntext = 'Daten holen'
        elif self.mode == 'push':
            command = self.push_to_main
            btntext = 'Daten hochladen'
        super().__init__(master, *args, text=btntext, command=command, **kwargs)

    def git_command(self, command, cwd="."):
        """
        Run a Git command and return its output.
        """

        try:
            result = subprocess.run(
                f'cd {mainpath}/databases && {command}',
                cwd=cwd,
                shell=True,
                check=True,
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            return result.stdout
        except subprocess.CalledProcessError as e:
            return f"Error: {e.stderr}"
        
    def push(self):
        push_result = self.git_command('git push')
        if 'Error' in push_result:
            ErrorWindow(self, f'Error while pushing: {push_result}')
            return 'err'
        return 'succ'

    def commit_databases(self, commit_message):
        self.git_command('git add .')
        status = self.git_command('git status')
        if 'no changes added to commit' in status or\
            'nothing to commit' in status:
            return 'succ'
        commit_result = self.git_command(
            f'git commit -m "{commit_message}"'
        )
        if "Error" in commit_result:
            ErrorWindow(self, f"Failed to commit: {commit_result}")
            return 'err'
        return 'succ'
        

    def pull_from_main(self):     
        # commit changes before pulling
        ret = self.commit_databases(f'{self.project.get('name')} before pulling at '
                f'{pd.to_datetime(datetime.datetime.now())}')
        if ret == 'err': return

        # Pull from main
        pull_result = self.git_command(f"git pull origin main")
        if "Error" in pull_result:
            ErrorWindow(self, f"pull from main fehlgeschlagen: {pull_result}")
            return
        elif 'merge conflict' in pull_result.lower():
            ErrorWindow(self, 'Merge conflict, bitte manuell mergen.')
            return
        else:
            ErrorWindow(self, "Daten erfolgreich heruntergeladen.")


    def push_to_main(self):
        # Add and commit changes
        ret = self.commit_databases(f'{self.project.get('name')} before pushing at '
                    f'{pd.to_datetime(datetime.datetime.now())}')
        if ret == 'err': return
        ret = self.push()
        if ret == 'err': return

        ErrorWindow(self, "Daten erfolgreich hochgeladen")



### ------------------ HELPER CLASSES ---------------------###
class ScrollFrame(tk.Frame):
    def __init__(self, parent, orient='vertical', use_mousewheel=True,
                 def_height=None, def_width=None):
        super().__init__(parent) # create a frame (self)
        
        self.orient = orient
        self.canvas = tk.Canvas(self, borderwidth=0, 
                                height=def_height, width=def_width)                     #place canvas on self
        self.viewPort = ttk.Frame(self.canvas)                                           #place a frame on the canvas, this frame will hold the child widgets 

        if self.orient == 'vertical':
            self.vsb = tk.Scrollbar(self, orient=orient, command=self.canvas.yview)     #place a scrollbar on self 
            self.canvas.configure(yscrollcommand=self.vsb.set)                          #attach scrollbar action to scroll of canvas

            self.vsb.pack(side="right", fill="y")                                       #pack scrollbar to right of self
            self.canvas.pack(side="right", fill="both", expand=True)                     #pack canvas to left of self and expand to fil

        elif orient == 'horizontal':
            self.hsb = tk.Scrollbar(self, orient=orient, command=self.canvas.xview) #place a scrollbar on self 
            self.canvas.configure(xscrollcommand=self.hsb.set)                          #attach scrollbar action to scroll of canvas

            self.hsb.pack(side="bottom", fill="x")                                       #pack scrollbar to right of self
            self.canvas.pack(side="left", fill="both", expand=True)                     #pack canvas to left of self and expand to fil


        self.canvas_window = self.canvas.create_window((4,4),
                                            window=self.viewPort, anchor="nw",            #add view port frame to canvas
                                            tags="self.viewPort")

        self.viewPort.bind("<Configure>", self.onFrameConfigure)                       #bind an event whenever the size of the viewPort frame changes.
        self.canvas.bind("<Configure>", self.onCanvasConfigure)                       #bind an event whenever the size of the canvas frame changes.
        
        if use_mousewheel:
            self.viewPort.bind('<Enter>', self.onEnter)                                 # bind wheel events when the cursor enters the control
            self.viewPort.bind('<Leave>', self.onLeave)                                 # unbind wheel events when the cursorl leaves the control

        self.onFrameConfigure(None)                                                 #perform an initial stretch on render, otherwise the scroll region has a tiny border until the first resize

    def onFrameConfigure(self, event):                                              
        '''Reset the scroll region to encompass the inner frame'''
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))                 #whenever the size of the frame changes, alter the scroll region respectively.

    def onCanvasConfigure(self, event):
        '''Reset the canvas window to encompass inner frame when required'''
        canvas_width = event.width
        canvas_height = event.height
        kw = {'width': canvas_width} if self.orient == 'vertical' else {'height': canvas_height}
        self.canvas.itemconfig(self.canvas_window, **kw)            #whenever the size of the canvas changes alter the window region respectively.

    def onMouseWheel(self, event):                                                  # cross platform scroll wheel event
        def func(self, *args): 
            if self.orient == 'vertical' and \
                    self.canvas.winfo_height() < self.viewPort.winfo_height():
                self.canvas.yview_scroll(*args)
            elif self.orient == 'horizontal' and \
                    self.canvas.winfo_width() < self.viewPort.winfo_width():
                self.canvas.xview_scroll(*args)
        fac = -1

        if platform.system() == 'Windows':
            func(self, int(fac*(event.delta/120)), "units")
        elif platform.system() == 'Darwin':
            func(self, int(-1 * event.delta), "units")
        else:
            if event.num == 4:
                func(self,  -1, "units" )
            elif event.num == 5:
                func(self, 1, "units" )

    def onEnter(self, event):                                                       # bind wheel events when the cursor enters the control
        if platform.system() == 'Linux':
            self.canvas.bind_all("<Button-4>", self.onMouseWheel)
            self.canvas.bind_all("<Button-5>", self.onMouseWheel)
        else:
            self.canvas.bind_all("<MouseWheel>", self.onMouseWheel)

    def onLeave(self, event):                                                       # unbind wheel events when the cursorl leaves the control
        if platform.system() == 'Linux':
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")
        else:
            self.canvas.unbind_all("<MouseWheel>")

    def gotoTop(self, *args):
        self.canvas.yview_moveto(0)

class Quitbutton(ttk.Frame):
    def __init__(self, parent, parent_window, text='Schließen', **pack_kwargs):
        super().__init__(parent)
        self.parent_window = parent_window
        qbtn = ttk.Button(self, text=text, command=self.kaputt)
        qbtn.pack(**pack_kwargs)

    def kaputt(self):
        self.parent_window.destroy()
        sys.exit()

class Label_w_Button:
    def __init__(self, labelmaster, labeltext: str,
                 command, command_args=[], command_kwargs={},
                 call_callables=True,
                 buttonmaster=None,
                 buttonkwargs={}, **label_kwargs):
        self.labelmaster = labelmaster
        self.buttonmaster = buttonmaster if buttonmaster is not None\
              else self.labelmaster
        self.command = command
        self.command_args = command_args
        self.labeltext = labeltext
        self.txtvar = tk.StringVar(self.labelmaster)
        self.txtvar.set(self.labeltext)
        self.call_callables = call_callables
        self.command_kwargs = command_kwargs
        self.label = MultilineLabel(self.labelmaster, textvar=self.txtvar,
                                    **label_kwargs)
        self.btn = ttk.Button(self.buttonmaster,
                              command=self.button_command,
                              **buttonkwargs)
        
    def button_command(self):
        if not self.call_callables:
            self.command(*self.command_args, **self.command_kwargs)
            return
        
        command_args, command_kwargs = self.execute_callables()
        self.command(*command_args, **command_kwargs)

    def execute_callables(self):
        command_kwargs = {}
        command_args = []
        for k in self.command_kwargs.keys():
            kwarg = self.command_kwargs[k]
            if callable(kwarg):
                command_kwargs[k] = kwarg()
                continue
            command_kwargs[k] = kwarg
        for arg in self.command_args:
            if callable(arg):
                command_args.append(arg())
                continue
            command_args.append(arg)
        return command_args, command_kwargs

class MultilineRadiobutton(ttk.Frame):
    def __init__(self, master, label_kwargs={}, **kwargs):
        if 'text' in kwargs.keys():
            text = kwargs.pop('text')
        super().__init__(master)
        self.radiobutton = ttk.Radiobutton(self, **kwargs)
        self.lblstrvar = tk.StringVar(self, value=text)
        self.label = MultilineLabel(self, self.lblstrvar, **label_kwargs)

        self.radiobutton.pack(side='left', padx=1, pady=1)
        self.label.pack(side='left', fill='x', expand=True, padx=1, pady=1)
        self.label.bind('<Button-1>', self.on_click)
        self.bind('<Button-1>', self.on_click)
        
    def on_click(self, *_):
        self.radiobutton.invoke()

class Conditional_Combobox(ttk.Combobox):
    def __init__(self, master, parent_strvar: tk.StringVar, value_dict: dict,
                 autoupdate_value=False,
                 **kwargs):
        '''
        Combobox whose selectable values change according to a dict and the
        value of the parent_strvar

        parent_strvar: tk.StringVar()
        value_dict: dict whose keys are possible values of parent_strvar
            and whose values are lists of corresponding selectables
        '''
        self.master = master
        self.autoupdate_value = autoupdate_value
        default_value = ''
        if 'textvariable' not in kwargs:
            self.txtvar = tk.StringVar(self, )
        else:
            self.txtvar = kwargs['textvariable']
            # getting default value of thhe given strvar
            default_value = self.txtvar.get()
            del kwargs['textvariable']
        if not isinstance(parent_strvar, tk.StringVar):
            raise TypeError(f'parent_strvar must be a tk.StringVar, but is {type(parent_strvar)}')
        self.parent_strvar = parent_strvar

        self.value_dict = value_dict
        super().__init__(self.master, textvariable=self.txtvar, **kwargs)

        self.parent_strvar.trace_add('write', self.update_values)
        self.update_values()
        # reset default value after updating the possible values
        self.txtvar.set(default_value)

    def update_values(self, *_):
        parent_val = self.parent_strvar.get()
        self.values = self.value_dict.get(parent_val, [''])
        if type(self.values) == str: self.values = [self.values]
        else:
            try: self.values = list(self.values)
            except TypeError: self.values = [self.values]
        self['values'] = self.values
        if self.autoupdate_value:
            val = self.values[0]
            if not (pd.isna(val) or str(val).lower() == 'nan'): self.set(val)
        self.unbind_class('TCombobox', '<MouseWheel>')
        self.unbind_class('TCombobox', '<ButtonPress-4>')
        self.unbind_class('TCombobox', '<ButtonPress-5>')

    def replace_value_dict(self, value_dict, **kwargs):
        self.value_dict = value_dict
        self.update_values(**kwargs)


class Entryline():
    '''
    creates a frame that contains a label and a box for data entering
    textbox can be prefilled, or have a button to select from prefill options
    '''
    def __init__(self, parent, label: str,
                 prefill: Optional[str|list]=None,
                 immutable: Optional[bool]=False,
                 multiline: Optional[bool]=False,
                 prefill_options: Optional[list|dict]=[],
                 **kwargs):
        '''initialize label and box for data entering
        label: str, text of the label
        prefill: str or list, text that is prefilled in the box
        immutable: bool, if True, the box is readonly
        multiline: bool, if True, the box is a multiline textbox
        prefill_options: list or dict, options for prefilling the box
            if list: entry is a combobox
            if dict: entry is a conditional combobox
            does not work with multiline=True
        kwargs are passed to the data entering box'''

        self.txtvar = tk.StringVar(parent)
        if prefill is not None:
            prefill = str(prefill).strip()
            if prefill.endswith('.0'): prefill = prefill[:-2]
            self.txtvar.set(prefill)
        self.label=label
        par =  parent
        self.btn = None

        self.tklabel = ttk.Label(par, text=self.label)

        if multiline: # case: make Text Widget
            self.entry = StrVarText(par, textvariable=self.txtvar,
                                    **kwargs)
            if not isinstance(prefill_options, list)\
                    and not callable(prefill_options)\
                    and prefill_options is not None:
                raise TypeError(f'prefill option must be None, function or list in combination with multiline, but is {type(prefill_options)}')
            if prefill_options:
                command = prefill_options if callable(prefill_options) else\
                    lambda *_: Prefiller(self, self.txtvar, prefill_options)
                self.btn = ttk.Button(par, text='+', width=3, command=command)
        elif callable(prefill_options) or not prefill_options: # case: make entry widget
            kwargs.pop('height', None)
            self.entry = ttk.Entry(par, textvariable=self.txtvar, **kwargs)
            if callable(prefill_options): self.btn = ttk.Button(
                                            par, text='+',
                                            width=3,
                                            command=lambda: prefill_options())
        else: # make combo/conditional combo
            kwargs.pop('height', None)
            if isinstance(prefill_options, list):
                self.entry = ttk.Combobox(par, textvariable=self.txtvar,
                                          values=prefill_options, **kwargs)
            elif isinstance(prefill_options, dict):
                try: parent_strvar = prefill_options.pop('depend_on')
                except KeyError: raise KeyError('prefill_options dict must have a key "depend_on". The key\'s value must be the strvar according to which the dict_values change.')
                self.entry = Conditional_Combobox(par, parent_strvar,
                                                  prefill_options,
                                                  textvariable=self.txtvar,
                                                  **kwargs)

        if immutable: self.entry.config(state='disabled')
    
    def get_textvar(self):
        return self.txtvar    
    def get_label(self):
        return self.tklabel    
    def get_labeltext(self):
        return self.label    
    def get_entry(self):
        return self.entry    
    def get_button(self):
        return self.btn
    
    def has_text(self):
        if self.txtvar.get(): return True
        return False

class EntryFrame(ttk.Frame):
    '''Frame that contains several labels with textboxes next to them
    can output a dict: {'property1': 'entry1', 'property2': 'entry2', ...}
    '''
    def __init__(self, parent, properties_clearname: list, 
                 prefilled_object=None,
                 clearname2argname: Optional[dict]=None,
                 autoplace_start_row=None,
                 immutable: list=[],
                 multiline: list=[],
                 prefill_options: dict={},
                 **frame_kwargs
                 ):
        '''initialize the frame with labels and boxes for data entering.
        properties_clearname: list of clearnames of the properties
        prefilled_object: obj/dict, object that contains the prefilled data
            if dict: puts value into the entry where key==property (clearname)
        clearname2argname: dict that maps clearnames to argnames
        autoplace_start_row: grids label, entry, prefill button start from this line
        immutable: list of clearnames that should be readonly
        multiline: list of clearnames that should have a multiline entry
        prefill_options: dict that maps clearnames to either
            lists of options (place combobox) or
            dict of options (place conditional combobox)
        '''
        self.clearname2argname = clearname2argname
        self.properties_clearname = properties_clearname
        self.properties_with_entries = {}
        self.entrylines = {}
        self.label_with_entries = []
        for property in self.properties_clearname:
            kwargs = {**frame_kwargs}
            if property in immutable: kwargs['immutable'] = True
            else: kwargs['immutable'] = False
            if property in multiline:
                kwargs['multiline'] = True
                if 'height' not in kwargs.keys(): kwargs['height'] = 3
                if 'wrap' not in kwargs.keys(): kwargs['wrap'] = 'word'
            else:
                kwargs.pop('multiline', None)
                kwargs.pop('wrap', None)
                kwargs.pop('height', None)

            prefills = prefill_options.get(property, None)
            if isinstance(prefills, dict):
                # --> dict means conditional combobox
                try: 
                    prefills['depend_on'] = self.get_textvars()[prefills['depend_on']]
                except KeyError:
                    raise KeyError(f'{prefills['depend_on']} not in textvariables. Make sure to put the the property being depended on before the dependent property in properties_clearnames.')
            # prefill: what is the text of the line
            prefill = gui_f.get_attribute_from_clearname(prefilled_object,
                                                         property,
                                                         clearname2argname) # returns None if any arg is None
            property_entryer = Entryline(parent,
                                         label=property,
                                         prefill=prefill,
                                         prefill_options=prefills,
                                         **kwargs)
            self.properties_with_entries[property_entryer.get_labeltext()] = \
                                                  property_entryer.get_textvar()
            self.entrylines[property] = property_entryer
            self.label_with_entries.append(property_entryer)
        if not autoplace_start_row: return
        for self.row, (label, entry, btn) in enumerate(
                            self.get_labels_entries_buttons(),
                            autoplace_start_row):
            stick = 'n' if isinstance(entry, StrVarText) else ''
            label.grid(row=self.row, column=0, pady=1, padx=1, sticky=f'{stick}w')
            entry.grid(row=self.row, column=1, pady=1, padx=5, sticky='ew')
            if btn is not None:
                btn.grid(row=self.row, column=2, padx=1, pady=1, sticky=f'{stick}e')

    def get_clearname_entryvalue_dict(self):
        '''puts the non-readable txtvar of the entries into readable
        str-containing dicts and prints them out
        '''
        clearname_entryvalue_dict = {}
        for clearname in self.properties_with_entries.keys():
            clearname_entryvalue_dict[clearname] = self.properties_with_entries[clearname].get()
        return clearname_entryvalue_dict
        
    def get_argname_entry_dict(self):
        argname_entry_dict = {}
        for clearname in self.properties_with_entries.keys():
            argname = self.clearname2argname[clearname]
            argname_entry_dict[argname] = self.properties_with_entries[clearname].get()
        return argname_entry_dict
    
    def is_completely_filled_out(self):
        for entry in self.label_with_entries:
            if not entry.has_text():
                return False
        return True

    def get_textvars(self) -> dict:
        '''return a dict like {clearname: textvar, ...}'''
        return self.properties_with_entries
    
    def get_labels_entries_buttons(self) -> list:
        '''returns a list of labels (ttk.Label) with corresponding entries
        (ttk.Entry, ttk.Combobox, Conditional_Combobox, StrVarText) and Buttons
        (ttk.Button/None) like
        [(label1, entry1, button1), (label2, entry2, button2), ...]
        '''
        lwe = []
        for label_with_entry in self.label_with_entries:
            lwe.append((label_with_entry.tklabel,
                        label_with_entry.entry,
                        label_with_entry.btn))
        return lwe
    def get_entrylines(self) -> dict:
        '''return a dict like {clearname: entryline, ...}'''
        return self.entrylines
    
    def get_next_free_row(self):
        return self.row + 1


class Checkbox_with_Label(ttk.Frame):
    '''single checkbox with a label, can output state and labeltext'''
    def __init__(self, parent, text, defaultvalue=True, command=None, **kwargs):
        super().__init__(parent, **kwargs)

        self.istrue = tk.IntVar(self, value=int(defaultvalue))
        self.text = text

        self.checkbut = ttk.Checkbutton(self, text=self.text, variable=self.istrue, command=command, **kwargs)
        self.checkbut.grid(sticky='nswe', padx=2, pady=2)

    def get_state(self):
        return self.istrue.get()
    
    def get_text(self):
        return self.text
        
class Checkboxes_Group(dict):
    """creates multiple checkboxes with some logic. access boxes through
    self['option'] -> (Checkbutton, IntVar)
    Key of select all box: all"""
    def __init__(self, master, options, select_all_box: bool=False, default_state=0, **kwargs):
        """creates multiple checkboxes with some logic. access boxes through
        self.boxes['option'] -> (Checkbutton, IntVar)
        Key of select all box: all"""
        super().__init__()
        if default_state not in [0, 1]:
            raise ValueError(f'default_state must be 0 or 1, but is {default_state}')
        if select_all_box:
            var = tk.IntVar(master, value=default_state)
            box = ttk.Checkbutton(master, text='Alle auswählen',
                                  command=self.toggle_all, variable=var)
            self['all'] = (box, var)

        for option in options:
            var = tk.IntVar(master, value=default_state)
            box = ttk.Checkbutton(master, text=option, variable=var,
                                  command=self.on_toggle)
            self[option] = (box, var)

    def get_selected(self) -> list:
        '''return a list of all selected options of the group'''
        selected = []
        for option in self.keys():
            if option == 'all':
                continue
            _, var = self[option]
            if int(var.get()):
                selected.append(option)
        return selected
    
    def get_states(self) -> dict:
        '''return a dict like {option: 0 / 1, ...}'''
        states = {}
        for option in self.keys():
            if option == 'all':
                continue
            _, var = self[option]
            states[option] = var.get()
        return states
    
    def activate_all(self):
        for option in self.keys():
            var = self[option][1]
            var.set(1)

    def deactivate_all(self):
        for option in self.keys():
            var = self[option][1]
            var.set(0)

    def toggle_all(self):
        if self['all'][1].get(): self.activate_all()
        else: self.deactivate_all()

    def on_toggle(self):
        if 'all' not in self.keys():
            return
        var = self['all'][1]
        selected = self.get_selected()
        all_length = len(self.keys()) - 1 # -1 because 'all is also a key'
        if len(selected) == all_length: var.set(1)
        else: var.set(0)

class MultilineLabel(ttk.Label):
    def __init__(self, master, textvar, **label_kwargs):
        self.text_var = textvar
        super().__init__(master, textvariable=self.text_var, **label_kwargs)
        self.bind('<Configure>', self._update_wraplength)

    def _update_wraplength(self, event):
        self.config(wraplength=event.width-1)

class StrVarText(tk.Text):
    def __init__(self, master, textvariable=None, **kwargs):
        '''tk.Text having a StringVar, like ttk.Entry. Replaces newlines with |.'''
        self.textvariable = textvariable
        if self.textvariable is None:
            self.textvariable = tk.StringVar(master)
        self.is_updating=False
        super().__init__(master, **kwargs)
        self.configure(font=font.nametofont('TkDefaultFont'))

        self.textvariable.trace_add('write', self.update_text)
        self.bind('<<Modified>>', self.update_strvar)
        self.update_text()

    def update_text(self, *args):
        if self.is_updating:
            return
        self.is_updating = True
        strvar_text = self.textvariable.get().replace(db_split_char, '\n')
        self.delete('1.0', 'end')
        self.insert('1.0', strvar_text)
        self.is_updating = False

    def update_strvar(self, *_):
        if self.is_updating:
            return
        if self.edit_modified():
            self.is_updating = True
            new_value = self.get('1.0', 'end-1c').replace('\n', db_split_char)
            self.textvariable.set(new_value)
            self.edit_modified(False)
            self.is_updating = False






