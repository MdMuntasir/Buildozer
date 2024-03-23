import os
import requests
from fpdf import FPDF
from kivy.utils import platform
from kivy.core.window import Window
from kivy.lang import Builder
from kivy.metrics import dp
from kivymd.uix.datatables import MDDataTable
from kivymd.uix.dialog import MDDialog
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.gridlayout import MDGridLayout
from kivymd.uix.screenmanager import MDScreenManager
from kivymd.uix.screen import MDScreen
from openpyxl import load_workbook as wb
from openpyxl import Workbook as wrk
from openpyxl.styles import Font, Alignment, Border, Side
from kivymd.app import MDApp
from kivy.uix.popup import Popup
from kivymd.uix.boxlayout import BoxLayout, MDBoxLayout
from kivymd.uix.button import MDRectangleFlatButton
from kivymd.uix.textfield import MDTextField
from kivymd.uix.label import MDLabel
from kivymd.uix.scrollview import MDScrollView

from Functions.functions import functions

if platform == "android":
    from androidstorage4kivy import SharedStorage
from kivymd.uix.behaviors import BackgroundColorBehavior, CommonElevationBehavior

hm = 'Follow these steps\n\n1. Click on "Download Routine" and wait a few seconds to download the latest routine. Click "Current Routine" to use the already downloaded routine.' \
     'You need to download the routine at least one time after installation to use "Current Routine"\n\n2.1 ' \
     'For individual section routine click "Student". Now enter your batch and section in the batch and section box \n\n2.2 ' \
     'To get routine for a specific teacher routine click "Teacher". Now enter teacher initial of that teacher \n\n3.1' \
     'Click on "Next" button to see and save routine. \n\n3.2' \
     'Click on "Empty Slots" button to see and save empty slots from the routine. \n\n4. ' \
     'Click "Show Routine" to see routine and "Save" to save your routine in pdf format.'



semester = {}

emt_days = {
    "Sat": {},
    "Sun": {},
    "Mon": {},
    "Tue": {},
    "Wed": {},
    "Thu": {}
}

days = {
            "Sat": [],
            "Sun": [],
            "Mon": [],
            "Tue": [],
            "Wed": [],
            "Thu": []
        }

day = {}
times = []
loctn = ['', '', '']
color = {}


# def extractor(loc):
#
#     clr_sheet = wb(loc)
#     info_sheet = clr_sheet.worksheets[1]
#     c_count = 66
#     cd_count = 3
#
#     while info_sheet[chr(c_count)+'1'].value:
#         color[str(info_sheet[chr(c_count)+'1'].value)] = info_sheet[chr(c_count)+'1'].fill.fgColor.rgb
#         c_count += 1
#
#     while info_sheet["A"+str(cd_count)].value:
#         semester[str(info_sheet["A"+str(cd_count)].value)] = info_sheet["B"+str(cd_count)].value
#         cd_count += 1
#     clr_sheet.close()





Generated = False

kv = '''

ScreenManager:

    HomePage:
    StudentPage:
    Student_Routine_Show:
    TeacherPage:
    Teacher_Routine_Show:
    Blank_Slot_Routine:
    Help:



<Help>:

    name: "help"

    BoxLayout:
        padding: 20
        orientation: "vertical"
        MDLabel:
            text : "Instructions"
            style : "Times New Roman"
            font_style: "H3"
            size_hint: 1,.1
            pos_hint: {"x":.15,"center_y": 1}

        MDLabel:
            size_hint: .5,.05
            pos_hint: {"center_x":.5,"center_y":1}  

        Scroll:
            size_hint: 1,.75
            do_scroll_y : True
            MDLabel:
                text: root.message
                text_size :  self.width, None
                pos_hint: {"center_x":1,"center_y":1}
                valign:"top"
                size_hint_y: None
                height: max(self.texture_size[1],self.parent.height)


        MDLabel:
            size_hint: .5,.05
            pos_hint: {"center_x":.5,"center_y":1}            

        MDRectangleFlatButton:
            text: "Close"
            pos_hint: {"center_x":.5,"center_y":1}
            on_press: root.manager.current = "Home"

        MDLabel:
            size_hint: .5,.1
            pos_hint: {"center_x":.5,"center_y":1}


<StudentPage>:
    name: "StudentPage"

<Student_Routine_Show>:
    name: "student_show"

<HomePage>:
    name: "Home"

<TeacherPage>:
    name: "TeacherPage"

<Teacher_Routine_Show>:
    name: "teacher_show"

<Blank_Slot_Routine>:
    name: "empty_slot"
<Scroll@MDScrollView>

'''


class HomePage(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        download_button = MDRectangleFlatButton(text="Download Routine", line_color="#30c6d1", size_hint=[.6, .04],
                                                pos_hint={"center_x": .5, "center_y": .85},
                                                on_press=self.download, md_bg_color="#30c6d1", text_color="#292929")
        self.current_button = MDRectangleFlatButton(text="Current Routine", line_color="#c97b22", size_hint=[.6, .04],
                                                    pos_hint={"center_x": .5, "center_y": .75},
                                                    on_press=self.current, md_bg_color="#c97b22", text_color="#121212")

        self.b1 = MDRectangleFlatButton(text="Student", line_color="#005252", size_hint=[.4, .04],
                                        pos_hint={"center_x": .5, "center_y": .6}, md_bg_color="#005252",
                                        text_color="#ffffff", on_press=self.student_page, disabled=True)
        self.b2 = MDRectangleFlatButton(text="Teacher", line_color="#005252", size_hint=[.4, .04],
                                        pos_hint={"center_x": .5, "center_y": .5}, md_bg_color="#005252",
                                        text_color="#ffffff", on_press=self.teacher_page, disabled=True)
        self.b3 = MDRectangleFlatButton(text="Empty Slots", line_color="#005252", size_hint=[.4, .04],
                                        pos_hint={"center_x": .5, "center_y": .4}, md_bg_color="#005252",
                                        text_color="#ffffff", on_press=self.emptySlote_page, disabled=True)
        self.help_button = MDRectangleFlatButton(text="Help", size_hint=[.2, .06],
                                                 pos_hint={"center_x": .5, "center_y": .15},
                                                 line_color="#3c3d3d", md_bg_color="#3c3d3d", on_press=self.helpPage,
                                                 text_color="white")

        self.add_widget(download_button)
        self.add_widget(self.current_button)
        self.add_widget(self.b1)
        self.add_widget(self.b2)
        self.add_widget(self.b3)
        self.add_widget(self.help_button)

        self.dialouge = MDDialog(text="Downloading...")

    def student_page(self, arg):
        self.manager.current = "StudentPage"

    def teacher_page(self, arg):
        self.manager.current = "TeacherPage"

    def emptySlote_page(self, arg):
        functions.times,functions.emt_days = functions.empty_slot(loctn[2])
        self.manager.current = "empty_slot"

    def helpPage(self, arg):
        self.manager.current = "help"

    def download(self, obj):

        if platform == 'android':
            from android.storage import app_storage_path
            from android import mActivity

            context = mActivity.getApplicationContext()
            result = context.getExternalFilesDir(None)
            if result:
                storage_path = str(result.toString())
            else:
                storage_path = app_storage_path()
        else:
            storage_path = os.curdir

        if not os.path.exists(os.path.join(storage_path, "App Data")):
            os.makedirs(os.path.join(storage_path, "App Data"))
        storage_path = os.path.join(storage_path, "App Data")

        temp_path = os.path.join(storage_path, "routine.xlsx")

        rtn = 'https://drive.google.com/uc?id=1e5_vL6oA4OtPBYb8Nyv1kizkdriBZ8eW'
        rtn_response = requests.get(rtn)
        if rtn_response.status_code == 200:
            with open(temp_path, 'wb') as file:
                file.write(rtn_response.content)

            down = BoxLayout()
            down.txt = MDLabel(text="Downloaded",
                               size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                               theme_text_color="Custom", text_color="#e6e5e3")
            down.add_widget(down.txt)
            pop = Popup(title=storage_path, content=down, size_hint=[.8, .2], pos_hint={"center_x": .5})
            # extractor(temp_path)

            pop.open()
            loctn[2] = temp_path
            self.b1.disabled = False
            self.b2.disabled = False
            self.b3.disabled = False
        else:
            down = BoxLayout()
            down.txt = MDLabel(text="Try Again",
                               size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                               theme_text_color="Custom", text_color="#e6e5e3")
            down.add_widget(down.txt)
            pop = Popup(title="Download Failed", content=down, size_hint=[.8, .2], pos_hint={"center_x": .5})
            # extractor(temp_path)
            pop.open()

    def current(self, obj):
        if platform == 'android':
            from android.storage import app_storage_path
            from android import mActivity

            context = mActivity.getApplicationContext()
            result = context.getExternalFilesDir(None)
            if result:
                storage_path = str(result.toString())
            else:
                storage_path = app_storage_path()
        else:
            storage_path = os.curdir

        storage_path = os.path.join(storage_path, "App Data")

        temp_path = os.path.join(storage_path, "routine.xlsx")

        if os.path.exists(temp_path):
            # extractor(temp_path)
            loctn[2] = temp_path

            self.b1.disabled = False
            self.b2.disabled = False
            self.b3.disabled = False
        else:
            down = BoxLayout()
            down.txt = MDLabel(text="Download the routine first",
                               size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                               theme_text_color="Custom", text_color="#e6e5e3")
            down.add_widget(down.txt)
            pop = Popup(title="No Routine Found", content=down, size_hint=[.8, .2], pos_hint={"center_x": .5})
            pop.open()


class Help(MDScreen):
    message = hm
    pass


class StudentPage(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)


        box = BoxLayout(padding=10, orientation="vertical")
        self.b2 = MDRectangleFlatButton(text="Next", size=[.25, .06], pos_hint={"center_x": .5, "center_y": .34},
                                        on_press=self.Generate_routine, line_color="#31de8d", md_bg_color="#31de8d",
                                        text_color="#292929", )
        b3 = MDRectangleFlatButton(text="Home", size=[.2, .06], pos_hint={"center_x": .5, "center_y": .2},
                                   on_press=self.change, line_color="#3c3d3d", md_bg_color="#3c3d3d",
                                   text_color="white")



        box.add_widget(
            MDLabel(text="SWE Routine", font_style="H4", pos_hint={"x": .25, "center_y": 1}, size_hint=[1, .1]))

        box.add_widget(MDLabel(size_hint=[.5, .05]))

        self.batch = MDTextField(multiline=False, size_hint=[.6, .1], pos_hint={"center_x": .5, "center_y": .8},
                                 hint_text="Enter Batch")
        box.add_widget(self.batch)

        box.add_widget(MDLabel(size_hint=[.5, .05]))

        self.sec = MDTextField(multiline=False, size_hint=[.25, .06], pos_hint={"center_x": .5, "center_y": .6},
                               hint_text="Enter Section")
        box.add_widget(self.sec)

        box.add_widget(MDLabel(size_hint=[.5, .05]))

        box.add_widget(self.b2)

        box.add_widget(MDLabel(size_hint=[.5, .05]))

        box.add_widget(b3)
        box.add_widget(
            MDLabel(text="0242320005341689", size_hint=[.5, .1], pos_hint={"center_x": .55, "center_y": .08},
                    theme_text_color="Custom", text_color="#171717"))
        self.add_widget(box)

    def Generate_routine(self, obj):
        loc = loctn[2]

        if platform == 'android':
            loc_list = loc.split('/')
            loc_list.pop()

        else:
            loc_list = loc.split('\\')
            loc_list.pop()
        loc_list.pop()
        loc_list.append("DIU Routine")

        loctn[0] = '/'.join(loc_list)
        loctn[1] = str(self.batch.text) + "_" + self.sec.text.upper()
        location = loc
        self.save_loc = '/'.join(loc_list)
        self.section = self.sec.text.upper()

        batch_txt = str(self.batch.text).upper()
        batch = str(batch_txt[0:2]) + self.section


        functions.times,functions.days = functions.routine_separateor(location, batch)
        for keys in functions.days:
            if functions.days[keys] != []:
                day[keys] = functions.days[keys]
        for keys in days:
            functions.days[keys] = []

        Generated = True

        self.manager.current = "student_show"
        # else:
        #     down = BoxLayout()
        #     down.txt = MDLabel(text="Wrong input\n\nTry reading the instruction again",
        #                        size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
        #                        theme_text_color="Custom", text_color="#e6e5e3")
        #     down.add_widget(down.txt)
        #     pop = Popup(title="Input Error", content=down, size_hint=[.8, .25], pos_hint={"center_x": .5})
        #     pop.open()

    def change(self, obj):
        self.manager.current = "Home"


class Student_Routine_Show(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.add_widget(
            MDRectangleFlatButton(text='Show Routine', on_press=self.show, pos_hint={'center_x': .5, 'center_y': .9},
                                  md_bg_color="#176963", line_color="#176963", text_color="white"))
        self.add_widget(
            MDRectangleFlatButton(text='Save PDF', pos_hint={"center_x": .5, "center_y": .2}, on_press=self.pdf_store,
                                  md_bg_color="#31de8d", line_color="#31de8d", text_color="#292929"))
        # self.add_widget(
        #     MDRectangleFlatButton(text='Save Excel', pos_hint={"center_x": .65, "center_y": .2}, on_press=self.excel_store,
        #                           md_bg_color="#31de8d", line_color="#31de8d", text_color="#292929"))
        self.add_widget(
            MDRectangleFlatButton(text='Back', pos_hint={"center_x": .5, "center_y": .1}, on_press=self.change,
                                  md_bg_color="#30c6d1", line_color="#30c6d1", text_color="#292929"))

        self.showed = False

    def show(self, obj):
        if not self.showed:
            times = functions.times
            row = len(day) + 1
            clm = len(times) + 1
            grid_table = MDGridLayout(pos_hint={'center_x': .5, 'center_y': .55}, size_hint=[None, None],
                                      md_bg_color="#2e2e2d", rows=row, cols=clm)
            grid_table.bind(minimum_height=grid_table.setter('height'), minimum_width=grid_table.setter('width'))
            DP = 60
            WDP = 110
            TDP = 10
            line = "#000000"
            txt = MDLabel(text="", size_hint=[None, None], size=[dp(WDP), dp(DP)], halign="center", line_color=line,
                          text_size=(None, dp(TDP)))
            grid_table.add_widget(txt)
            for t in times:
                tim = MDLabel(text=t, size_hint=[None, None], size=[dp(WDP), dp(DP)], halign="center", line_color=line,
                              text_size=(None, dp(TDP)))
                grid_table.add_widget(tim)

            for key in day.keys():
                d = MDLabel(text=key, size_hint=[None, None], size=[dp(WDP), dp(DP)], halign="center", line_color=line,
                            text_size=(None, dp(TDP)))
                grid_table.add_widget(d)

                temp = {}  # creates a dummy dictionary
                dummy = {}  # creates a dummy dictionary

                for lst in day[key]:
                    dummy[lst[1]] = lst

                for i, t in enumerate(times):
                    if t in dummy.keys():
                        temp[i] = dummy[t]

                dummy.clear()

                for i, t in enumerate(times):
                    if i in temp.keys() and temp[i][1] == t:
                        tim = MDLabel(text=f"{temp[i][0]}\nRoom : {temp[i][2]}", size_hint=[None, None],
                                      size=[dp(WDP), dp(DP)], halign="center",
                                      pos_hint={"center_x": .5, "center_y": .5},
                                      line_color=line, text_size=(None, dp(TDP)))
                        grid_table.add_widget(tim)

                    else:
                        txt = MDLabel(text="-", size_hint=[None, None], size=[dp(WDP), dp(DP)], halign="center",
                                      line_color=line,
                                      text_size=(None, dp(TDP)))
                        grid_table.add_widget(txt)
                temp.clear()

            self.table = MDScrollView(pos_hint={'center_x': .5, 'center_y': .55}, size=self.size, size_hint=[.9, .56],
                                      do_scroll_y=True, do_scroll_x=True)
            self.table.add_widget(grid_table)
            self.add_widget(self.table)
            self.showed = True

    def pdf_store(self, obj):
        print(day)

        if not os.path.exists(loctn[0]):
            os.makedirs(loctn[0])
        functions.routine_pdf(day, functions.times, loctn[1], loctn[0])

        if platform == "android":
            from android import autoclass
            Environment = autoclass('android.os.Environment')
            ss = SharedStorage()
            ss.copy_to_shared(f"{loctn[0]}/{loctn[1]}.pdf",
                              collection=Environment.DIRECTORY_DOWNLOADS)

        success = BoxLayout()
        success.txt = MDLabel(text="Successfully saved the routine ",
                              size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                              theme_text_color="Custom", text_color="#e6e5e3")
        success.add_widget(success.txt)
        pop = Popup(title=f"Downloads/DIU Routine/{loctn[1]}.pdf", content=success, size_hint=[.8, .2],
                    pos_hint={"center_x": .5})
        pop.open()

    def excel_store(self, obj):
        if not os.path.exists(loctn[0]):
            os.makedirs(loctn[0])
        functions.routine_excel(day, times, loctn[1], loctn[0])
        success = BoxLayout()
        success.txt = MDLabel(text="Successfully saved the routine ",
                              size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                              theme_text_color="Custom", text_color="#e6e5e3")
        success.add_widget(success.txt)
        pop = Popup(title=f"{loctn[0]}/{loctn[1]}.excel", content=success, size_hint=[.8, .2],
                    pos_hint={"center_x": .5})
        pop.open()

    def change(self, obj):
        if self.showed == True:
            # self.remove_widget(self.scrl)
            self.remove_widget(self.table)
            self.showed = False
        self.manager.current = 'StudentPage'

        day.clear()
        times.clear()


class TeacherPage(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)


        box = BoxLayout(padding=10, orientation="vertical")
        self.b2 = MDRectangleFlatButton(text="Next", size=[.25, .06], pos_hint={"center_x": .5, "center_y": .34},
                                        on_press=self.Generate_routine, line_color="#31de8d", md_bg_color="#31de8d",
                                        text_color="#292929", )
        b3 = MDRectangleFlatButton(text="Home", size=[.2, .06], pos_hint={"center_x": .5, "center_y": .2},
                                   on_press=self.change, line_color="#3c3d3d", md_bg_color="#3c3d3d",
                                   text_color="white")


        box.add_widget(
            MDLabel(text="SWE Routine", font_style="H4", pos_hint={"x": .25, "center_y": 1}, size_hint=[1, .1]))

        box.add_widget(MDLabel(size_hint=[.5, .05]))

        self.teacher_initial = MDTextField(multiline=False, size_hint=[.6, .1],
                                           pos_hint={"center_x": .5, "center_y": .8},
                                           hint_text="Teacher's Initial")
        box.add_widget(self.teacher_initial)

        box.add_widget(MDLabel(size_hint=[.5, .1], pos_hint={"center_x": .5, "center_y": .5}))

        box.add_widget(self.b2)

        box.add_widget(MDLabel(size_hint=[.5, .05]))

        box.add_widget(b3)
        box.add_widget(
            MDLabel(text="0242320005341689", size_hint=[.5, .1], pos_hint={"center_x": .55, "center_y": .08},
                    theme_text_color="Custom", text_color="#171717"))
        self.add_widget(box)

    def Generate_routine(self, obj):
        loc = loctn[2]

        if platform == 'android':
            loc_list = loc.split('/')


            loc_list.pop()

        else:
            loc_list = loc.split('\\')
            loc_list.pop()
        loc_list.pop()
        loc_list.append("DIU Routine")

        loctn[0] = '/'.join(loc_list)
        loctn[1] = self.teacher_initial.text.upper()
        location = loc
        self.save_loc = '/'.join(loc_list)

        ti = str(self.teacher_initial.text).upper()

        functions.times,functions.days = functions.teacher(location, ti)
        for keys in functions.days:
            if functions.days[keys] != []:
                day[keys] = functions.days[keys]
        for keys in functions.days:
            functions.days[keys] = []

        Generated = True

        self.manager.current = "teacher_show"

    def change(self, obj):
        self.manager.current = "Home"


class Teacher_Routine_Show(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.add_widget(
            MDRectangleFlatButton(text='Show Routine', on_press=self.show, pos_hint={'center_x': .5, 'center_y': .9},
                                  md_bg_color="#176963", line_color="#176963", text_color="white"))
        self.add_widget(
            MDRectangleFlatButton(text='Save PDF', pos_hint={"center_x": .5, "center_y": .2}, on_press=self.pdf_store,
                                  md_bg_color="#31de8d", line_color="#31de8d", text_color="#292929"))

        self.add_widget(
            MDRectangleFlatButton(text='Back', pos_hint={"center_x": .5, "center_y": .1}, on_press=self.change,
                                  md_bg_color="#30c6d1", line_color="#30c6d1", text_color="#292929"))

        self.showed = False

    def show(self, obj):
        if not self.showed:
            times = functions.times
            row = len(day) + 1
            clm = len(times) + 1
            grid_table = MDGridLayout(pos_hint={'center_x': .5, 'center_y': .55}, size_hint=[None, None],
                                      md_bg_color="#2e2e2d", rows=row, cols=clm)
            grid_table.bind(minimum_height=grid_table.setter('height'), minimum_width=grid_table.setter('width'))
            DP = 60
            WDP = 110
            TDP = 10
            line = "#000000"
            txt = MDLabel(text="", size_hint=[None, None], size=[dp(WDP), dp(DP)], halign="center", line_color=line,
                          text_size=(None, dp(TDP)))
            grid_table.add_widget(txt)
            for t in times:
                tim = MDLabel(text=t, size_hint=[None, None], size=[dp(WDP), dp(DP)], halign="center", line_color=line,
                              text_size=(None, dp(TDP)))
                grid_table.add_widget(tim)

            for key in day.keys():
                d = MDLabel(text=key, size_hint=[None, None], size=[dp(WDP), dp(DP)], halign="center", line_color=line,
                            text_size=(None, dp(TDP)))
                grid_table.add_widget(d)
                # count = len(day[key])-1
                # end = len(day[key])

                temp = {}  # creates a dummy dictionary
                dummy = {}  # creates a dummy dictionary

                for lst in day[key]:
                    dummy[lst[1]] = lst

                for i, t in enumerate(times):
                    if t in dummy.keys():
                        temp[i] = dummy[t]

                dummy.clear()

                for i, t in enumerate(times):
                    if i in temp.keys() and temp[i][1] == t:
                        tim = MDLabel(text=f"{temp[i][0]}\nRoom : {temp[i][2]}", size_hint=[None, None],
                                      size=[dp(WDP), dp(DP)], halign="center",
                                      pos_hint={"center_x": .5, "center_y": .5},
                                      line_color=line, text_size=(None, dp(TDP)))
                        grid_table.add_widget(tim)

                    else:
                        txt = MDLabel(text="-", size_hint=[None, None], size=[dp(WDP), dp(DP)], halign="center",
                                      line_color=line,
                                      text_size=(None, dp(TDP)))
                        grid_table.add_widget(txt)
                temp.clear()

            self.table = MDScrollView(pos_hint={'center_x': .5, 'center_y': .55}, size=self.size, size_hint=[.9, .56],
                                      do_scroll_y=True, do_scroll_x=True)
            self.table.add_widget(grid_table)
            self.add_widget(self.table)
            self.showed = True

    def pdf_store(self, obj):
        print(day)
        if not os.path.exists(loctn[0]):
            os.makedirs(loctn[0])
        functions.routine_pdf(day, functions.times, loctn[1], loctn[0])

        if platform == "android":
            from android import autoclass
            Environment = autoclass('android.os.Environment')
            ss = SharedStorage()
            ss.copy_to_shared(f"{loctn[0]}/{loctn[1]}.pdf",
                              collection=Environment.DIRECTORY_DOWNLOADS)

        success = BoxLayout()
        success.txt = MDLabel(text="Successfully saved the routine ",
                              size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                              theme_text_color="Custom", text_color="#e6e5e3")
        success.add_widget(success.txt)
        pop = Popup(title=f"Downloads/DIU Routine/{loctn[1]}.pdf", content=success, size_hint=[.8, .2],
                    pos_hint={"center_x": .5})
        pop.open()

    def excel_store(self, obj):
        if not os.path.exists(loctn[0]):
            os.makedirs(loctn[0])
        functions.routine_excel(day, times, loctn[1], loctn[0])
        success = BoxLayout()
        success.txt = MDLabel(text="Successfully saved the routine ",
                              size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                              theme_text_color="Custom", text_color="#e6e5e3")
        success.add_widget(success.txt)
        pop = Popup(title=f"Downloads/DIU Routine/{loctn[1]}.xlsx", content=success, size_hint=[.8, .2],
                    pos_hint={"center_x": .5})
        pop.open()

    def change(self, obj):
        if self.showed == True:
            # self.remove_widget(self.scrl)
            self.remove_widget(self.table)
            self.showed = False
        self.manager.current = 'TeacherPage'

        day.clear()
        times.clear()


class Blank_Slot_Routine(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.add_widget(
            MDRectangleFlatButton(text='Show Routine', on_press=self.show, pos_hint={'center_x': .5, 'center_y': .9},
                                  md_bg_color="#176963", line_color="#176963", text_color="white"))
        self.add_widget(
            MDRectangleFlatButton(text='Save PDF', pos_hint={"center_x": .5, "center_y": .2}, on_press=self.pdf_store,
                                  md_bg_color="#31de8d", line_color="#31de8d", text_color="#292929"))
        # self.add_widget(
        #     MDRectangleFlatButton(text='Save Excel', pos_hint={"center_x": .65, "center_y": .2}, on_press=self.excel_store,
        #                           md_bg_color="#31de8d", line_color="#31de8d", text_color="#292929"))
        self.add_widget(
            MDRectangleFlatButton(text='Home', pos_hint={"center_x": .5, "center_y": .1}, on_press=self.change,
                                  md_bg_color="#30c6d1", line_color="#30c6d1", text_color="#292929"))

        self.showed = False

    def show(self, obj):
        if not self.showed:
            emt_days = functions.emt_days
            times = functions.times

            row = len(emt_days) + 1
            clm = len(times) + 1

            grid_table = MDGridLayout(pos_hint={'center_x': .5, 'center_y': .55}, size_hint=[None, None],
                                      md_bg_color="#2e2e2d", rows=row, cols=clm)
            grid_table.bind(minimum_height=grid_table.setter('height'), minimum_width=grid_table.setter('width'))
            DP = 60
            WDP = 110
            TDP = 10
            line = "#000000"
            txt = MDLabel(text="", size_hint=[None, None], size=[dp(WDP), dp(DP)], halign="center", line_color=line,
                          text_size=(None, dp(TDP)))
            grid_table.add_widget(txt)
            for t in times:
                tim = MDLabel(text=t, size_hint=[None, None], size=[dp(WDP), dp(DP)], halign="center", line_color=line,
                              text_size=(None, dp(TDP)))
                grid_table.add_widget(tim)

            for key in emt_days.keys():
                max_len = 1
                for t in times:
                    max_len = max(max_len, len(emt_days[key][t]) + 1)
                DP = max_len * 20

                d = MDLabel(text=key, size_hint=[None, None], size=[dp(WDP), dp(DP)], halign="center", line_color=line,
                            text_size=(None, dp(TDP)))
                grid_table.add_widget(d)
                # count = len(emt_days[key])-1
                # end = len(emt_days[key])
                for time in emt_days[key]:
                    rooms = ""
                    for room in emt_days[key][time]:
                        rooms += f"\n{room[0]}"
                    rm = MDLabel(text=rooms, size_hint=[None, None], size=[dp(WDP), dp(DP)], halign="center",
                                 line_color=line,
                                 text_size=(None, dp(TDP)))
                    grid_table.add_widget(rm)

            self.table = MDScrollView(pos_hint={'center_x': .5, 'center_y': .55}, size=self.size, size_hint=[.9, .56],
                                      do_scroll_y=True, do_scroll_x=True)
            self.table.add_widget(grid_table)
            self.add_widget(self.table)
            self.showed = True

    def pdf_store(self, obj):
        loc = loctn[2]
        if platform == 'android':
            loc_list = loc.split('/')
            loc_list.pop()

        else:
            loc_list = loc.split('\\')
            loc_list.pop()
        loc_list.pop()
        loc_list.append("DIU Routine")
        loctn[0] = '/'.join(loc_list)

        if not os.path.exists(loctn[0]):
            os.makedirs(loctn[0])
        functions.blank_pdf(functions.emt_days, functions.times, loctn[0])

        if platform == "android":
            from android import autoclass
            Environment = autoclass('android.os.Environment')
            ss = SharedStorage()
            ss.copy_to_shared(f"{loctn[0]}/Empty Slots.pdf",
                              collection=Environment.DIRECTORY_DOWNLOADS)

        success = BoxLayout()
        success.txt = MDLabel(text="Successfully saved the routine ",
                              size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                              theme_text_color="Custom", text_color="#e6e5e3")
        success.add_widget(success.txt)
        pop = Popup(title=f"Downloads/DIU Routine/Empty Slots.pdf", content=success, size_hint=[.8, .2],
                    pos_hint={"center_x": .5})
        pop.open()

    def change(self, obj):
        if self.showed == True:
            self.remove_widget(self.table)
            self.showed = False
        self.manager.current = 'Home'
        for keys in functions.emt_days:
            functions.emt_days[keys] = {}
        times.clear()


class DIU_Routine(MDApp):
    def build(self):
        Window.bind(on_keyboard=self.key)
        self.theme_cls.theme_style = "Dark"
        kk = Builder.load_string(kv)
        kk.message = hm
        return kk
        # return main_manager()

    def on_start(self):
        if platform == "android":
            from android.permissions import request_permissions, Permission
            request_permissions([Permission.WRITE_EXTERNAL_STORAGE])

    def close_dialouge(self):
        self.dialog = MDDialog(
            title="Exit Application?",
            elevation=0,
            buttons=[
                MDRectangleFlatButton(
                    text="Cancel",
                    theme_text_color="Custom",
                    text_color="#121212",
                    on_press=self.dialog_close,
                    md_bg_color="#c97b22",
                    line_color="#c97b22",
                    size_hint=[.5, .7]
                ),
                MDRectangleFlatButton(
                    text="Exit",
                    theme_text_color="Custom",
                    text_color="#121212",
                    on_press=self.app_close,
                    md_bg_color="#c97b22",
                    line_color="#c97b22",
                    size_hint=[.5, .7]
                ),
            ],
        )
        self.dialog.open()

    def dialog_close(self, obj):
        self.dialog.dismiss()

    def app_close(self, obj):
        self.stop()

    def key(self, window, key, scancode, codepoint, modifier):
        if key == 27:
            self.close_dialouge()
            return True
        else:
            return False


if __name__ == "__main__":
    DIU_Routine().run()