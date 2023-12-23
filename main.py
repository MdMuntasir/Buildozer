import os
import requests
from kivy import platform
from kivy.core.window import Window
from kivy.lang import Builder
from kivy.metrics import dp
from kivymd.uix.datatables import MDDataTable
from kivymd.uix.dialog import MDDialog
from kivymd.uix.gridlayout import MDGridLayout
from kivymd.uix.screenmanager import MDScreenManager
from kivymd.uix.screen import MDScreen
from openpyxl import load_workbook as wb
from openpyxl import Workbook as wrk
from openpyxl.styles import Font, Alignment, Border, Side
from kivymd.app import MDApp
from kivy.uix.popup import Popup
from kivymd.uix.boxlayout import BoxLayout
from kivymd.uix.button import MDRectangleFlatButton
from kivymd.uix.textfield import MDTextField
from kivymd.uix.label import MDLabel
from kivymd.uix.scrollview import MDScrollView


hm = 'Follow these steps\n\n1. Click on "Download Routine" wait wait few seconds to download the latest routine. Click "Current Routine" to use the already downloaded routine.' \
     'You need to download the routine at least one time after installation to use "Current Routine"\n\n2. ' \
     'Now enter your batch and section in the batch and section box \n\n3.' \
     'Click on "Routine Page" button to see and save routine. \n\n4. Click "Show Routine" to see routine and "Save" to save your routine in exccel format.'

semester = {}

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
loctn = ['', '']
color = {}

def extractor(loc):

    clr_sheet = wb(loc)
    info_sheet = clr_sheet.worksheets[1]
    c_count = 66
    cd_count = 3

    while info_sheet[chr(c_count)+'1'].value:
        color[str(info_sheet[chr(c_count)+'1'].value)] = info_sheet[chr(c_count)+'1'].fill.fgColor.rgb
        c_count += 1

    while info_sheet["A"+str(cd_count)].value:
        semester[str(info_sheet["A"+str(cd_count)].value)] = info_sheet["B"+str(cd_count)].value
        cd_count += 1
    clr_sheet.close()

def routine_separateor(loc,lst,bach):
    sheet = wb(loc)
    routine = sheet.worksheets[0]
    ltrs = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
    x = 6
    times.append(routine["A3"].value)
    for i in range(2, len(ltrs), 2):
        times.append(routine[ltrs[i] + '3'].value)
    for i in range(x, x + 19):
        if routine['A' + str(i)].value == "Sunday":
            x = i
            break
        for ltr in ltrs:
            cursor = ltr + str(i)
            if routine[cursor].fill and routine[cursor].fill.fgColor.rgb == color[bach] and routine[
                cursor].value in lst:
                if ltr == 'B':
                    time = routine['A' + '3'].value
                else:
                    time = routine[ltr + '3'].value
                room = routine['A' + str(i)].value
                sub = routine[cursor].value
                days["Sat"].append([sub, time, str(room)])

    i = x
    while True:
        if routine['A' + str(i)].value == "Monday":
            break
        for ltr in ltrs:
            cursor = ltr + str(i)
            if routine[cursor].fill and routine[cursor].fill.fgColor.rgb == color[bach] and routine[
                cursor].value in lst:
                if ltr == 'B':
                    time = routine['A' + '3'].value
                else:
                    time = routine[ltr + '3'].value
                room = routine['A' + str(i)].value
                sub = routine[cursor].value
                days["Sun"].append([sub, time, str(room)])
        i += 1

    while True:
        if routine['A' + str(i)].value == "Tuesday":
            break
        for ltr in ltrs:
            cursor = ltr + str(i)
            if routine[cursor].fill and routine[cursor].fill.fgColor.rgb == color[bach] and routine[
                cursor].value in lst:
                if ltr == 'B':
                    time = routine['A' + '3'].value
                else:
                    time = routine[ltr + '3'].value
                room = routine['A' + str(i)].value
                sub = routine[cursor].value
                days["Mon"].append([sub, time, str(room)])
        i += 1

    while True:
        if routine['A' + str(i)].value == "Wednesday":
            break
        for ltr in ltrs:
            cursor = ltr + str(i)
            if routine[cursor].fill and routine[cursor].fill.fgColor.rgb == color[bach] and routine[
                cursor].value in lst:
                if ltr == 'B':
                    time = routine['A' + '3'].value
                else:
                    time = routine[ltr + '3'].value
                room = routine['A' + str(i)].value
                sub = routine[cursor].value
                days["Tue"].append([sub, time, str(room)])
        i += 1

    while True:
        if routine['A' + str(i)].value == "Thursday":
            break
        for ltr in ltrs:
            cursor = ltr + str(i)
            if routine[cursor].fill and routine[cursor].fill.fgColor.rgb == color[bach] and routine[
                cursor].value in lst:
                if ltr == 'B':
                    time = routine['A' + '3'].value
                else:
                    time = routine[ltr + '3'].value
                room = routine['A' + str(i)].value
                sub = routine[cursor].value
                days["Wed"].append([sub, time, str(room)])
        i += 1

    while True:
        if routine['A' + str(i)].value == None:
            break
        for ltr in ltrs:
            cursor = ltr + str(i)
            if routine[cursor].fill and routine[cursor].fill.fgColor.rgb == color[bach] and routine[
                cursor].value in lst:
                if ltr == 'B':
                    time = routine['A' + '3'].value
                else:
                    time = routine[ltr + '3'].value
                room = routine['A' + str(i)].value
                sub = routine[cursor].value
                days["Thu"].append([sub, time, str(room)])
        i += 1


def create_routine(dic, times, sec, path):
    border = Border(
        left=Side(style='medium'),
        right=Side(style='medium'),
        top=Side(style='medium'),
        bottom=Side(style='medium')
    )
    workbook = wrk()
    rtn = workbook.active
    rtn.title=sec
    lst = [chr(i) for i in range(66,66+len(times))]

    for cr in lst:
        rtn[cr+"1"].border = border
    clm = len(times)-1
    row = 2*len(dic)+3
    for i in range(1,row):
        rtn["A"+str(i)].border = border

    rtn.merge_cells(f"A1:{lst[clm]}1")
    rtn["A1"].value = sec
    rtn["A1"].font = Font(name="Times New Roman", size=20, bold=True)
    rtn["A1"].alignment = Alignment(horizontal="center")
    rtn["A1"].border = border
    rtn[lst[clm] + '1'].border = border
    rtn.column_dimensions["A"].width = 16.5
    for i in range(clm + 1):
        rtn.column_dimensions[lst[i]].width = 17

    i = 0
    for time in times:
        rtn[lst[i] + '2'].value = time
        rtn[lst[i] + '2'].font = Font(name="Times New Roman", size=14, bold=True)
        rtn[lst[i] + '2'].alignment = Alignment(vertical="center")
        rtn[lst[i] + '2'].alignment = Alignment(horizontal="center")
        rtn[lst[i] + '2'].border = border
        i += 1
    i = 3

    for key in dic.keys():
        rtn['A' + str(i + 1)].border = border
        rtn.merge_cells(f"A{str(i)}:A{str(i + 1)}")
        rtn['A' + str(i)].value = key
        rtn['A' + str(i)].font = Font(name="Times New Roman", size=14, bold=True)
        rtn['A' + str(i)].alignment = Alignment(vertical="center")
        rtn['A' + str(i)].alignment = Alignment(horizontal="center")
        rtn['A' + str(i)].border = border
        i += 2

    i = 0
    for j in range(3, row, 2):
        key = dic[rtn["A" + str(j)].value]
        for k in range(1, clm + 2):
            c = lst[k - 1]
            rtn[c + str(j)].border = border
            rtn[c + str(j + 1)].border = border
            for x in key:
                if rtn[c + '2'].value == x[1]:
                    rtn[c + str(j)].value = str(x[0])
                    rtn[c + str(j + 1)].value = f"(Room : {str(x[2])})"
                    rtn[c + str(j)].font = Font(name="Times New Roman", size=13.5)
                    rtn[c + str(j + 1)].font = Font(name="Times New Roman", size=12)
                    rtn[c + str(j)].alignment = Alignment(horizontal="center")
                    rtn[c + str(j + 1)].alignment = Alignment(horizontal="center")

                    break

        i += 1

    workbook.save(f"{path}/{sec}.xlsx")





Generated = False





kv = '''

ScreenManager:
    Front:
    Second:
    Routine_Show:
    

<Front>:

    name: "front"
    
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
            text: "Routine Scrapper"
            pos_hint: {"center_x":.5,"center_y":1}
            on_press: root.manager.current = "function"
        
        MDLabel:
            size_hint: .5,.1
            pos_hint: {"center_x":.5,"center_y":1}
            
            
<Second>:
    name: "function"
    
<Routine_Show>:
    name: "show"


<Scroll@MDScrollView>

'''






class Front(MDScreen):
    message = hm
    pass


class Second(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.orientation = "vertical"
        self.spacing = 15
        self.padding = 20
        b1 = MDRectangleFlatButton(text="Download Routine",line_color="#30c6d1" ,size_hint=[.5, .04], pos_hint={"center_x": .5, "center_y": .85},
                            on_press=self.download, md_bg_color="#30c6d1", text_color="#292929")
        b12 = MDRectangleFlatButton(text="Current Routine",line_color="#c97b22" ,size_hint=[.5, .04], pos_hint={"center_x": .5, "center_y": .75},
                            on_press=self.current ,md_bg_color="#c97b22", text_color="#121212")
        self.b2 = MDRectangleFlatButton(text="Next", size_hint=[.3, .07], pos_hint={"center_x": .5, "center_y": .34},
                                 on_press=self.Generate_routine,line_color="#31de8d", md_bg_color="#31de8d", text_color="#292929", disabled=True)
        b3 = MDRectangleFlatButton(text="Help", size_hint=[.2, .07], pos_hint={"center_x": .5, "center_y": .2},
                            on_press=self.change,line_color="#3c3d3d" ,md_bg_color="#3c3d3d", text_color="white")

        self.add_widget(b1)
        self.add_widget(b12)

        self.batch = MDTextField(multiline=False, size_hint=[.6, .1], pos_hint={"center_x": .5, "center_y": .65},
                                     hint_text="Enter Batch")
        self.add_widget(self.batch)

        self.sec = MDTextField(multiline=False, size_hint=[.25, .06], pos_hint={"center_x": .5, "center_y": .5},
                               hint_text="Enter Section")
        self.add_widget(self.sec)

        self.add_widget(self.b2)

        self.add_widget(b3)
        self.add_widget(
            MDLabel(text="0242320005341689", size_hint=[.5, .1], pos_hint={"center_x": .55, "center_y": .08},
                    theme_text_color="Custom", text_color="#171717"))

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
            storage_path = "G:\Pycharm\Projects\Routine"

        if not os.path.exists(os.path.join(storage_path, "App Data")):
            os.makedirs(os.path.join(storage_path, "App Data"))
        storage_path = os.path.join(storage_path, "App Data")

        temp_path = os.path.join(storage_path, "routine.xlsx")
        cd_path = os.path.join(storage_path, "codes.xlsx")

        rtn = 'https://drive.google.com/uc?id=1e5_vL6oA4OtPBYb8Nyv1kizkdriBZ8eW'

        rtn_response = requests.get(rtn)

        if rtn_response.status_code == 200 :
            with open(temp_path, 'wb') as file:
                file.write(rtn_response.content)

            down = BoxLayout()
            down.txt = MDLabel(text="Downloaded",
                               size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                               theme_text_color="Custom", text_color="#e6e5e3")
            down.add_widget(down.txt)
            pop = Popup(title=storage_path, content=down, size_hint=[.8, .2], pos_hint={"center_x": .5})
            extractor(temp_path)
            pop.open()
            self.path = temp_path
            self.b2.disabled = False

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
            storage_path = "G:\Pycharm\Projects\Routine"

        storage_path = os.path.join(storage_path, "App Data")

        temp_path = os.path.join(storage_path, "routine.xlsx")


        if os.path.exists(temp_path):

            extractor(temp_path)
            self.path = temp_path
            self.b2.disabled = False
        else:
            down = BoxLayout()
            down.txt = MDLabel(text="Download the routine first",
                               size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                               theme_text_color="Custom", text_color="#e6e5e3")
            down.add_widget(down.txt)
            pop = Popup(title="No Routine Found", content=down, size_hint=[.8, .2], pos_hint={"center_x": .5})
            pop.open()

    def Generate_routine(self, obj):
        loc = self.path
        if platform == 'android':
            loc_list = loc.split('/')
        else:
            loc_list = loc.split('\\')
        loc_list.pop()
        loc_list.pop()
        loctn[0] = '/'.join(loc_list)
        loctn[1] = str(self.batch.text) + "_" +self.sec.text.upper()
        location = loc
        self.save_loc = '/'.join(loc_list)
        self.section = self.sec.text.upper()
        
        batch_txt = str(self.batch.text).upper()
        batch = str(batch_txt[0:2])
        if batch_txt in semester.keys() and batch in color.keys():
            txt = semester[batch_txt].split()

            subjects = []
            for ele in txt:
                if ele[-1] == 'L':
                    lab = ele.split('_')
                    subjects.append(lab[0] + self.section + '1')
                    subjects.append(lab[0] + self.section + '2')
                else:
                    subjects.append(ele + self.section)
            routine_separateor(location, subjects, batch)
            for keys in days:
                if days[keys] != []:
                    day[keys] = days[keys]
            for keys in days:
                days[keys] = []

            Generated = True



            self.manager.current = "show"
        else:
            down = BoxLayout()
            down.txt = MDLabel(text="Wrong input\n\nTry reading the instruction again",
                               size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                               theme_text_color="Custom", text_color="#e6e5e3")
            down.add_widget(down.txt)
            pop = Popup(title="Input Error", content=down, size_hint=[.8, .25], pos_hint={"center_x": .5})
            pop.open()
            print(color)
            print(semester)
            print(batch,batch_txt)





    def change(self, obj):
        self.manager.current = "front"





class Routine_Show(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.add_widget(
            MDRectangleFlatButton(text='Show Routine', on_press=self.show, pos_hint={'center_x': .5, 'center_y': .9},
                           md_bg_color="#176963",line_color="#176963" ,text_color="white"))
        self.add_widget(MDRectangleFlatButton(text='Save', pos_hint={"center_x": .5, "center_y": .2}, on_press=self.store,
                                       md_bg_color="#31de8d" ,line_color="#31de8d"  ,text_color="#292929"))
        self.add_widget(
            MDRectangleFlatButton(text='Back', pos_hint={"center_x": .5, "center_y": .1}, on_press=self.change,
                           md_bg_color="#30c6d1",line_color="#30c6d1" ,text_color="#292929"))

        self.showed = False
    def show(self, obj):

        row = len(day)+1
        clm = len(times)+1
        table = MDGridLayout(pos_hint={'center_x':.5,'center_y':.55},size_hint = [None,None] ,md_bg_color="#2e2e2d",rows= row,cols= clm)
        table.bind(minimum_height=table.setter('height'),minimum_width=table.setter('width'))
        DP = 60
        WDP = 110
        TDP = 10
        line = "#000000"
        txt = MDLabel(text="", size_hint=[None,None],size=[dp(WDP),dp(DP)], halign="center", line_color=line, text_size=(None, dp(TDP)))
        table.add_widget(txt)
        for t in times:
            tim = MDLabel(text=t, size_hint=[None,None], size=[dp(WDP), dp(DP)], halign="center", line_color=line, text_size=(None, dp(TDP)))
            table.add_widget(tim)

        for key in day.keys():
            d = MDLabel(text=key, size_hint=[None,None], size=[dp(WDP),dp(DP)], halign="center", line_color=line,  text_size=(None, dp(TDP)))
            table.add_widget(d)
            
            temp = {} #creates a dummy dictionary
            dummy = {} #creates a dummy dictionary

            for lst in day[key]:
                dummy[lst[1]] = lst

            for i,t in enumerate(times):
                if t in dummy.keys():
                    temp[i] = dummy[t]

            dummy.clear()

            for i,t in enumerate(times):
                if i in temp.keys() and temp[i][1] == t:
                    tim = MDLabel(text=f"{temp[i][0]}\nRoom : {temp[i][2]}", size_hint=[None,None], size=[dp(WDP),dp(DP)], halign="center", pos_hint={"center_x": .5, "center_y": .5},
                                  line_color=line, text_size=(None, dp(TDP)))
                    table.add_widget(tim)

                else:
                    txt = MDLabel(text="-", size_hint=[None,None], size=[dp(WDP),dp(DP)], halign="center", line_color=line,
                                  text_size=(None, dp(TDP)))
                    table.add_widget(txt)
            temp.clear()


        self.scrl = MDScrollView(pos_hint={'center_x':.5,'center_y':.55},size=self.size,size_hint=[.9,.56],do_scroll_y = True,do_scroll_x=True)
        self.scrl.add_widget(table)
        self.add_widget(self.scrl)
        self.showed = True

    def store(self, obj):

        create_routine(day, times,loctn[1], loctn[0])
        success = BoxLayout()
        success.txt = MDLabel(text="Successfully saved the routine ",
                           size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                           theme_text_color="Custom", text_color="#e6e5e3")
        success.add_widget(success.txt)
        pop = Popup(title=f"{loctn[0]}/{loctn[1]}.xlsx", content=success, size_hint=[.8, .2], pos_hint={"center_x": .5})
        pop.open()

    def change(self, obj):
        if self.showed==True:
            self.remove_widget(self.scrl)
        self.manager.current = 'function'

        day.clear()
        times.clear()

class main_manager(MDScreenManager):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.add_widget(Front(name='front'))
        self.add_widget(Second(name='function'))
        self.add_widget(Routine_Show(name='show'))


class DIU_Routine(MDApp):
    def build(self):
        Window.bind(on_keyboard=self.key)
        self.theme_cls.theme_style= "Dark"
        kk = Builder.load_string(kv)
        kk.message=hm
        return kk

    def close_dialouge(self):

        self.dialog = MDDialog(
            text="Exit Application?",
            buttons=[
                MDRectangleFlatButton(
                    text="Cancel",
                    theme_text_color="Custom",
                    text_color="#121212",
                    on_press=self.dialog_close,
                    md_bg_color="#c97b22",
                    line_color="#c97b22"
                ),
                MDRectangleFlatButton(
                    text="Close",
                    theme_text_color="Custom",
                    text_color="#121212",
                    on_press=self.app_close,
                    md_bg_color="#c97b22",
                    line_color="#c97b22"
                ),
            ],
        )
        self.dialog.open()
    def dialog_close(self):
        self.dialog.dismiss()
    def app_close(self):
        self.stop()

    def key(self,window,key,scancode,codepoint,modifier):
        if key==27:
            self.close_dialouge()
            return True
        else:
            return False




if __name__ == "__main__":
    DIU_Routine().run()
