import os
import requests
from kivy import platform
from kivy.metrics import dp
from kivymd.uix.datatables import MDDataTable
from kivymd.uix.screenmanager import MDScreenManager
from kivymd.uix.screen import MDScreen
from openpyxl import load_workbook as wb
from openpyxl import Workbook as wrk
from openpyxl.styles import Font, Alignment, Border, Side
from kivymd.app import MDApp
from kivy.uix.popup import Popup
from kivymd.uix.boxlayout import BoxLayout
from kivymd.uix.button import MDRaisedButton
from kivymd.uix.button import MDRectangleFlatButton
from kivymd.uix.textfield import MDTextField
from kivymd.uix.label import MDLabel
from kivymd.uix.scrollview import MDScrollView

hm = 'Follow these steps\n\n1. Click on "Choose Routine" and select your routine(excel format) from folder.\n\n2. Now enter your course codes separeted by spaces or your semester no with major(If you have any.[NM,DS,RE,CS]).\n\nEx: "MAT101 ENG101 AOL101 SE111 SE112_L SE113" or "1" or "7_NM" \nwithout quotation marks.\n\n3. Now enter your section in section box.\n\n4. Click on "Routine Page" button to see and save routine. \n\n5. Click "Show Routine" to see routine and "Save" to save your routine in exccel format in the same directory. \n\nNOTE: This works with only excel file. '

semester = {
    '1': "MAT101 ENG101 AOL101 SE111 SE112_L SE113",
    '2': "MAT102 PHY101 SE121 SE122_L SE123 SE212 SE213",
    '3': "SE131 SE132_L SE133_L SE211 SE222 STA101 BNS101",
    '4': "SE214 SE215_L SE221 SE223 SE224_L SE232 SE233_L GE235 SE532",
    '5': "SE225 SE226_L SE231_L SE234 SE311 SE312 SE313_L GE324",
    '6': "SE321 SE322_L SE323 SE332 SE333 SE334_L SE411 SE544",
    '7_NM': "SE331 EMP101 SE442 SE535 SE447 SE599",
    '7_DS': "SE331 EMP101 SE442 DS331 DS332 DS411 DS412 DS421 DS422",
    '7_CS': "SE331 EMP101 SE442 CS211 CS418 CS422",
    '7_RE': "SE331 EMP101 SE442 RE331 RE411 RE412 RE421 RE422",
    '8_NM': "SE341 SE431",
    '8_DS': "CS334 CS335 CS439",
    '8_CS': "DS432 DS424 DS431",
    '8_RE': "RE423 RE424 RE431"
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
loctn = ['', '']


def routine_separateor(loc, lst):
    sheet = wb(loc)
    routine = sheet.active
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
            if routine[cursor].value in lst:
                if ltr == 'B':
                    time = routine['A' + '3'].value
                else:
                    time = routine[ltr + '3'].value
                room = routine['A' + str(i)].value
                sub = routine[cursor].value
                days["Sat"].append([sub, time, room])
    i = x
    while True:
        if routine['A' + str(i)].value == "Monday":
            break
        for ltr in ltrs:
            cursor = ltr + str(i)
            if routine[cursor].value in lst:
                if ltr == 'B':
                    time = routine['A' + '3'].value
                else:
                    time = routine[ltr + '3'].value
                room = routine['A' + str(i)].value
                sub = routine[cursor].value
                days["Sun"].append([sub, time, room])
        i += 1

    while True:
        if routine['A' + str(i)].value == "Tuesday":
            break
        for ltr in ltrs:
            cursor = ltr + str(i)
            if routine[cursor].value in lst:
                if ltr == 'B':
                    time = routine['A' + '3'].value
                else:
                    time = routine[ltr + '3'].value
                room = routine['A' + str(i)].value
                sub = routine[cursor].value
                days["Mon"].append([sub, time, room])
        i += 1

    while True:
        if routine['A' + str(i)].value == "Wednesday":
            break
        for ltr in ltrs:
            cursor = ltr + str(i)
            if routine[cursor].value in lst:
                if ltr == 'B':
                    time = routine['A' + '3'].value
                else:
                    time = routine[ltr + '3'].value
                room = routine['A' + str(i)].value
                sub = routine[cursor].value
                days["Tue"].append([sub, time, room])
        i += 1

    while True:
        if routine['A' + str(i)].value == "Thursday":
            break
        for ltr in ltrs:
            cursor = ltr + str(i)
            if routine[cursor].value in lst:
                if ltr == 'B':
                    time = routine['A' + '3'].value
                else:
                    time = routine[ltr + '3'].value
                room = routine['A' + str(i)].value
                sub = routine[cursor].value
                days["Wed"].append([sub, time, room])
        i += 1

    while True:
        if routine['A' + str(i)].value == None:
            break
        for ltr in ltrs:
            cursor = ltr + str(i)
            if routine[cursor].value in lst:
                if ltr == 'B':
                    time = routine['A' + '3'].value
                else:
                    time = routine[ltr + '3'].value
                room = routine['A' + str(i)].value
                sub = routine[cursor].value
                days["Thu"].append([sub, time, room])
        i += 1


def create_routine(dic, sec, times, path):
    border = Border(
        left=Side(style='medium'),  # You can use 'thick' for thicker borders
        right=Side(style='medium'),
        top=Side(style='medium'),
        bottom=Side(style='medium')
    )
    workbook = wrk()
    rtn = workbook.active
    rtn.title = "routine"
    lst = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
    clm = len(times) - 1
    row = 2 * len(dic) + 3
    rtn.merge_cells(f"A1:{lst[clm]}1")
    rtn["A1"].value = "Section : " + sec
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


class Front(MDScreen):

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        message=hm
        box = BoxLayout(orientation='vertical', padding=20)
        heading = MDLabel(text="Instructions",font_style= "H3",size_hint= [1,.1],pos_hint= {"x":.15,"center_y": 1})
        box.add_widget(heading)
        box.add_widget(MDLabel(size_hint=[.5,.05],pos_hint={"center_x":.5,"center_y":1}))
        scroll = MDScrollView(size=self.size,size_hint=[1,.75],do_scroll_y = True)
        msg = MDLabel(
                text= hm,
                valign="top",
                size_hint_y= None,
                height=450
                )
        msg.bind(height=self.update_label_height)
        scroll.add_widget(msg)
        box.add_widget(scroll)
        box.add_widget(MDLabel(size_hint= [.5, .05],pos_hint={"center_x": .5, "center_y": 1}))
        box.add_widget(MDRectangleFlatButton(text="Routine Scrapper",pos_hint= {"center_x": .5, "center_y": 1},on_press= self.change))
        box.add_widget(MDLabel(size_hint= [.5, .1],pos_hint= {"center_x": .5, "center_y": 1}))
        self.add_widget(box)
    def change(self,obj):
        self.manager.current = 'function'
    def update_label_height(self, instance, value):
        instance.height = instance.parent.height


class Second(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.orientation = "vertical"
        self.spacing = 15
        self.padding = 20
        b1 = MDRectangleFlatButton(text="Download Routine", size_hint=[.5, .08], pos_hint={"center_x": .5, "center_y": .85},
                            on_press=self.test, md_bg_color="#30c6d1", text_color="#292929")
        self.b2 = MDRectangleFlatButton(text="Routine Page", size_hint=[.5, .07], pos_hint={"center_x": .5, "center_y": .34},
                                 on_press=self.Generate_routine, md_bg_color="#31de8d", text_color="#292929", disabled=True)
        b3 = MDRectangleFlatButton(text="Help", size_hint=[.2, .07], pos_hint={"center_x": .5, "center_y": .2},
                            on_press=self.change, md_bg_color="#3c3d3d", text_color="white")

        self.add_widget(b1)

        self.code_text = MDTextField(multiline=False, size_hint=[.6, .1], pos_hint={"center_x": .5, "center_y": .65},
                                     hint_text="Enter Course Codes or Semester")
        self.add_widget(self.code_text)

        self.sec = MDTextField(multiline=False, size_hint=[.25, .06], pos_hint={"center_x": .5, "center_y": .5},
                               hint_text="Enter Section")
        self.add_widget(self.sec)

        self.add_widget(self.b2)

        self.add_widget(b3)
        self.add_widget(
            MDLabel(text="0242320005341689", size_hint=[.5, .1], pos_hint={"center_x": .55, "center_y": .08},
                    theme_text_color="Custom", text_color="#171717"))

    def test(self, obj):
        if platform == 'android':
            from android.storage import app_storage_path
            from android import mActivity

            context = mActivity.getApplicationContext()
            result = context.getExternalFilesDir(None)  # don't forget the argument
            if result:
                storage_path = str(result.toString())
            else:
                storage_path = app_storage_path()  # NOT SECURE
        else:
            storage_path = "G:\Pycharm\Projects\Routine"

        temp_path = os.path.join(storage_path, "routine.xlsx")

        url = 'https://drive.google.com/uc?id=1e5_vL6oA4OtPBYb8Nyv1kizkdriBZ8eW'
        response = requests.get(url)
        if response.status_code == 200:
            with open(temp_path, 'wb') as file:
                file.write(response.content)
            down = BoxLayout()
            down.txt = MDLabel(text="Downloaded",
                               size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                               theme_text_color="Custom", text_color="#e6e5e3")
            down.add_widget(down.txt)
            pop = Popup(title=temp_path, content=down, size_hint=[.8, .2], pos_hint={"center_x": .5})
            pop.open()
            self.path = temp_path
            self.b2.disabled = False


    def Generate_routine(self, obj):
        loc = self.path
        if platform == 'android':
            loc_list = loc.split('/')
        else:
            loc_list = loc.split('\\')
        loc_list.pop()
        loctn[0] = '/'.join(loc_list)
        loctn[1] = self.sec.text.upper()
        location = loc
        self.save_loc = '/'.join(loc_list)
        self.section = self.sec.text.upper()
        code = self.code_text.text
        if code[0] in semester.keys():
            txt = semester[code[0]].split()
        else:
            txt = code.split()
        subjects = []
        for ele in txt:
            if ele[-1] == 'L':
                lab = ele.split('_')
                subjects.append(lab[0] + self.section + '1')
                subjects.append(lab[0] + self.section + '2')
            else:
                subjects.append(ele + self.section)
        routine_separateor(location, subjects)
        for keys in days:
            if days[keys] != []:
                day[keys] = days[keys]
        for keys in days:
            days[keys] = []

        Generated = True

        self.manager.current = "show"



    def change(self, obj):
        self.manager.current = "front"





class Routine_Show(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.add_widget(
            MDRectangleFlatButton(text='Show Routine', on_press=self.show, pos_hint={'center_x': .5, 'center_y': .9},
                           md_bg_color="#30c6d1", text_color="#292929"))
        self.add_widget(MDRectangleFlatButton(text='Save', pos_hint={"center_x": .5, "center_y": .2}, on_press=self.store,
                                       md_bg_color="#31de8d", text_color="#292929"))
        self.add_widget(
            MDRectangleFlatButton(text='Routine Scrapper', pos_hint={"center_x": .5, "center_y": .1}, on_press=self.change,
                           md_bg_color="#176963", text_color="white"))

    def show(self, obj):
        clmn = [('', dp(22))]
        row = []
        for time in times:
            clmn.append((time, dp(25)))
        for key in day.keys():
            lst = [key]
            k = day[key]
            for time in times:
                hagu = False
                for subs in k:
                    if subs[1] == time:
                        lst.append(f"{subs[0]}\nRoom: {subs[2]}")
                        hagu = True
                        break
                if hagu == False:
                    lst.append('')
            row.append(lst)

        table = MDDataTable(
            pos_hint={'center_x': .5, 'center_y': .55},
            size_hint=[.9, .56],
            column_data=clmn,
            row_data=row,
            rows_num=8
        )
        self.add_widget(table)

    def store(self, obj):

        create_routine(day, loctn[1], times, loctn[0])
        success = BoxLayout()
        success.txt = MDLabel(text="Successfully saved the routine ",
                           size_hint=[.5, 1], size=self.size, pos_hint={"top": 1.1},
                           theme_text_color="Custom", text_color="#e6e5e3")
        success.add_widget(success.txt)
        pop = Popup(title=f"{loctn[0]}/{loctn[1]}.xlsx", content=success, size_hint=[.8, .2], pos_hint={"center_x": .5})
        pop.open()

    def change(self, obj):
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
        self.theme_cls.theme_style = "Dark"
        return main_manager()



if __name__ == "__main__":
    DIU_Routine().run()