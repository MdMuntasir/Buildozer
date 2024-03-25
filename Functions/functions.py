from fpdf import FPDF
from openpyxl import load_workbook as wb
from openpyxl import Workbook as wrk
from openpyxl.styles import Font, Alignment, Border, Side


class functions:
    times = []
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

    dummy = {}

    # Extracts color of each semester information
    def extractor(loc):
        color = {}
        semester = {}

        clr_sheet = wb(loc)
        info_sheet = clr_sheet.worksheets[1]
        c_count = 66
        cd_count = 3

        while info_sheet[chr(c_count) + '1'].value:
            color[str(info_sheet[chr(c_count) + '1'].value)] = info_sheet[chr(c_count) + '1'].fill.fgColor.rgb
            c_count += 1

        while info_sheet["A" + str(cd_count)].value:
            semester[str(info_sheet["A" + str(cd_count)].value)] = info_sheet["B" + str(cd_count)].value
            cd_count += 1
        clr_sheet.close()

        return color, semester

    # This generates individual section routine
    def routine_separateor(loc, bach):
        times = functions.times
        days = functions.days

        sheet = wb(loc)
        routine = sheet.worksheets[0]
        dimension = routine.calculate_dimension().split(':')
        last_ltr = ord(dimension[1][0])

        ltrs = [chr(i) for i in range(66, last_ltr + 1)]
        x = 6

        if times == []:
            times.append(routine["A3"].value)
            for i in range(2, len(ltrs), 2):
                times.append(routine[ltrs[i] + '3'].value)

        while True:
            if routine['A' + str(x)].value == "Sunday":
                break
            for ltr in ltrs:
                cursor = ltr + str(x)
                if routine[cursor].value:
                    temp = routine[cursor].value.split('-')
                    if len(temp) == 2:
                        if len(temp[1]) == 4:
                            temp[0] += temp[1][2:5]
                        if temp[1][0:3] == bach:
                            # if not color:
                            #     color.append(routine[cursor].fill.fgColor.rgb)
                            if ltr == 'B':
                                time = routine['A' + '3'].value
                            else:
                                time = routine[ltr + '3'].value
                            room = routine['A' + str(x)].value
                            sub = temp[0]
                            days["Sat"].append([sub, time, str(room)])
            x += 1

        i = x
        while True:
            if routine['A' + str(i)].value == "Monday":
                break
            for ltr in ltrs:
                cursor = ltr + str(i)
                if routine[cursor].value:
                    temp = routine[cursor].value.split('-')
                    if len(temp) == 2:
                        if len(temp[1]) == 4:
                            temp[0] += temp[1][2:5]
                        if temp[1][0:3] == bach:
                            # if not color:
                            #     color.append(routine[cursor].fill.fgColor.rgb)
                            if ltr == 'B':
                                time = routine['A' + '3'].value
                            else:
                                time = routine[ltr + '3'].value
                            room = routine['A' + str(i)].value
                            sub = temp[0]
                            days["Sun"].append([sub, time, str(room)])
            i += 1

        while True:
            if routine['A' + str(i)].value == "Tuesday":
                break
            for ltr in ltrs:
                cursor = ltr + str(i)
                if routine[cursor].value:
                    temp = routine[cursor].value.split('-')
                    if len(temp) == 2:
                        if len(temp[1]) == 4:
                            temp[0] += temp[1][2:5]
                        if temp[1][0:3] == bach:
                            # if not color:
                            #     color.append(routine[cursor].fill.fgColor.rgb)
                            if ltr == 'B':
                                time = routine['A' + '3'].value
                            else:
                                time = routine[ltr + '3'].value
                            room = routine['A' + str(i)].value
                            sub = temp[0]
                            days["Mon"].append([sub, time, str(room)])
            i += 1

        while True:
            if routine['A' + str(i)].value == "Wednesday":
                break
            for ltr in ltrs:
                cursor = ltr + str(i)
                if routine[cursor].value:
                    temp = routine[cursor].value.split('-')
                    if len(temp) == 2:
                        if len(temp[1]) == 4:
                            temp[0] += temp[1][2:5]
                        if temp[1][0:3] == bach:
                            # if not color:
                            #     color.append(routine[cursor].fill.fgColor.rgb)
                            if ltr == 'B':
                                time = routine['A' + '3'].value
                            else:
                                time = routine[ltr + '3'].value
                            room = routine['A' + str(i)].value
                            sub = temp[0]
                            days["Tue"].append([sub, time, str(room)])
            i += 1

        while True:
            if routine['A' + str(i)].value == "Thursday":
                break
            for ltr in ltrs:
                cursor = ltr + str(i)
                if routine[cursor].value:
                    temp = routine[cursor].value.split('-')
                    if len(temp) == 2:
                        if len(temp[1]) == 4:
                            temp[0] += temp[1][2:5]
                        if temp[1][0:3] == bach:
                            # if not color:
                            #     color.append(routine[cursor].fill.fgColor.rgb)
                            if ltr == 'B':
                                time = routine['A' + '3'].value
                            else:
                                time = routine[ltr + '3'].value
                            room = routine['A' + str(i)].value
                            sub = temp[0]
                            days["Wed"].append([sub, time, str(room)])
            i += 1

        while True:
            if routine['A' + str(i)].value == None:
                break
            for ltr in ltrs:
                cursor = ltr + str(i)
                if routine[cursor].value:
                    temp = routine[cursor].value.split('-')
                    if len(temp) == 2:
                        if len(temp[1]) == 4:
                            temp[0] += temp[1][2:5]
                        if temp[1][0:3] == bach:
                            # if not color:
                            #     color.append(routine[cursor].fill.fgColor.rgb)
                            if ltr == 'B':
                                time = routine['A' + '3'].value
                            else:
                                time = routine[ltr + '3'].value
                            room = routine['A' + str(i)].value
                            sub = temp[0]
                            days["Thu"].append([sub, time, str(room)])
            i += 1
        return times, days

    # This generates teacher routine with teacher initial
    def teacher(loc, ti):
        times = functions.times
        days = functions.days

        sheet = wb(loc)
        routine = sheet.worksheets[0]
        dimension = routine.calculate_dimension().split(':')
        last_ltr = ord(dimension[1][0])
        ltrs = [chr(i) for i in range(66, last_ltr + 1)]
        x = 6

        if times == []:
            times.append(routine["A3"].value)
            for i in range(2, len(ltrs), 2):
                times.append(routine[ltrs[i] + '3'].value)

        while True:
            if routine['A' + str(x)].value == "Sunday":
                break
            for index, ltr in enumerate(ltrs):
                cursor = ltr + str(x)
                clmn = ltrs[index - 1]
                if routine[cursor].value == ti:
                    if clmn == 'B':
                        time = routine['A' + '3'].value
                    else:
                        time = routine[clmn + '3'].value
                    room = routine['A' + str(x)].value
                    temp = routine[clmn + str(x)].value.split('-')
                    if len(temp[1]) == 4:
                        temp[0] += temp[1][2:5]
                    else:
                        temp[0] += temp[1][2:4]
                    sub = temp[0]
                    days["Sat"].append([sub, time, str(room)])
            x += 1

        i = x
        while True:
            if routine['A' + str(i)].value == "Monday":
                break
            for index, ltr in enumerate(ltrs):
                cursor = ltr + str(i)
                clmn = ltrs[index - 1]
                if routine[cursor].value == ti:
                    if clmn == 'B':
                        time = routine['A' + '3'].value
                    else:
                        time = routine[clmn + '3'].value
                    room = routine['A' + str(i)].value
                    temp = routine[clmn + str(i)].value.split('-')
                    if len(temp[1]) == 4:
                        temp[0] += temp[1][2:5]
                    else:
                        temp[0] += temp[1][2:4]
                    sub = temp[0]
                    days["Sun"].append([sub, time, str(room)])
            i += 1

        while True:
            if routine['A' + str(i)].value == "Tuesday":
                break
            for index, ltr in enumerate(ltrs):
                cursor = ltr + str(i)
                clmn = ltrs[index - 1]
                if routine[cursor].value == ti:
                    if clmn == 'B':
                        time = routine['A' + '3'].value
                    else:
                        time = routine[clmn + '3'].value
                    room = routine['A' + str(i)].value
                    temp = routine[clmn + str(i)].value.split('-')
                    if len(temp[1]) == 4:
                        temp[0] += temp[1][2:5]
                    else:
                        temp[0] += temp[1][2:4]
                    sub = temp[0]
                    days["Mon"].append([sub, time, str(room)])
            i += 1

        while True:
            if routine['A' + str(i)].value == "Wednesday":
                break
            for index, ltr in enumerate(ltrs):
                cursor = ltr + str(i)
                clmn = ltrs[index - 1]
                if routine[cursor].value == ti:
                    if clmn == 'B':
                        time = routine['A' + '3'].value
                    else:
                        time = routine[clmn + '3'].value
                    room = routine['A' + str(i)].value
                    temp = routine[clmn + str(i)].value.split('-')
                    if len(temp[1]) == 4:
                        temp[0] += temp[1][2:5]
                    else:
                        temp[0] += temp[1][2:4]
                    sub = temp[0]
                    days["Tue"].append([sub, time, str(room)])
            i += 1

        while True:
            if routine['A' + str(i)].value == "Thursday":
                break
            for index, ltr in enumerate(ltrs):
                cursor = ltr + str(i)
                clmn = ltrs[index - 1]
                if routine[cursor].value == ti:
                    if clmn == 'B':
                        time = routine['A' + '3'].value
                    else:
                        time = routine[clmn + '3'].value
                    room = routine['A' + str(i)].value
                    temp = routine[clmn + str(i)].value.split('-')
                    if len(temp[1]) == 4:
                        temp[0] += temp[1][2:5]
                    else:
                        temp[0] += temp[1][2:4]
                    sub = temp[0]
                    days["Wed"].append([sub, time, str(room)])
            i += 1

        while True:
            if routine['A' + str(i)].value == None:
                break
            for index, ltr in enumerate(ltrs):
                cursor = ltr + str(i)
                clmn = ltrs[index - 1]
                if routine[cursor].value == ti:
                    if clmn == 'B':
                        time = routine['A' + '3'].value
                    else:
                        time = routine[clmn + '3'].value
                    room = routine['A' + str(i)].value
                    temp = routine[clmn + str(i)].value.split('-')
                    if len(temp[1]) == 4:
                        temp[0] += temp[1][2:5]
                    else:
                        temp[0] += temp[1][2:4]
                    sub = temp[0]
                    days["Thu"].append([sub, time, str(room)])
            i += 1

        return times, days

    # This genarates empty slot info
    def empty_slot(loc):
        times = functions.times
        emt_days = functions.emt_days

        sheet = wb(loc)
        routine = sheet.worksheets[0]
        dimension = routine.calculate_dimension().split(':')
        last_ltr = ord(dimension[1][0])
        ltrs = [chr(i) for i in range(66, last_ltr)]
        x = 6

        if times == []:  # Insert times in the times list from excel
            times.append(routine["A3"].value)
            for i in range(2, len(ltrs), 2):
                times.append(routine[ltrs[i] + '3'].value)

        for key in emt_days:
            for t in times:
                emt_days[key][t] = []

        while True:
            if routine['A' + str(x)].value == "Sunday":
                break
            for index, ltr in enumerate(ltrs):
                cursor = ltr + str(x)
                clmn = ltrs[index]
                if routine[clmn + str(x)].value == None and routine[cursor].value == None:
                    if clmn == 'B':
                        time = routine['A' + '3'].value
                    else:
                        time = routine[clmn + '3'].value
                    room = routine['A' + str(x)].value
                    if time != None and str(room) != "Online":
                        emt_days["Sat"][time].append([str(room)])
            x += 1

        i = x + 1
        while True:
            if routine['A' + str(i)].value == "Monday":
                break
            for index, ltr in enumerate(ltrs):
                cursor = ltr + str(i)
                clmn = ltrs[index]
                if routine[clmn + str(i)].value == None and routine[cursor].value == None:
                    if clmn == 'B':
                        time = routine['A' + '3'].value
                    else:
                        time = routine[clmn + '3'].value
                    room = routine['A' + str(i)].value
                    if time != None and str(room) != "Online":
                        emt_days["Sun"][time].append([str(room)])
            i += 1
        i += 1
        while True:
            if routine['A' + str(i)].value == "Tuesday":
                break
            for index, ltr in enumerate(ltrs):
                cursor = ltr + str(i)
                clmn = ltrs[index]
                if routine[clmn + str(i)].value == None and routine[cursor].value == None:
                    if clmn == 'B':
                        time = routine['A' + '3'].value
                    else:
                        time = routine[clmn + '3'].value
                    room = routine['A' + str(i)].value
                    if time != None and str(room) != "Online":
                        emt_days["Mon"][time].append([str(room)])
            i += 1
        i += 1
        while True:
            if routine['A' + str(i)].value == "Wednesday":
                break
            for index, ltr in enumerate(ltrs):
                cursor = ltr + str(i)
                clmn = ltrs[index]
                if routine[clmn + str(i)].value == None and routine[cursor].value == None:
                    if clmn == 'B':
                        time = routine['A' + '3'].value
                    else:
                        time = routine[clmn + '3'].value
                    room = routine['A' + str(i)].value
                    if time != None and str(room) != "Online":
                        emt_days["Tue"][time].append([str(room)])
            i += 1
        i += 1
        while True:
            if routine['A' + str(i)].value == "Thursday":
                break
            for index, ltr in enumerate(ltrs):
                cursor = ltr + str(i)
                clmn = ltrs[index]
                if routine[clmn + str(i)].value == None and routine[cursor].value == None:
                    if clmn == 'B':
                        time = routine['A' + '3'].value
                    else:
                        time = routine[clmn + '3'].value
                    room = routine['A' + str(i)].value
                    if time != None and str(room) != "Online":
                        emt_days["Wed"][time].append([str(room)])
            i += 1
        i += 1
        while True:
            if routine['A' + str(i)].value == None:
                break
            for index, ltr in enumerate(ltrs):
                cursor = ltr + str(i)
                clmn = ltrs[index]
                if routine[clmn + str(i)].value == None and routine[cursor].value == None:
                    if clmn == 'B':
                        time = routine['A' + '3'].value
                    else:
                        time = routine[clmn + '3'].value
                    room = routine['A' + str(i)].value
                    if time != None and str(room) != "Online":
                        emt_days["Thu"][time].append([str(room)])
            i += 1

        return times, emt_days

    # Creates routine of blank slots as pdf file
    def blank_pdf(dic, times, path):
        tbl = []
        for key in dic.keys():
            lst = [key]
            for time in dic[key]:
                rooms = "Rooms:"
                for room in dic[key][time]:
                    rooms += f"\n{room[0]}"
                lst.append(rooms)
            tbl.append(lst)

        class PDF(FPDF):
            def header(self):
                self.set_font("Times", 'B', size=23)
                self.set_fill_color(107, 106, 106)
                self.set_text_color(235, 235, 235)
                self.cell(0, 15, txt="Empty Slots", align="C", border=1, fill=True)
                self.ln(15)

        pdf = PDF('P', 'mm', 'Letter')
        pdf.add_page()
        pdf.title = "Empty Slots"

        pdf.set_font("Times", 'B', size=8)

        x = 10
        y = 35
        cell_len = 25.115
        day_len = 20
        pdf.set_fill_color(56, 56, 56)
        pdf.set_text_color(230, 230, 230)
        pdf.cell(day_len, 10, txt='', border=1, fill=True)
        pdf.set_x(x + day_len)
        x += day_len
        for t in times:
            pdf.cell(cell_len, 10, txt=t, border=1, fill=True, align="C")
            pdf.set_x(x + cell_len)
            x += cell_len

        pdf.set_xy(10, y)
        pdf.set_fill_color(84, 84, 84)
        for subs in tbl:
            max_len = 1
            for t in times:
                max_len = max(max_len, len(dic[subs[0]][t]) + 1)
            highest_num = max_len * 2.9
            x = 10
            pdf.set_text_color(230, 230, 230)
            pdf.multi_cell(day_len, highest_num, txt=subs[0], border=1, fill=True, align="C")
            pdf.set_xy(x + day_len, y)
            x += day_len
            pdf.set_text_color(15, 15, 15)
            for i in range(1, len(subs)):
                num = len(dic[subs[0]][times[i - 1]]) + 1
                single_cell_height = highest_num / num
                if subs[i] != '':
                    pdf.multi_cell(cell_len, single_cell_height, txt=subs[i], border=1, align="C")
                else:
                    pdf.multi_cell(cell_len, single_cell_height, txt=subs[i], border=1, align="C")
                pdf.set_xy(x + cell_len, y)
                x += cell_len
            y += highest_num
            pdf.set_xy(10, y)

        pdf.output(path + "/Empty Slots.pdf")

    # Creates routine as pdf file
    def routine_pdf( dic, times, sec, path):
        tbl = []

        for key in dic.keys():
            lst = [key]
            k = dic[key]
            for time in times:
                fnd = False
                for subs in k:
                    if subs[1] == time:
                        lst.append(f"{subs[0]}\n{str(subs[2])}")
                        fnd = True
                        break
                if fnd == False:
                    lst.append('')
            tbl.append(lst)

        class PDF(FPDF):
            def header(self):
                self.set_font("Times", 'B', size=28)
                self.set_fill_color(107, 106, 106)
                self.set_text_color(235, 235, 235)
                self.cell(0, 15, txt=sec, align="C", border=1, fill=True)
                self.ln(15)

        pdf = PDF('P', 'mm', 'Letter')
        pdf.add_page()
        pdf.title = sec

        pdf.set_font("Times", 'B', size=10)
        x = 10
        y = 35
        cell_len = 25.1
        day_len = 20
        cell_height = 16
        pdf.set_fill_color(56, 56, 56)
        pdf.set_text_color(230, 230, 230)
        pdf.cell(day_len, 10, txt='', border=1, fill=True)
        pdf.set_x(x + day_len)
        x += day_len
        for t in times:
            pdf.cell(cell_len, 10, txt=t, border=1, fill=True, align="C")
            pdf.set_x(x + cell_len)
            x += cell_len

        pdf.set_xy(10, y)

        pdf.set_fill_color(84, 84, 84)
        for subs in tbl:
            x = 10
            pdf.set_text_color(230, 230, 230)
            pdf.multi_cell(day_len, cell_height, txt=subs[0], border=1, fill=True, align="C")
            pdf.set_xy(x + day_len, y)
            x += day_len
            pdf.set_text_color(15, 15, 15)
            for i in range(1, len(subs)):
                if subs[i] != '':
                    pdf.multi_cell(cell_len, cell_height // 2, txt=subs[i], border=1, align="C")
                else:
                    pdf.multi_cell(cell_len, cell_height, txt=subs[i], border=1, align="C")
                pdf.set_xy(x + cell_len, y)
                x += cell_len
            y += cell_height
            pdf.set_xy(10, y)

        pdf.output(f"{path}/{sec}.pdf")

    # Creates routine as excel file
    def routine_excel(dic, times, sec, path):
        border = Border(
            left=Side(style='medium'),
            right=Side(style='medium'),
            top=Side(style='medium'),
            bottom=Side(style='medium')
        )
        workbook = wrk()
        rtn = workbook.active
        rtn.title = sec
        lst = [chr(i) for i in range(66, 66 + len(times))]

        for cr in lst:
            rtn[cr + "1"].border = border
        clm = len(times) - 1
        row = 2 * len(dic) + 3
        for i in range(1, row):
            rtn["A" + str(i)].border = border

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
