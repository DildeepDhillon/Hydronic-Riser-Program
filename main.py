import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from schedule import Schedule
from levels import Levels


def pipe_size(flow_gpm, pipe_list):
    if flow_gpm == 0:
        return "N/A"
    for gpm in pipe_list:
        if gpm[0] is not None and flow_gpm <= gpm[0]:
            return gpm[1]


def find_schedule(value, schedule_lists):
    for sched in schedule_lists:
        if value == sched.tag:
            return sched
        else:
            pass
    print("Unable to find a schedule that matched tag provided")


def piping_chart():
    pipe = []
    while True:
        question = input("Do you want to enter your own pipe sizing information (Yes/No): ")
        if question.upper() == "YES":
            break
        elif question.upper() == "NO":
            break
        else:
            print("Invalid input, try again")

    if question.upper() == "YES":
        print("For each one, leave blank if not applicable")
        print("For 1/4\" Piping")
        max_14 = int(input("Max. GPA: "))
        pipe.append(tuple((max_14, "1/4")))

        print("For 1/2\" Piping")
        max_12 = int(input("Max. GPA: "))
        pipe.append(tuple((max_12, "1/2")))

        print("For 3/4\" Piping")
        max_34 = int(input("Max. GPA: "))
        pipe.append(tuple((max_34, "3/4")))

        print("For 1\" Piping")
        max_1 = int(input("Max. GPA: "))
        pipe.append(tuple((max_1, "1")))

        print("For 1 1/4\" Piping")
        max_114 = int(input("Max. GPA: "))
        pipe.append(tuple((max_114, "1 1/4")))

        print("For 1 1/2\" Piping")
        max_112 = int(input("Max. GPA: "))
        pipe.append(tuple((max_112, "1 1/2")))

        print("For 2\" Piping")
        max_2 = int(input("Max. GPA: "))
        pipe.append(tuple((max_2, "2")))

        print("For 2 1/2\" Piping")
        max_212 = int(input("Max. GPA: "))
        pipe.append(tuple((max_212, "2 1/2")))

        print("For 3\" Piping")
        max_3 = int(input("Max. GPA: "))
        pipe.append(tuple((max_3, "3")))

        print("For 4\" Piping")
        max_4 = int(input("Max. GPA: "))
        pipe.append(tuple((max_4, "4")))

        print("For 6\" Piping")
        max_6 = int(input("Max. GPA: "))
        pipe.append(tuple((max_6, "6")))

        print("For 8\" Piping")
        max_8 = int(input("Max. GPA: "))
        pipe.append(tuple((max_8, "8")))

    else:
        pipe.append(tuple((None, "1/4")))
        pipe.append(tuple((None, "1/2")))
        pipe.append(tuple((4, "3/4")))
        pipe.append(tuple((8, "1")))
        pipe.append(tuple((16, "1 1/4")))
        pipe.append(tuple((24, "1 1/2")))
        pipe.append(tuple((48, "2")))
        pipe.append(tuple((75, "2 1/2")))
        pipe.append(tuple((140, "3")))
        pipe.append(tuple((280, "4")))
        pipe.append(tuple((700, "6")))
        pipe.append(tuple((1300, "8")))

    return pipe


def write_excel(vec, cells):
    floor_vec = []
    riser_vec = []
    suite_vec = []
    tag_vec = []
    model_vec = []
    flow_vec = []
    supply_below_vec = []
    supply_top_vec = []
    pipe_bottom_vec = []
    pipe_top_vec = []
    for v in vec:
        floor_vec.append(v.floor)
        riser_vec.append(v.riser)
        suite_vec.append(v.suite)
        tag_vec.append(v.schedule.tag)
        model_vec.append(v.schedule.model)
        flow_vec.append(v.schedule.flow)
        supply_below_vec.append(v.bottomFlow)
        supply_top_vec.append(v.topFlow)
        pipe_bottom_vec.append(v.bottomPipe)
        pipe_top_vec.append(v.topPipe)

    final_list = list(zip(floor_vec, riser_vec, suite_vec, tag_vec, model_vec, flow_vec, supply_below_vec,
                          supply_top_vec, pipe_bottom_vec, pipe_top_vec))

    columns = ['Floor', 'Risers', 'Suite', 'Tag', 'Model', 'Flow', 'Supply Flow Below',
               'Supply FLow Top', 'Supply Size Bottom', 'Supply Size Top']

    df = pd.DataFrame(final_list, columns=columns)

    df.to_excel('Output.xlsx', index=False)

    blue = "000000FF"
    double = Side(border_style="thick", color=blue)
    border = None
    if supply_loc.upper() == "BOTTOM":
        border = Border(top=None, left=None, right=None, bottom=double)
    elif supply_loc.upper() == "TOP":
        border = Border(top=double, left=None, right=None, bottom=None)

    workbook = load_workbook(filename="Output.xlsx")
    sheet = workbook.active

    cell_range = []
    for cell in cells:
        F_num = "F" + str(cell)
        cell_range.append(F_num)
        G_num = "G" + str(cell)
        cell_range.append(G_num)
        H_num = "H" + str(cell)
        cell_range.append(H_num)
        I_num = "I" + str(cell)
        cell_range.append(I_num)
        J_num = "J" + str(cell)
        cell_range.append(J_num)

    for each_cell in cell_range:
        sheet[each_cell].border = border

    workbook.save(filename='Output.xlsx')

    return df


def index_finder(vec, floor_num, riser_spec):
    start = 0
    end = len(vec)
    for v in range(start, end):
        if floor_num == vec[v].floor and riser_spec == vec[v].riser:
            return v


def table_fill(floor_list_1, riser_vec, supply_level_1, supply_loc_1):
    for floors in floor_list_1:
        if str(floors.suite).upper() == "SPECIAL":
            index_floor = index_finder(floor_list_1, floors.floor, floors.riser)
            if floors.floor > supply_level_1:
                # floors.bottomFlow += floors.schedule.flow
                # floors.topFlow += floors.schedule.flow
                floor_list_1[index_floor-1].topFlow = floors.schedule.flow
            elif floors.floor <= supply_level_1:
                # floors.bottomFlow += floors.schedule.flow
                # floors.topFlow += floors.schedule.flow
                floor_list_1[index_floor+1].bottomFlow = floors.schedule.flow
    for r in riser_vec:
        start = supply_level_1
        total = 0
        for each_total in floor_list_1:
            if each_total.riser == r and each_total.floor > start and str(each_total.suite).upper() != "SPECIAL":
                total += each_total.schedule.flow
        print(total)

        for each in floor_list_1:
            if each.riser == r and str(each.suite).upper() != "SPECIAL":
                if each.floor == 1:
                    each.bottomFlow += 0
                    each.bottomPipe = pipe_size(each.bottomFlow, pipe_List)
                    print("floor 1 bottom: ", each.bottomFlow)

                    each.topFlow += each.schedule.flow
                    each.topPipe = pipe_size(each.topFlow, pipe_List)
                    print("floor 1 top: ", each.topFlow)
                elif each.floor == start:
                    if supply_loc_1.upper() == "TOP":
                        each.bottomFlow += floor_list_1[(index_finder(floor_list_1, each.floor,
                                                                      each.riser) - 1)].topFlow
                        each.bottomPipe = pipe_size(each.bottomFlow, pipe_List)
                        print("floor ", each.floor, " bottom: ", each.bottomFlow)

                        each.topFlow += total
                        each.topPipe = pipe_size(each.topFlow, pipe_List)
                        print("floor ", each.floor, " top: ", each.topFlow)

                    elif supply_loc_1.upper() == "BOTTOM":
                        each.bottomFlow += (total + each.schedule.flow)
                        each.bottomPipe = pipe_size(each.bottomFlow, pipe_List)
                        print("floor ", each.floor, " bottom: ", each.bottomFlow)

                        each.topFlow += each.bottomFlow - each.schedule.flow
                        each.topPipe = pipe_size(each.topFlow, pipe_List)
                        print("floor ", each.floor, " top: ", each.topFlow)
                elif each.floor > start:
                    pass
                else:
                    each.bottomFlow += floor_list_1[(index_finder(floor_list_1, each.floor,
                                                                  each.riser) - 1)].topFlow
                    each.bottomPipe = pipe_size(each.bottomFlow, pipe_List)
                    print("floor ", each.floor, " bottom: ", each.bottomFlow)

                    each.topFlow += each.bottomFlow + each.schedule.flow
                    each.topPipe = pipe_size(each.topFlow, pipe_List)
                    print("floor ", each.floor, " top: ", each.topFlow)
            else:
                pass

        for upper in floor_list_1:
            if upper.riser == r and str(upper.suite).upper() != "SPECIAL":
                if upper.floor > start:
                    upper.bottomFlow += floor_list_1[(index_finder(floor_list_1, upper.floor,
                                                                   upper.riser) - 1)].topFlow
                    upper.bottomPipe = pipe_size(upper.bottomFlow, pipe_List)
                    print("floor ", upper.floor, " bottom: ", upper.bottomFlow)

                    upper.topFlow += upper.bottomFlow - upper.schedule.flow
                    upper.topPipe = pipe_size(upper.topFlow, pipe_List)
                    print("floor ", upper.floor, " top: ", upper.topFlow)
            else:
                pass

    return floor_list_1


if __name__ == '__main__':
    """
    while True:
        try:
            number = int(input("How many different tags are there: "))
            break
        except ValueError:
            print("You have to enter numbers, try again")
    """

    schedule_list = []

    dataframe2 = pd.read_excel('input.xlsx', sheet_name='SCHEDULE')
    tupleScheduleList = list(dataframe2.itertuples(index=False, name=None))

    for x in tupleScheduleList:
        tag = None
        model = None
        flow = None
        for y in range(3):
            if y == 0:
                tag = x[y]
            elif y == 1:
                model = x[y]
            elif y == 2:
                flow = x[y]
        schedule_list_object = Schedule(tag, model, flow)
        schedule_list.append(schedule_list_object)

    """
    for x in range(number):
        print("")
        print("Information for Tag " + str(x + 1))
        tag = input("Enter the tag: ")
        model = input("Enter the model number: ")
        flow = int(input("Enter the flow: "))
        list1.append(tuple((tag.upper(), model.upper(), flow)))
    """

    schedule_list.append(Schedule("SPECIAL", None, None))

    pipe_List = piping_chart()

    dataframe1 = pd.read_excel('book1.xlsx', sheet_name='TAKEOFF')

    tupleList = list(dataframe1.itertuples(index=False, name=None))

    floor_list = []
    riser_list = []

    for x in tupleList:
        floor = None
        riser = None
        suite = None
        schedule_object = None
        special = 0
        for y in range(4):
            if y == 0:
                floor = x[y]
            elif y == 1:
                riser = x[y]
                riser_list.append(riser)
            elif y == 2:
                suite = x[y]
                if str(suite).upper() == "SPECIAL":
                    special = 1
            elif y == 3:
                if special == 1:
                    schedule_object = Schedule("SPECIAL", "Supply", int(x[y]))
                else:
                    schedule_object = find_schedule(x[y], schedule_list)
        level = Levels(floor, riser.upper(), suite, schedule_object)
        floor_list.append(level)

    floor_list.sort(key=lambda a: a.floor)
    floor_list.sort(key=lambda a: a.riser)

    final_riser_list = []
    [final_riser_list.append(x) for x in riser_list if x not in final_riser_list]

    while True:
        try:
            supply_level = int(input("What level is the supply run: "))
            break
        except ValueError:
            print("You have to enter numbers, try again")

    while True:
        supply_loc = input("Is it at the Top/Bottom: ")
        if supply_loc.upper() == "TOP":
            break
        elif supply_loc.upper() == "BOTTOM":
            break
        else:
            print("Invalid input, try again")

    final_floor_list = table_fill(floor_list, final_riser_list, supply_level, supply_loc)
    final_floor_list.sort(key=lambda a: a.floor, reverse=True)
    final_floor_list.sort(key=lambda a: a.riser)

    write_excel(final_floor_list, excel_index)
