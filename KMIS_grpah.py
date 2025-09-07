import os
import pandas as pd
import matplotlib.pyplot as plt

if os.name == 'posix':
    plt.rc("font", family="AppleGothic")
else:
    plt.rc("font", family="Malgun Gothic")

# color = []
# for i in range(100):
#     color.append("#"+''.join([rd.choice('0123456789ABCDEF') for j in range(6)]))
# load_color =

load_color = ['C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', '#8C2FAB', '#CC45C7', '#EB9091', '#5B5D07', '#58B19C', '#EBA1DF', '#95B8B1', '#4D1981', '#893EAB', '#E816B8', '#D76D06', '#9374AB', '#8CDA26', '#E387EF', '#F3CBFC', '#7918F5', '#710DB5', '#982DC4', '#2E26C6', '#1BCD33', '#9866D2', '#1F9952', '#E7954C', '#6605F1', '#A0C2DC', '#BE862C', '#7EF8CF', '#B6B302', '#9F7F6A', '#F17C50', '#277BB3', '#E100D4', '#F89245', '#CD951E', '#5D0C95', '#A59E33', '#1E0B0E', '#6948EA', '#6A9560', '#C52723', '#DD8A55', '#329FF8', '#750B47', '#8D1CD3', '#4971C5', '#A5A5C9', '#9CD755', '#41EC1A', '#C89CAC', '#48B91E', '#1F4973', '#058395', '#D40056', '#6B61F9', '#DEA6D4', '#8F47E7', '#681D80', '#7E846F', '#E76C75', '#ACB59E', '#13F78A', '#971DAC', '#547578', '#30EB6F', '#347ABC', '#D40204', '#F33F06', '#608773', '#49DA31', '#69007B', '#644636', '#0D6FB9', '#E2BB0C', '#24A40A', '#4E99B4', '#9DD948', '#F58D0D', '#C4190E', '#FE951D', '#F19E55', '#49BCB5', '#B9FC72', '#C1FD05', '#12067E', '#F36659', '#3BFE54', '#B99301', '#E0F5D9', '#817E53', '#202897']
load_df = pd.read_excel('KMIS_data.xls')

col_com = list(load_df.keys()).index('법인')
col_typ = list(load_df.keys()).index('품목')
col_pro = list(load_df.keys()).index('품종')
col_day = list(load_df.keys()).index('일자')
col_kg = list(load_df.keys()).index('물량(kg)')
col_cost = list(load_df.keys()).index('금액(원)')

type_set = set() # 품목 set
product_set = set() # 품종 set
day_set = set()

'''
법인 : {(품목, 품종) : {날짜 : [물량 :, 금액:]}}
'''
df = {}

for value in load_df.values:
    type_set.add(value[col_typ])
    product_set.add(value[col_pro])
    day_set.add(value[col_day])
    if value[col_com] not in df.keys():
        df[value[col_com]] = {(value[col_typ], value[col_pro]): {value[col_day]: [value[col_kg], value[col_cost]]}}
    else:
        if (value[col_typ], value[col_pro]) not in df[value[col_com]].keys():
            df[value[col_com]][(value[col_typ], value[col_pro])] = {value[col_day]: [value[col_kg], value[col_cost]]}
        else:
            if value[col_day] not in df[value[col_com]][(value[col_typ], value[col_pro])].keys():
                df[value[col_com]][(value[col_typ], value[col_pro])][value[col_day]] = [value[col_kg], value[col_cost]]

company_list = list(df.keys())
type_list = list(type_set)
product_list = list(product_set)
day_dict, max_date = {}, len(day_set)
for idx, day in enumerate(sorted(day_set)):
    day_dict[day] = idx

# # random
# rand_int1 = rd.randint(0, len(type_list))
# rand_int2 = rd.randint(0, len(product_list))
# print(rand_int1, rand_int2)
# kind = (type_list[rand_int1], product_list[rand_int2]) # (품목, 품종)

# if you want manual
kind = ('감귤', '천혜향')

print(kind)

# company_color
cmp_color = {}
for i, name in enumerate(company_list):
    cmp_color[name] = load_color[i]

# ------------------------------ line graph ------------------------------------
graph_value = []
for company in df.values():
    temp = [0]*max_date
    if kind in company.keys():
        for key, cost in company[kind].items():
            temp[day_dict[key]] = round(cost[1] / cost[0], 2)
    graph_value.append(temp)

date_str = list(day_dict.keys())
date_str.reverse()

# x_label setting
date_x, day_term = [], 5
for idx, date in enumerate(date_str):
    if idx % day_term == 0:
        date_x.append(str(date))
    else:
        date_x.append('')
date_x.reverse()

plt.figure(figsize=(12, 8))

x_label = range(max_date)
for id in range(len(graph_value)):
    plt.scatter(x_label, graph_value[id], color=list(cmp_color.values())[id], label=company_list[id])

plt.legend()
plt.xticks(x_label, date_x)

# ------------------------ bar graph -------------------------------------------

bar_graph_value = {} # X-label
graph_label = ["Max", "Min", "Avg"] # Y-label
max_temp = 0
for cmp_code, ava_prod in df.items():
    if kind in ava_prod.keys():
        temp = []
        for cost in ava_prod[kind].values():
            temp.append(round(cost[1] / cost[0], 2))
        bar_graph_value[cmp_code] = [max(temp), min(temp), round(sum(temp) / len(temp), 2)]
        if max_temp < max(temp):
            max_temp = max(temp)

def present_value(ax, barh):
    for data in barh:
        width = data.get_width()
        locx = width*1.001
        locy = data.get_y()+data.get_height()*0.5
        ax.text(locx, locy, '%.2f' % width, ha='left', va='center')

fig, ax = plt.subplots(1, 1, figsize=(10, 8))
fig.suptitle(str(list(day_dict.keys())[0]) + " ~ " + str(list(day_dict.keys())[-1]))

available_cmp = list(bar_graph_value.keys())

for id, bar_data in enumerate(list(bar_graph_value.keys())):

    height = [x + 0.15 * (id-0.5*(len(available_cmp)-1)) for x in range(len(graph_label))]
    barh = ax.barh(height, bar_graph_value[bar_data], label= bar_data, color=cmp_color[bar_data], height=0.1)
    present_value(ax, barh)
    ax.set_xlim(0, max_temp + 1000)

ax.xaxis.set_tick_params(labelsize=10)
ax.set_xlabel('Price per 1kg', fontsize=14)

ax.set_yticks(range(len(graph_label)))
ax.set_yticklabels(graph_label, fontsize=15)
ax.set_ylabel('Data type', fontsize=12)

frame = ax.get_position()
ax.set_position([frame.x0, frame.y0, frame.width * 0.9, frame.height])
ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), ncol=1)

ax.set_axisbelow(True)
ax.xaxis.grid(True, color='gray', linestyle='dashed', linewidth=0.5)


plt.show()