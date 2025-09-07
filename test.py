# import pandas as pd
#
# default = "Final_Result"
# load_file_name = "PRD.xlsx"
# data = pd.read_excel(load_file_name, sheet_name=default, na_values='=')
#
# load_df = pd.read_csv('sample.csv')
# print(load_df)


import random as rd
cmp_color = []
for i in range(100):
    cmp_color.append("#"+''.join([rd.choice('0123456789ABCDEF') for j in range(6)]))

print(cmp_color)

# Floor = 3
# CELL = 4
# COLO_list = []
#
# temp_data = {}
#
#
# col_Device = list(data.keys()).index('EndDevice')
# col_port = list(data.keys()).index('EndPort')
#
# usage_device = ['C0', 'M0']
#
# for idx, value in enumerate(data.values):
#     if value[col_Device][-2:] in usage_device:
#         Rack = value[col_port - 1][:5]
#         Device = value[col_Device]
#         Port = value[col_port]
#         print(temp_data)
#         if Rack not in temp_data.keys():
#             temp_data[Rack] = {Device: [Port]}
#         else:
#             if Device not in temp_data[Rack].keys():
#                 temp_data[Rack][Device] = [Port]
#             else:
#                 if Port not in temp_data[Rack][Device]:
#                     temp_data[Rack][Device].append(Port)
#                 else:
#                     print("해당포트에 중복이 있습니다. {} 장비의 {} 포트가 중복입니다.".format(Device, Port))
#
# print(temp_data)
#
#
#
# frame = pd.DataFrame(temp_data)
# print(frame)
# frame.to_csv('sample.csv')