import pandas as pd
import OrcFxAPI
import numpy as np



# 打开模型文件
model_file = r'./example.sim'
excel_file = r'./example.xlsx'

model = OrcFxAPI.Model()

model.LoadData(model_file) # load the data from a simulation file
model.LoadSimulation(model_file) # load entire simulation file

model.RunSimulation()

line = model['riser']
n, m = 20.0, 79.0 # 设置周期范围
period = OrcFxAPI.SpecifiedPeriod(n, m)
sample = 0.049998 # 设置采样间隔

times = [] # 保存采样序列
i = 1
while n < m:
    n = 20.0
    n = n+i * sample
    if n <m:
        times.append(n)
    i += 1
 
nums = 105 #点的个数设置需要生成的

# 获取endA
x = line.TimeHistory('X', period=period, objectExtra=OrcFxAPI.oeEndA) # 端点A的接口
x_datas = [np.array(times), x]
                     
for i in range(1,nums+1):
    times.append(20.0 + sample)
    x = line.TimeHistory('X', period=period, objectExtra=OrcFxAPI.oeNodeNum(i)) # 其他点的接口
    y = line.TimeHistory('Y', period=period, objectExtra=OrcFxAPI.oeNodeNum(i))
    x_datas.append(x)
    x_datas.append(y)
print(len(x_datas), len(x_datas[0]))
columns = ['X', 'Y']*nums
# 增加列
columns.insert(0, 'Time')


columns.insert(1, 'X')

x_datas = np.asarray(x_datas)
df = pd.DataFrame(x_datas.T,columns=columns)
df.to_excel(excel_file, index=False)
print("运行成功：excel保存在：", excel_file)

