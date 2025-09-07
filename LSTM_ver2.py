import numpy as np
import quantmod as qd
import pandas as pd

import torch
import torch.nn as nn
import torch.optim as optim
from torchvision import datasets, transforms
from torch.utils.data import DataLoader
from torch.autograd import Variable
import matplotlib.pyplot as plt
import pickle

df = qd.get_symbol(ticker="005930.KS", src = 'yahoo', start= "2018-12-31", end = "2021-12-31", to_frame=True)
# High  Low   Open  Close  Volume  Close
input_data = df.values
input_data = input_data[:, :5]

T = 20 # LSTM에서 고려할 타임 시퀀스 (갈수록 반영이 잘되지만 연산속도가 느리고 계산값이 ㅄ이다.)
length = len(input_data) - T
print(len(input_data))
n_class = 5 # 입력 피쳐
n_hidden = 5 # 히든 레이어
output = 1 # 결과
learning_rate = 0.01

train_size = int(len(input_data) * 0.8) # 80%를 학습
train_set = input_data[0:train_size]  # 80%를 학습
test_set = input_data[train_size - T:] # 20%를 테스트
print(train_set.shape)
# 데이터 노말리이제이션
def normalization(data):
    dataN = []
    for n in range(n_class):
        numerator = data[:, n] - np.min(data[:, n], 0)
        deneminator = np.max(data[:, n], 0) - np.min(data[:, n], 0)
        dataN.append(numerator / (deneminator + 1e-7))
    return np.array(dataN).T

# 데이터 디노말리이제이션
def denormalization(data, x, n):
    deneminator = np.max(data[:, n], 0) - np.min(data[:, n], 0)
    x = x * (deneminator + 1e-7)
    return x + np.min(data[:, n], 0)

# 타임 시퀀스 만큼의 데이터 생성
def bulild_dataset(time_series, seq_length, case):
    dataX = []
    dataY = []
    for i in range(0, len(time_series) - seq_length):
        _x = time_series[i:i + seq_length, :]
        _y = time_series[i + seq_length, [case]]
        dataX.append(_x)
        dataY.append(_y)

    return np.array(dataX), np.array(dataY)

def dataset(train_set, test_set, n):
    train_X, train_Y = bulild_dataset(train_set, T, n)
    test_X, test_Y = bulild_dataset(test_set, T, n)
    return train_X, train_Y, test_X, test_Y

def Tesor_transfer(trainX, trainY, testX, testY):
    trainX_tensor = torch.FloatTensor(trainX)
    trainY_tensor = torch.FloatTensor(trainY)

    testX_tensor = torch.FloatTensor(testX)
    testY_tensor = torch.FloatTensor(testY)
    return trainX_tensor, trainY_tensor, testX_tensor, testY_tensor

train_set = normalization(train_set)
test_set = normalization(test_set)
print(train_set.shape)


train_X, train_Y, test_X, test_Y = dataset(train_set, test_set, 3)
print(train_X.shape)
train_X_tensor, train_Y_tensor, test_X_tensor, test_Y_tensor = Tesor_transfer(train_X, train_Y, test_X, test_Y)

class LSTM_NET(torch.nn.Module):
    def __init__(self, input_dim, hidden_dim, output_dim, layers):
        super(LSTM_NET, self).__init__()
        self.rnn = torch.nn.LSTM(input_dim, hidden_dim, num_layers=layers, batch_first=True)
        self.fc = torch.nn.Linear(hidden_dim, output_dim)

    def forward(self, x):
        x, _status = self.rnn(x)

        x = self.fc(x[:, -1])
        return x

net = LSTM_NET(n_class, n_hidden, output, 3)

criterion = torch.nn.SmoothL1Loss()
optimizer = optim.Adam(net.parameters(), lr=learning_rate)

def Train(X, Y, net):
    for i in range(500):
        optimizer.zero_grad()
        outputs = net(X)
        loss = criterion(outputs, Y)
        loss.backward()
        optimizer.step()
        if i % 100 == 0:
            print(i, loss.item())

Train(train_X_tensor, train_Y_tensor, net)


all_set = normalization(input_data)
test_cost = []
origin = []
for i in range(0, len(all_set) - T + 1):
    _x = all_set[i:i + T, :]
    test_cost.append(_x)
    if i == len(all_set) - T:
        pass
    else:
        _y = all_set[i + T, [3]]
        origin.append(_y)

test_tensor = torch.FloatTensor(test_cost)
aa = test_Y_tensor
print(test_tensor.shape)
close_cost = net(test_X_tensor).data.numpy()
print("Predict ?? : ", close_cost[-1])

bb = train_Y_tensor
close_cost_B = net(train_X_tensor).data.numpy()

plt.subplot(1, 2, 2)
plt.plot(aa, label='original')
plt.plot(close_cost, label='predict')
plt.subplot(1, 2, 1)
plt.plot(bb, label='original')
plt.plot(close_cost_B, label='predict')
plt.legend()
plt.show()