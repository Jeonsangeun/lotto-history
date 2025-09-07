import numpy as np
import random as rd
import torch
import torch.nn as nn
import torch.nn.functional as F
import torch.optim as optim
import collections
use_cuda = torch.cuda.is_available()
device = torch.device('cuda:0' if use_cuda else "cpu")
buffer_limit = 100000

class ReplayBuffer():
    def __init__(self):
        self.buffer = collections.deque(maxlen=buffer_limit)

    def put(self, transition):
        self.buffer.append(transition)

    def sample(self, n):
        mini_batch = rd.sample(self.buffer, n)
        s_lst, a_lst, r_lst, s_prime_lst, done_mask_lst = [], [], [], [], []

        for transition in mini_batch:
            s, a, r, s_prime, done_mask = transition
            s_lst.append(s)
            a_lst.append([a])
            r_lst.append([r])
            s_prime_lst.append(s_prime)
            done_mask_lst.append([done_mask])

        return s_lst, torch.tensor(a_lst), \
               torch.tensor(r_lst), s_prime_lst, \
               torch.tensor(done_mask_lst)

    def size(self):
        return len(self.buffer)

# class Net(nn.Module):
#     def __init__(self, input_size, output_size):
#         super(Net, self).__init__()
#         self.input_size = input_size
#         self.output_size = output_size
#
#         conv1 = nn.Conv1d(self.input_size, 50, kernel_size=1, stride=1)  # n x 60 x 5
#         # conv2 = nn.Conv1d(60, 30, kernel_size=1, stride=1)  # n x 30 x 5
#         self.conv_module = nn.Sequential(conv1, nn.ReLU()).to(device)  # n x 30 x 5
#         self.relu = nn.ELU(inplace=False).to(device)
#
#         fc1 = nn.Linear(250, 200)
#         fc2 = nn.Linear(200, 150)
#         fc3 = nn.Linear(150, 100)
#         fc4 = nn.Linear(100, self.output_size)
#
#         self.fc = nn.Sequential(fc1, self.relu, fc2, self.relu, fc3, self.relu, fc4).to(device)
#
#         for m in self.modules():
#             if isinstance(m, nn.Linear):
#                 nn.init.kaiming_uniform_(m.weight.data)  # Kaming He Initialization
#                 m.bias.data.fill_(0)  # 편차를 0으로 초기화
#
#     def forward(self, x):
#         x = x.to(device)
#         print(x.shape)
#         cnn1 = self.conv_module(x)
#         cnn2 = cnn1.reshape(-1, 250)
#
#         out = self.fc(cnn2)
#         return out

class Net2(nn.Module):
    def __init__(self, input_size, output_size,  hidden_size):
        super(Net2, self).__init__()
        self.input_size = input_size
        self.output_size = output_size
        self.hidden = hidden_size

        self.relu = nn.ELU(inplace=False).to(device)

        self.fc_h1 = nn.Linear(self.input_size, self.hidden)
        self.fc_h2 = nn.Linear(self.hidden, self.output_size)

        self.fc_l1 = nn.Linear(self.input_size, self.hidden)
        self.fc_l2 = nn.Linear(self.hidden, self.output_size)

        self.fc_o1 = nn.Linear(self.input_size, self.hidden)
        self.fc_o2 = nn.Linear(self.hidden, self.output_size)

        self.fc_c1 = nn.Linear(self.input_size, self.hidden)
        self.fc_c2 = nn.Linear(self.hidden, self.output_size)

        self.fc_v1 = nn.Linear(self.input_size, self.hidden)
        self.fc_v2 = nn.Linear(self.hidden, self.output_size)

        # for m in self.modules():
        #     if isinstance(m, nn.Linear):
        #         nn.init.xavier_uniform_(m.weight.data)  # Kaming He Initialization
        #         # m.bias.data.fill_(0)  # 편차를 0으로 초기화

    def forward(self, x):
        x = x.to(device)
        x = torch.reshape(x, [-1, 5, self.input_size])
        x_High = x[:, 0] * 1
        x_Low = x[:, 1] * 1
        x_Open = x[:, 2] * 1
        x_Close = x[:, 3] * 1
        x_Volume = x[:, 4] * 1

        xout_h = self.fc_h2(F.elu(self.fc_h1(x_High)))
        xout_l = self.fc_l2(F.elu(self.fc_l1(x_Low)))
        xout_o = self.fc_o2(F.elu(self.fc_o1(x_Open)))
        xout_c = self.fc_c2(F.elu(self.fc_c1(x_Close)))
        xout_v = self.fc_v2(F.elu(self.fc_v1(x_Volume)))

        out = xout_h + xout_l + xout_o + xout_c + xout_v
        return out


class Net(nn.Module):
    def __init__(self, input_size, output_size, one_hot):
        super(Net, self).__init__()
        self.input_size = input_size
        self.output_size = output_size
        self.time_hidden = 20
        self.one_hot = one_hot

        conv1 = nn.Conv1d(self.input_size, self.time_hidden, kernel_size=1, stride=1)  # n x 60 x 5
        # conv2 = nn.Conv1d(60, 30, kernel_size=1, stride=1)  # n x 30 x 5
        self.conv_module = nn.Sequential(conv1, nn.ReLU()).to(device)  # n x 30 x 5
        self.relu = nn.ELU(inplace=False).to(device)

        fc1 = nn.Linear(self.one_hot, self.one_hot)
        fc2 = nn.Linear(self.one_hot, 5*self.time_hidden)

        self.fc_num = nn.Sequential(fc1, self.relu, fc2).to(device)

        fc3 = nn.Linear(5*self.time_hidden, 50)
        fc4 = nn.Linear(50, self.output_size)

        self.fc_total = nn.Sequential(fc3, self.relu, fc4).to(device)

        for m in self.modules():
            if isinstance(m, nn.Linear):
                nn.init.kaiming_uniform_(m.weight.data)  # Kaming He Initialization
                m.bias.data.fill_(0)  # 편차를 0으로 초기화

    def forward(self, x):
        stock_data = list(x[:, 0]**1)
        stock_num = list(x[:, 1]**1)
        stock_data = torch.from_numpy(np.asarray(stock_data)).float().to(device)
        stock_num = torch.from_numpy(np.asarray(stock_num)).float().to(device)

        cnn1 = self.conv_module(stock_data)
        cnn1 = cnn1.reshape(-1, 5 * self.time_hidden)
        fcn1 = self.fc_num(stock_num)

        out = cnn1 + fcn1
        out = self.fc_total(out)

        return out