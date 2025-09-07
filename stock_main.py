import numpy as np
import quantmod as qd
import pandas as pd
import torch
import torch.nn as nn
import torch.nn.functional as F
import torch.optim as optim
from torchvision import datasets, transforms
from torch.utils.data import DataLoader
from torch.autograd import Variable
import matplotlib.pyplot as plt
import pickle

import stock_env as stock
import stock_net
use_cuda = torch.cuda.is_available()
device = torch.device('cuda:0' if use_cuda else "cpu")

learning_rate = 0.001
gamma = 0.99

batch_size = 32

stock_value = "005930.KS"
start_day = "2016-12-01" #25300
end_day = "2021-12-01" #86000
Now_money = 1000000
risk = 0.75 # ~ 1.0
Now_stock = 0
sight_day = 200
Period = 100 # 'default' # = 100
one_hot = 100
variance_level = 5 # 한 액션이 몇 주인지

validation_len = 100
# stock, start_time, end_time, start_money, start_seed, target_cost, input_size, output_size
# 주식종류, 시작일, 종료일, 보유금액, 보유주, 목표가, 과거에 볼데이터수, 매수 매도 범위(2n + 1)
env = stock.STOCK(stock_value, start_day, end_day, Now_money, Now_stock,
                  sight_day, variance_level, Period, risk, validation_len, one_hot)

print(env.end_time)
y_layer = []
input_size = sight_day #(현가, 현주, 주가, 과거데이터)
output_size = 3

def Train(Q, Q_target, memory, optimizer):
    for i in range(20):
        state, action, reward, next_state, done = memory.sample(batch_size)

        # state = state.cuda(device)
        action = action.cuda(device)
        reward = reward.cuda(device)
        # next_state = next_state.cuda(device)
        done = done.cuda(device)

        s = np.reshape(state, [batch_size, -1])
        s_n = np.reshape(next_state, [batch_size, -1])

        # # # DDQN
        # Q_out = Q(s)
        # Q_value = Q_out.gather(1, action)
        # Q_argmax_value = Q_out.max(1)[1].unsqueeze(1)
        # Q_prime = Q_target(s_n)
        # Q_prime = Q_prime.gather(1, Q_argmax_value)

        # # DQN
        Q_out = Q(s)
        Q_value = Q_out.gather(1, action)
        Q_prime = Q_target(s_n).max(1)[0].unsqueeze(1)

        target = reward + gamma * Q_prime * done
        loss = F.smooth_l1_loss(Q_value, target)

        optimizer.zero_grad()
        loss.backward()
        optimizer.step()

def main():
    main_Net = stock_net.Net(input_size, output_size, one_hot).to(device)
    target_Net = stock_net.Net(input_size, output_size, one_hot).to(device)

    target_Net.load_state_dict(main_Net.state_dict())
    target_Net.eval()

    memory = stock_net.ReplayBuffer()
    optimizer = optim.Adam(main_Net.parameters(), lr=learning_rate)
    interval = 10
    max_episode = 1000
    cost = 0
    for episode in range(max_episode):
        state = env.reset()
        done = False
        e = max((1. / ((episode // 200) + 1)), 0.1)
        while not done:
            # print("-----------------Step------------------ :", env.df.index[sight_day+env.count])
            # print("가치총액:", env.seed + env.num_stock * int(env.data[sight_day+env.count, -1]))
            # print("현재 돈", env.seed)
            # print("현재 주식", env.num_stock)
            # print("현재 가격", int(env.data[sight_day+env.count, -1]))
            if np.random.rand(1) < e:
                able = env.action_select(int(env.data[sight_day+env.count - 1, -1]))
                aa = np.where(able == 1)[0]
                # print(aa)
                action = np.random.choice(aa, 1)[0]
            else:
                s = np.reshape(state, [1, -1])
                with torch.no_grad():
                    aa = main_Net(s).cpu().detach().numpy()
                action = env.action_filter(aa, env.data[sight_day+env.count - 1, -1])

            next_state, reward, done = env.step(action)
            reward = reward / (risk * 100000)
            # print("시도", action - 1)
            # print("보상", reward)

            done_mask = 0.0 if done else 1.0
            memory.put((state, action, reward, next_state, done_mask))
            state = next_state

        cost += (env.seed + env.num_stock * env.data[sight_day+env.count - 1, -1])
        if episode % interval == (interval - 1):
            rate = np.round(100*((cost / interval) - Now_money) / Now_money)
            y_layer.append(rate)
            print("Episode: {} cost: {}%".format(episode, rate, 3))
            Train(main_Net, target_Net, memory, optimizer)
            target_Net.load_state_dict(main_Net.state_dict())
            target_Net.eval()
            cost = 0

    torch.save(main_Net.state_dict(), "check_point.pth")

    price_value = []
    action_value = []

    state = env.test_reset()
    done = False
    while not done:
        print("-----------------Step------------------ :", env.df.index[sight_day+env.count])
        print("가치총액:", env.seed + env.num_stock * int(env.data[sight_day+env.count - 1, -1]))
        print("현재 돈", env.seed)
        print("현재 주식", env.num_stock)
        print("현재 가격", int(env.data[sight_day+env.count, -1]))

        price_value.append(env.seed + env.num_stock * int(env.data[sight_day+env.count - 1, -1]))

        s = np.reshape(state, [1, -1])
        with torch.no_grad():
            aa = main_Net(s).cpu().detach().numpy()

        action = env.action_filter(aa, env.data[sight_day + env.count - 1, -1])
        action_value.append(action)

        next_state, done = env.test_step(action)
        print("시도", action - 1)

        state = next_state
    cost += (env.seed + env.num_stock * env.data[sight_day + env.count - 1, -1])
    final_result = np.round(100 * (cost - Now_money) / Now_money)
    print(final_result, "%")

    state = env.today_reset()
    print("그럼 오늘은 사야할까?", env.df.index[sight_day+env.count])
    print(env.data[-1, :5])
    print(env.data[sight_day + env.count - 1, -1])
    s = np.reshape(state, [1, -1])
    with torch.no_grad():
        aa = main_Net(s).cpu().detach().numpy()
    print(aa)
    action = env.action_filter(aa, env.data[sight_day + env.count - 1, -1])
    result = action - 1
    if result > 0:
        print("매수 가즈아아아아")
        print("매주할 주 : ", abs(result))
    elif result < 0:
        print("매도 가즈아아아아")
        print("매도할 주 : ", abs(result))
    else:
        print("중입기어 씨게 박자 후우...")

    plt.subplot(311)
    plt.plot(y_layer)
    plt.subplot(312)
    plt.plot(price_value)
    plt.subplot(313)
    plt.plot(action_value)
    plt.show()

if __name__ == '__main__':
    main()