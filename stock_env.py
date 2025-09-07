import numpy as np
import quantmod as qd
import pandas as pd

import matplotlib.pyplot as plt
import pickle

class STOCK:
    def __init__(self, stock, start_time, end_time, start_money, start_seed, input_size, variance_level, period, risk, reward_period, one_hot):
        # High  Low   Open  Close  Volume  Close
        self.df = qd.get_symbol(ticker=stock, src = "yahoo", start= start_time, end = end_time, to_frame=True)
        print(self.df.index[0])
        self.data = self.df.values
        self.seed_constant = start_money
        self.num_stock_constant = start_seed
        self.seed = 0
        self.num_stock = 0
        self.input_size = input_size
        self.buy_stock = variance_level # 무조건 홀수
        self.end_time = len(self.data)
        self.period = 100 if period == 'default' else period
        self.reward_period = reward_period
        self.time_able = self.end_time - self.period - self.input_size - self.reward_period
        self.one_hot_size = one_hot
        self.count, self.do = 0, 1
        self.center = 1
        self.risk = [(risk ** i) for i in range(self.reward_period)]

    def normalization(self, data):
        numerator = data - np.min(data, 0)
        deneminator = np.max(data, 0) - np.min(data, 0)
        return numerator / (deneminator + 1e-7)

    def reset(self):
        self.seed = self.seed_constant
        self.num_stock = self.num_stock_constant
        self.count = np.random.randint(self.time_able)
        self.do = 1
        state = self.state_make(self.count)
        return state

    def step(self, action):
        num = int(self.buy_stock * (action - self.center))
        present_price = int(self.data[self.input_size + self.count, -1])
        # tommorow_cost = int(self.data[day + self.count + 1, -1])
        present_value = present_price*np.ones(self.reward_period)
        future_values = self.data[(self.input_size + self.count+1):(self.input_size + self.count+self.reward_period+1), -1]
        diff = future_values - present_value

        done = False

        self.num_stock += num
        self.seed -= present_price * num

        # total_reward = float(num*(np.sum(np.array(self.risk) * diff)))
        total_reward = np.sum(np.array(self.risk) * diff)
        if action == 1 and self.num_stock > 0:
            total_reward = float(-1 * total_reward)
        else:
            total_reward = float(num * total_reward)

        self.count += 1
        self.do += 1
        next_state = self.state_make(self.count)

        if self.do > self.period:
            done = True

        return next_state, total_reward, done

    def action_select(self, cost):
        if self.num_stock < self.buy_stock:
            sell = np.zeros(self.center)
            if (self.seed // cost) >= self.buy_stock:
                buy = np.ones(self.center)
            else:
                buy = np.zeros(self.center)
        else:
            sell = np.ones(self.center)
            if (self.seed // cost) >= self.buy_stock:
                buy = np.ones(self.center)
            else:
                buy = np.zeros(self.center)

        able = np.hstack([sell, [1], buy])
        return able

    def action_filter(self, Q, cost):
        able = self.action_select(cost)
        table = Q * able
        table = np.where(table == 0, -1000, table)
        return np.argmax(table[0])

    def state_make(self, count):
        temp = self.normalization(self.data[count:self.input_size+count, :5])
        stock_save = np.zeros(self.one_hot_size)
        stock_save[:int(self.num_stock) + 1] = round(self.seed / self.seed_constant, 2)
        state = np.asarray([temp, stock_save], dtype=object)
        return state

    def test_reset(self):
        self.seed = self.seed_constant
        self.num_stock = self.num_stock_constant
        self.count = self.end_time - self.input_size - self.reward_period
        self.do = 1
        state = self.state_make(self.count)
        return state

    def test_step(self, action):
        num = int(self.buy_stock * (action - self.center))
        present_price = int(self.data[self.input_size + self.count, -1])

        done = False

        self.num_stock += num
        self.seed -= present_price * num

        self.count += 1
        self.do += 1
        next_state = self.state_make(self.count)

        if self.do > self.reward_period:
            done = True

        return next_state, done

    def today_reset(self):
        self.seed = self.seed_constant
        self.num_stock = self.buy_stock
        self.count = self.end_time - self.input_size - 1
        self.do = 1
        state = self.state_make(self.count)
        return state