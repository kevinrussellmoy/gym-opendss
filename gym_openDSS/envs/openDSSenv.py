import gym

from gym import error, spaces, utils

from gym.utils import seeding


class openDSSenv(gym.Env):
    metadata = {'render.modes': ['human']}

    def __init__(self):
        print('Env initialized')

    def step(self):
        print('Step success')

    def reset(self):
        print('env reset')