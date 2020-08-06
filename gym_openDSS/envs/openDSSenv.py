import gym
import win32com.client
from gym import error, spaces, utils
import numpy as np

from gym.utils import seeding


class openDSSenv(gym.Env):
    metadata = {'render.modes': ['human']}

    def __init__(self):
        self.DSSObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
        self.DSSText = self.DSSObj.Text
        self.DSSCircuit = self.DSSObj.ActiveCircuit
        self.DSSSolution = self.DSSCircuit.Solution

        self.DSSText.Command = r"Compile 'C:\Program Files\OpenDSS\IEEETestCases\13Bus\IEEE13Nodeckt.dss'"

        self.DSSText.Command = "Disable regcontrol.Reg1"
        self.DSSText.Command = "Disable regcontrol.Reg2"
        self.DSSText.Command = "Disable regcontrol.Reg3"

        self.loadNames = self.DSSCircuit.Loads.AllNames
        self.VoltageMag = self.DSSCircuit.AllBusVmagPu

        n_actions = 2
        self.action_space = spaces.Discrete(n_actions)
        self.observation_space = spaces.Box(low=0, high=2, shape=(len(self.VoltageMag), 1), dtype=np.float32)

        print('Env initialized')

    def step(self):
        print('Step success')

    def reset(self):
        print('env reset')