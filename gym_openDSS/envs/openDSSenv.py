import gym
import win32com.client
from gym import spaces
import numpy as np
import logging

from gym_openDSS.envs.bus13_state_reward import *
from gym_openDSS.envs.find_load_config import *
from gym.utils import seeding

# Upper and lower bounds of voltage zones:
ZONE2_UB = 1.10
ZONE1_UB = 1.05
ZONE1_LB = 0.95
ZONE2_LB = 0.90

# Penalties for each zone:
# TODO: Tune these hyperparameters
ZONE1_PENALTY = -200
ZONE2_PENALTY = -400

logging.basicConfig(format='%(asctime)s %(levelname)s: %(message)s', level=logging.WARNING)


class openDSSenv(gym.Env):
    metadata = {'render.modes': ['human']}

    def __init__(self):
        print("initializing 13-bus env")
        self.DSSObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
        self.DSSText = self.DSSObj.Text
        self.DSSCircuit = self.DSSObj.ActiveCircuit
        self.DSSSolution = self.DSSCircuit.Solution

        self.DSSText.Command = r"Compile 'C:\Program Files\OpenDSS\IEEETestCases\13Bus\IEEE13Nodeckt.dss'"

        print("disabling voltage regulators")
        self.DSSText.Command = "Disable regcontrol.Reg1"
        self.DSSText.Command = "Disable regcontrol.Reg2"
        self.DSSText.Command = "Disable regcontrol.Reg3"

        # Initially disable both capacitor banks and set both to 1500 KVAR rating
        capNames = self.DSSCircuit.Capacitors.AllNames
        for cap in capNames:
            self.DSSCircuit.SetActiveElement("Capacitor." + cap)
            self.DSSCircuit.ActiveDSSElement.Properties("kVAR").Val = 1500
        self.DSSCircuit.Capacitors.Name = "Cap1"
        self.DSSCircuit.Capacitors.States = (0,)
        self.DSSCircuit.Capacitors.Name = "Cap2"
        self.DSSCircuit.Capacitors.States = (0,)

        self.loadNames = self.DSSCircuit.Loads.AllNames
        self.VoltageMag = self.DSSCircuit.AllBusVmagPu

        # Set up action and observation space variables
        n_actions = 4
        self.action_space = spaces.Discrete(n_actions)
        self.observation_space = spaces.Box(low=0, high=2, shape=(len(self.DSSCircuit.AllBusVmagPu),), dtype=np.float32)

        print('Env initialized')

    def step(self, action):
        # Execute action based on control
        # Expect action in range [0 3] for capacitor control
        if action == 0:
            # Both capacitors off:
            self.DSSCircuit.Capacitors.Name = "Cap1"
            self.DSSCircuit.Capacitors.States = (0,)
            self.DSSCircuit.Capacitors.Name = "Cap2"
            self.DSSCircuit.Capacitors.States = (0,)
        elif action == 1:
            # Capacitor 1 on, Capacitor 2 off:
            self.DSSCircuit.Capacitors.Name = "Cap1"
            self.DSSCircuit.Capacitors.States = (1,)
            self.DSSCircuit.Capacitors.Name = "Cap2"
            self.DSSCircuit.Capacitors.States = (0,)
        elif action == 2:
            # Capacitor 1 off, Capacitor 2 on:
            self.DSSCircuit.Capacitors.Name = "Cap1"
            self.DSSCircuit.Capacitors.States = (0,)
            self.DSSCircuit.Capacitors.Name = "Cap2"
            self.DSSCircuit.Capacitors.States = (1,)
        elif action == 3:
            # Both capacitors on:
            self.DSSCircuit.Capacitors.Name = "Cap1"
            self.DSSCircuit.Capacitors.States = (1,)
            self.DSSCircuit.Capacitors.Name = "Cap2"
            self.DSSCircuit.Capacitors.States = (1,)
        else:
            inval_action_message = "Invalid action " + str(action) + ", action in range [0 3] expected"
            print(inval_action_message)
            logging.warning(inval_action_message)

        self.DSSSolution.Solve()  # Solve Circuit
        observation = get_state(self.DSSCircuit)
        reward = quad_reward(observation)
        done = True
        info = {}
        logging.info('Step success')

        return observation, reward, done, info

    def reset(self):
        logging.info('resetting environment...')
        logging.info("Set new loads")
        # Set new loads
        # TODO: Handle loads as an input
        loadKws = new_load_config()
        logging.info("New loads obtained")
        for loadnum in range(np.size(self.loadNames)):
            self.DSSCircuit.SetActiveElement("Load." + self.loadNames[loadnum])
            # Set load with new loadKws
            self.DSSCircuit.ActiveDSSElement.Properties("kW").Val = loadKws[loadnum]
        # Get state observations from initial default load configuration
        self.DSSSolution.solve()
        logging.info("reset complete\n")
        obs = get_state(self.DSSCircuit)
        return obs

    def render(self, mode='human', close=False):
        pass
