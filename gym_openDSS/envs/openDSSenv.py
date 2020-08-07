import gym
import win32com.client
from gym import spaces
import numpy as np

from gym_openDSS.envs.find_load_config import new_load_config

# Upper and lower bounds of voltage zones:
ZONE2_UB = 1.10
ZONE1_UB = 1.05
ZONE1_LB = 0.95
ZONE2_LB = 0.90

# Penalties for each zone:
# TODO: Tune these hyperparameters
ZONE1_PENALTY = -200
ZONE2_PENALTY = -400

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

        self.loadNames = np.array(self.DSSCircuit.Loads.AllNames)

        n_actions = 4
        self.action_space = spaces.Discrete(n_actions)
        self.observation_space = spaces.Box(low=0, high=2, shape=(len(self.DSSCircuit.AllBusVmagPu), 1), dtype=np.float32)

        print('Env initialized')

    def step(self, action):
        # Execute action based on control
        # Expect action in range [0 3] for capacitor control
        # TODO: put in its own method? idk -km
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
            raise ValueError("Received invalid action={} which is not part of the action space".format(action))

        # Solve new circuit with new capacitor states
        self.DSSSolution.solve()

        # Get state observations
        obs = np.array(self.DSSCircuit.AllBusVmagPu)

        # Calculate reward from states
        # Number of buses in voltage zone 1
        num_zone1 = np.size(np.nonzero(np.logical_and(obs >= ZONE1_UB, obs < ZONE2_UB))) \
                    + np.size(np.nonzero(np.logical_and(obs <= ZONE1_LB, obs > ZONE2_LB)))

        # Number of buses in voltage zone 2
        num_zone2 = np.size(np.nonzero(obs >= ZONE2_UB)) \
                    + np.size(np.nonzero(obs <= ZONE2_LB))

        reward = num_zone1 * ZONE1_PENALTY + num_zone2 * ZONE2_PENALTY

        print('Step success')

        return obs, reward

    def reset(self):
        print('env reset')

        print("Set new loads")
        # Set new loads
        # TODO: Handle loads as an input
        loadKws = new_load_config()
        print("New loads obtained")
        for loadnum in range(np.size(self.loadNames)):
            self.DSSCircuit.SetActiveElement("Load." + self.loadNames[loadnum])
            # Set load with new loadKws
            self.DSSCircuit.ActiveDSSElement.Properties("kW").Val = loadKws[loadnum]
        # Get state observations from initial default load configuration
        self.DSSSolution.solve()
        obs = np.array(self.DSSCircuit.AllBusVmagPu)
        return obs
