# Kevin Moy, 8/8/2020
# Uses IEEE 13-bus system to scale loads and determine state space for RL agent learning
# Spins up COM interface to determine a suitable load configuration

import win32com.client
import pandas as pd
import os
import numpy as np
from generate_state_space import load_states

currentDirectory = os.getcwd()  # Will switch to OpenDSS directory after OpenDSS Object is instantiated!

MAX_NUM_CONFIG = 30
MIN_BUS_VOLT = 0.8
MAX_BUS_VOLT = 1.2


def new_load_config():
    # Generate a new load configuration
    # Instantiate the OpenDSS Object
    try:
        DSSObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
    except:
        print("Unable to start the OpenDSS Engine")
        raise SystemExit

    # Set up the Text, Circuit, and Solution Interfaces
    DSSText = DSSObj.Text
    DSSCircuit = DSSObj.ActiveCircuit
    DSSSolution = DSSCircuit.Solution

    # Load in an example circuit
    DSSText.Command = r"Compile 'C:\Program Files\OpenDSS\IEEETestCases\13Bus\IEEE13Nodeckt.dss'"

    # Disable voltage regulators
    DSSText.Command = "Disable regcontrol.Reg1"
    DSSText.Command = "Disable regcontrol.Reg2"
    DSSText.Command = "Disable regcontrol.Reg3"

    loadNames = np.array(DSSCircuit.Loads.AllNames)
    loadKwdf = pd.DataFrame(loadNames)

    loadKws = load_states(loadNames, DSSCircuit, DSSSolution)

    return loadKws