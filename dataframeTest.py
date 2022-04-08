import numpy as np
import modules.ApiClass as Api
import time
import json
import os

import pandas as pd
import requests as requests
import pathlib
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

edit = pd.read_excel("GGIP + PLC Collection\PLC.xlsx", 0, index_col=False)
edit["Account"].replace(" ", np.nan, inplace=True)
edit.dropna(subset=["Account"], inplace=True)
