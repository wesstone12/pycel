# Import necessary classes from statsforecast
from statsforecast import StatsForecast
from statsforecast.models import AutoARIMA
from statsforecast.utils import AirPassengersDF
from utilsforecast.plotting import plot_series
import pandas as pd
import os
    # Set the environment variable
os.environ['NIXTLA_ID_AS_COL'] = '1'

# AirPassengersDF is provided in the execution environment
df = AirPassengersDF

# Initialize the StatsForecast object
sf = StatsForecast(
    models=[AutoARIMA(season_length=12)],
    freq='ME'
)

# Fit the model
sf.fit(df)

# Predict future values
forecast_df = sf.predict(h=12, level=[95])

print(forecast_df)