from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import pandas as pd
import io
import base64
import contextlib
import pandas as pd
import matplotlib.pyplot as plt
from statsforecast import StatsForecast
from statsforecast.models import AutoARIMA
from statsforecast.utils import AirPassengersDF


app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # For development; specify exact origins in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class ScriptData(BaseModel):
    script: str  # The Python script to execute

@app.post("/execute-script")
def execute_script(script_data: ScriptData):
    try:
        stdout = io.StringIO()
        with contextlib.redirect_stdout(stdout):
            exec_namespace = {
                "plt": plt,
                "pd": pd,
                "StatsForecast": StatsForecast,
                "AutoARIMA": AutoARIMA,
                "AirPassengersDF": AirPassengersDF,
            }
            allowed_builtins = {
                'len': len,
                'range': range,
                'print': print,
                'min': min,
                'max': max,
                'list': list,
                'dict': dict,
                'abs': abs,
                'sum': sum,
                'enumerate': enumerate,
                'zip': zip,
                'int': int,
                'float': float,
                'str': str,
                'type': type,
                'isinstance': isinstance,
            }
            exec(
                script_data.script,
                {"__builtins__": allowed_builtins},
                exec_namespace
            )
        if not plt.get_fignums():
            raise HTTPException(status_code=400, detail="No plot was generated in the script.")
        img_bytes = io.BytesIO()
        plt.savefig(img_bytes, format='png')
        plt.close()
        img_bytes.seek(0)
        img_base64 = base64.b64encode(img_bytes.read()).decode('utf-8')
        output = stdout.getvalue()
        return {"plot": img_base64, "output": output}
    except Exception as e:
        output = stdout.getvalue()
        error_message = str(e)
        print("Script execution error:", error_message)
        raise HTTPException(status_code=500, detail=error_message)
