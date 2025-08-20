import requests
import zipfile
import io
import pandas as pd
from datetime import datetime, timedelta

def get_soil_temp_full():
    url = "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/bodemtemps/bodemtemps_260.zip"

    resp = requests.get(url)
    resp.raise_for_status()

    with zipfile.ZipFile(io.BytesIO(resp.content)) as z:
        txt_files = [f for f in z.namelist() if f.endswith(".txt")]
        if not txt_files:
            raise FileNotFoundError("Geen .txt bestand in zip gevonden.")

        with z.open(txt_files[0]) as f:
            lines = f.read().decode("utf-8").splitlines()

    header_idx = next(i for i, line in enumerate(lines) if line.startswith("# STN"))

    df = pd.read_csv(
        io.StringIO("\n".join(lines[header_idx+1:])),
        header=None,
        names=["STN","YYYYMMDD","HH","TB1","TB2","TB3","TB4","TB5","TNB1","TNB2","TXB1","TXB2","NaN"],
        na_values=["     "],
        skip_blank_lines=True
    )

    # Fix datetime
    def parse_datetime(row):
        date_str = str(int(row["YYYYMMDD"]))
        hour = int(row["HH"])
        dt = datetime.strptime(date_str, "%Y%m%d")
        if hour == 24:
            dt += timedelta(days=1)
            hour = 0
        return dt + timedelta(hours=hour)

    df["datetime"] = df.apply(parse_datetime, axis=1)

    # Zet TB-waarden van tienden °C naar °C
    for col in ["TB1","TB2","TB3","TB4","TB5","TNB1","TNB2","TXB1","TXB2"]:
        df[col] = df[col] / 10.0

    return df
