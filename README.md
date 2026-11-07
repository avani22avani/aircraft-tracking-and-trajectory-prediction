# ADS-B Tracker

A Flask-based aircraft tracking dashboard for PingStation ADS-B data. It can replay a bundled sample capture for a demo or receive a live raw data stream from a PingStation device. The dashboard displays aircraft positions and details on a map and records trajectory history locally.

## Features

- Simulation mode for trying the dashboard without a receiver
- Live PingStation TCP input
- Aircraft map, callsigns, flight details, and observed tracks
- Optional predicted tracks based on recently recorded positions
- Local SQLite trajectory storage and CSV/Excel exports
- Aircraft type and photo lookups when internet access is available

## Requirements

- Python 3.10 or newer
- Windows, macOS, or Linux
- A PingStation device is only needed for live mode

## Run locally

From the project directory, create and activate a virtual environment, install dependencies, and start the app.

### Windows PowerShell

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

### macOS or Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

Open [http://localhost:5000](http://localhost:5000) in a browser. The app uses simulation mode by default and replays the sample capture in `pingstation_data_log/`.

## Configuration

Edit `ADSBTracker.cfg` and restart the app to change its mode. Set `simulation_mode = true` to replay the bundled capture. For live input, set `simulation_mode = false`, then configure the PingStation to send its raw stream to the computer running this app on TCP port `30003`. The default listener binds to `0.0.0.0`; make sure your network and firewall allow the device to connect.

Environment variables can override settings. For example, `SIMULATION_MODE` overrides the mode in the config file. The data source also accepts `PINGSTATION_LISTEN_HOST`, `PINGSTATION_LISTEN_PORT`, and `SIMULATION_SOURCE_FILE`.

## Local data

Runtime data is written under `logs/` by default, including the SQLite trajectory database, session captures, and exports. These files are generated while the app runs. Do not commit private captures or data you do not have permission to share.

## Notes

- Map tiles come from OpenStreetMap and require internet access. Aircraft tracking can continue when map tiles are unavailable.
- Aircraft type, registration, route, and photo details depend on external lookup services and may be unavailable or incomplete.
- The app listens locally on port `5000` by default.
