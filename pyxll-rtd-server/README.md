# PyXLL RTD Server

A PyXLL package that exposes the repository's WebSocket market data as Excel RTD topics.

This is separate from the Office.js add-in in `addin/`. The Office.js add-in uses the `COG` namespace; this project only appears in Excel when the PyXLL Excel add-in is installed and loaded.

## Setup

PyXLL is a commercial Excel add-in and must be installed separately with a valid license. Install PyXLL into Excel first, then from this directory create an environment and install the package dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e .
```

In the active `pyxll.cfg` used by Excel, add the project source directory and module. Use an absolute path for `pythonpath` if the configuration file is outside this project:

```ini
[PYXLL]
pythonpath = D:\Workspace\Github-Repo\Other Repo\excel-rtd-websocket-addin\pyxll-rtd-server\src
modules = pyxll_rtd_server
```

Install `websocket-client` into the Python environment used by PyXLL, which may differ from this project's `.venv`. The `run-pyxll-rtd-server.bat` file installs the package into the project environment; it cannot change the Python environment embedded/configured by PyXLL.

After reloading PyXLL, verify that `WEBSOCKET_RTD` appears in Excel's PyXLL function wizard before entering the formula.

## Run with the mock server

From this directory, run `run-pyxll-rtd-server.bat` to create or reuse the virtual environment, install the package, install the mock server dependencies, and start the mock WebSocket server. Keep that command window running.

Then use the PyXLL function in Excel:

```excel
=WEBSOCKET_RTD("ACC.NS","LAST")
```

The optional third argument changes the WebSocket endpoint:

```excel
=WEBSOCKET_RTD("ACC.NS","VOLUME","ws://localhost:8080")
```

Supported fields are `LAST`, `OPEN`, `HIGH`, `LOW`, and `VOLUME`.
