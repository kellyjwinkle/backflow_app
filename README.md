# United Fire Backflow Report App

## Setup
1. Place **backflow_template.pdf** in this folder (same dir as app.py)
2. Place **jacksonville_template.pdf** in this folder ← NEW (JEA/Jacksonville form)
3. `pip install -r requirements.txt`
4. `streamlit run app.py`

## GitHub Pages / Streamlit Cloud
1. Create GitHub repo, push all files + both template PDFs
2. Go to https://share.streamlit.io → New app → select repo → app.py
3. Deploy — get a URL for phone/browser access

## Forms
The app now shows a **Select Form** toggle at the top:
- **United Fire (Standard)** — original form (backflow_template.pdf)
- **Jacksonville / JEA** — JEA Cross Connection Control form (jacksonville_template.pdf)

Each form has its own session state so switching between forms does not clear data.

## Adding the Jacksonville Template
Rename `Jacksonville-form-original.pdf` → `jacksonville_template.pdf` and place it
in the repo root alongside `backflow_template.pdf`.

## Coordinate Calibration
Coordinates were derived from pixel positions scaled to PDF 612×792 pts.
To adjust: edit `JAX_TEXT_FIELDS`, `JAX_CHECKBOXES`, `UNITED_TEXT_FIELDS`,
or `UNITED_CHECKBOXES` dicts in app.py.
x increases left→right, y increases bottom→top (PDF coordinate system).

## Sticky Fields (United Fire)
Branch, AHJ, Manufacturer, Model, Size, Gauge Mfg, Gauge Serial,
Date Calibrated, Technician, Cert No., Re-Cert Date

## Sticky Fields (Jacksonville)
Next Report keeps: premises info, device info, tester names/certs.
New Job keeps: tester names/certs only.

## PDF Filename
United Fire: `CustomerName - StreetAddress - Location.pdf`
Jacksonville: `JAX PremisesName - ServiceAddress - PhysicalLocation.pdf`
