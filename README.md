# 🏦 CVC Analytics Dashboard

A Streamlit web app for searching and analysing **Corporate Venture Capital (CVC)** company information.

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://your-app-url.streamlit.app)

---

## Features

- 🔍 **Full-text search** across company names, descriptions, and sectors
- 🎛️ **Sidebar filters** — Sector, Investment Stage, City, GP/LP activity
- 📊 **Analytics charts** — Sector distribution, Stage breakdown, Geographic map, GP/LP bars
- 📋 **Data table** with CSV export
- 🌙 **Dark glassmorphism UI** with animated hover effects

## Data

The app uses `CVC_Company_Info.xlsx` (404 CVC companies) with the following fields:

| Column | Description |
|--------|-------------|
| Company Name | CVC entity name |
| Business Overview | Description of the CVC and its parent |
| Sector | Investment focus areas |
| Stage | Investment stage preference |
| Location | Headquarters city/country |
| Invest in GPs | Whether the CVC invests in GP funds |
| LP in Secondary Funds | Whether the CVC is an LP in secondary funds |

## Run Locally

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/cvc-analytics.git
cd cvc-analytics

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch the app
streamlit run app.py
```

Then open [http://localhost:8501](http://localhost:8501) in your browser.

## Deploy on Streamlit Community Cloud (Free)

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **"New app"** → select your repo → set `app.py` as the main file
4. Click **Deploy** — your app gets a public URL in ~2 minutes

## Tech Stack

- [Streamlit](https://streamlit.io/) — UI framework
- [Pandas](https://pandas.pydata.org/) — Data manipulation
- [Plotly](https://plotly.com/python/) — Interactive charts
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel file reading
