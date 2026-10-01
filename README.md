# Wetland QC Online v1
Browser version of the working 2019/2021/2024 Wetland QC Tool.

## Use
Open the deployed URL → upload Excel/CSV → upload Maps ZIP → RUN QC → download result Excel.

## Streamlit Community Cloud
1. Create a GitHub repo.
2. Upload all files in this folder.
3. Create a Streamlit app from that repo.
4. Main file: `app.py`.
5. Deploy.

## Private/internal server
Recommended for office-sensitive data.

    python -m venv .venv
    source .venv/bin/activate
    pip install -r requirements.txt
    streamlit run app.py --server.address 0.0.0.0 --server.port 8501

The office PC only needs a browser once hosted.
