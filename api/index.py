import os
import sys

# This allows Vercel to find the streamlit executable
def handler(event, context):
    os.system("streamlit run api/app.py --server.port 8080 --server.address 0.0.0.0")
