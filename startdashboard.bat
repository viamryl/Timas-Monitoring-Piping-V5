@echo off
cd \d "D:\Monitoring Piping Fix\"
start /b streamlit run dashboard.py --server.headless true
exit