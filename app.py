"""
EdgeFinder Pro Cloud Launcher
Entrypoint wrapper for Streamlit Community Cloud and cloud container environments.
"""
import os
import runpy

script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "edgefinder_dashboard.py")
runpy.run_path(script_path, run_name="__main__")
