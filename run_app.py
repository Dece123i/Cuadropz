import os
import sys
import streamlit.web.bootstrap

if __name__ == '__main__':
    # Locate the Streamlit app.py script inside PyInstaller's temporary folder (_MEIPASS)
    base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    app_path = os.path.join(base_dir, 'app.py')
    
    # Force configuration programmatically to prevent Streamlit from reverting to development defaults (port 3000)
    import streamlit.config
    streamlit.config.set_option("global.developmentMode", False)
    streamlit.config.set_option("browser.gatherUsageStats", False)
    streamlit.config.set_option("server.port", 8501)
    streamlit.config.set_option("server.headless", True)
    
    # Configure Streamlit execution arguments
    # Disable telemetry and set to run on default port 8501
    sys.argv = [
        "streamlit", 
        "run", 
        app_path, 
        "--browser.gatherUsageStats=false", 
        "--server.port=8501"
    ]
    
    flag_options = {
        "global.developmentMode": False,
        "browser.gatherUsageStats": False,
        "server.port": 8501,
        "server.headless": True,
    }
    
    # Run Streamlit bootstrap
    streamlit.web.bootstrap.run(app_path, True, [], flag_options=flag_options)

