import sys
import os
import threading
import time
import webview
from app import app

def start_flask():
    """Starts local Flask backend server silently in background thread."""
    app.run(host="127.0.0.1", port=5000, debug=False)

def main():
    # Start Flask server in background thread
    server_thread = threading.Thread(target=start_flask, daemon=True)
    server_thread.start()
    
    # Small pause to allow server socket initialization
    time.sleep(1.0)
    
    # Create Native GUI Desktop Application Window
    webview.create_window(
        "ખેરાલુ નગરપાલિકા - જન્મ નોંધણી સુધારા હુકમ (Desktop Software)",
        "http://127.0.0.1:5000",
        width=1180,
        height=840,
        resizable=True,
        min_size=(900, 600)
    )
    webview.start()

if __name__ == "__main__":
    main()
