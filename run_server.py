import uvicorn
import os
import sys

# Ensure current directory is in Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if __name__ == "__main__":
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "8000"))
    
    print("=" * 60)
    print("  EFFILUME Automated Excel Consolidation Dashboard")
    print(f"  Local Access:   http://localhost:{port}")
    print(f"  Network Access: http://127.0.0.1:{port} (or LAN IP:{port})")
    print("=" * 60)
    
    uvicorn.run("app:app", host=host, port=port, log_level="info")
