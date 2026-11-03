"""Launch the local demo: generate synthetic data anchored to today, then start the server.

    python demo/run_demo.py [port]

Requires Python 3 and Node 18+ (no packages to install). Binds to 127.0.0.1 only.
"""
import os
import subprocess
import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent


def main():
    port = sys.argv[1] if len(sys.argv) > 1 else "8765"
    data_file = HERE / ".demo_data_today.json"
    sys.path.insert(0, str(HERE))
    import generate_demo_data

    import json
    data = generate_demo_data.build(date.today())
    data_file.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    env = dict(os.environ, DEMO_DATA=str(data_file))
    print("Synthetic data generated for", date.today().isoformat())
    subprocess.run(["node", str(HERE / "server.js"), port], env=env, check=False)


if __name__ == "__main__":
    main()
