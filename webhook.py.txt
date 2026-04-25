"""
webhook.py — SiteInsight
Receives trigger from Make.com and runs pipeline
"""
import os
import subprocess
import requests
from flask import Flask, request, jsonify

app = Flask(__name__)
SECRET = os.environ.get("WEBHOOK_SECRET", "siteinsight2026")

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "SiteInsight"}), 200

@app.route("/run-pipeline", methods=["POST"])
def run_pipeline():
    token = request.headers.get("X-Secret-Token", "")
    if token != SECRET:
        return jsonify({"error": "Unauthorized"}), 401

    data = request.get_json()
    if not data:
        return jsonify({"error": "No data received"}), 400

    file_name = data.get("file_name", "WPR.xlsx")
    file_url  = data.get("file_url")

    if not file_url:
        return jsonify({"error": "No file_url provided"}), 400

    try:
        response = requests.get(file_url, timeout=30)
        response.raise_for_status()

        os.makedirs("wpr_files", exist_ok=True)
        file_path = f"wpr_files/{file_name}"
        with open(file_path, "wb") as f:
            f.write(response.content)

        result = subprocess.run(
            ["python", "pipeline.py", "--file", file_path],
            capture_output=True,
            text=True,
            timeout=300
        )

        if result.returncode == 0:
            return jsonify({
                "status":  "success",
                "message": f"Pipeline completed for {file_name}",
                "output":  result.stdout[-500:]
            }), 200
        else:
            return jsonify({
                "status": "error",
                "message": result.stderr[-500:]
            }), 500

    except Exception as e:
        return jsonify({
            "status":  "error",
            "message": str(e)
        }), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)