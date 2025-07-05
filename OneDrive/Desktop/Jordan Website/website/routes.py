from flask import Flask, Blueprint, send_from_directory
import os

app = Blueprint('app', __name__)

# Serve React frontend for all non-API routes
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve_react(path):
    if path != "" and os.path.exists(os.path.join('frontend/build', path)):
        return send_from_directory('frontend/build', path)
    else:
        return send_from_directory('frontend/build', 'index.html')

# Remove old routes that used render_template
# API routes can be added here if needed


