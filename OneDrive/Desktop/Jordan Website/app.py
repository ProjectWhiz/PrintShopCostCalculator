from flask import Flask, send_from_directory, jsonify
from dotenv import load_dotenv
import os


load_dotenv()

app = Flask(__name__, static_folder='frontend/build')

# Serve React frontend
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve_react(path):
    # Serve static files (e.g., CSS, JS, images) if they exist
    if path != "" and os.path.exists(os.path.join(app.static_folder, path)):
        return send_from_directory(app.static_folder, path)
    # Serve the React app's index.html for all other routes
    else:
        return send_from_directory(app.static_folder, 'index.html')

# Fallback route for unmatched API or frontend routes
@app.errorhandler(404)
def not_found(e):
    return send_from_directory(app.static_folder, 'index.html')

# API routes
@app.route('/api/example', methods=['GET'])
def example_api():
    return jsonify({"message": "This is an example API route."})



if __name__ == '__main__':
    app.run(debug=True)










