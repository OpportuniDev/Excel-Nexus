from flask import Flask, request, send_from_directory
import tempfile
app = Flask(__name__)

temp_dir = tempfile.TemporaryDirectory()

Upload_folder = temp_dir.name
Uploadedfile = None
 
@app.route("/file", methods = [ "GET", "POST"])
def File_handler():
    global Uploadedfile

    if request.method == "POST":
     file = request.files.get("file")
     file.save(f"{Upload_folder}/{file.filename}")
     Uploadedfile = file.filename
     return f"File {file.filename} uploaded successfully"

    if request.method == "GET":
        if Uploadedfile is None:
            return "No file uploaded yet", 404
    if Uploadedfile:
        return send_from_directory(Upload_folder, Uploadedfile, as_attachment=True)
    

@app.route("/")
def index():
 return "Opportunidev VBA Updater"
    
if __name__ == "__main__":    app.run(debug=True)