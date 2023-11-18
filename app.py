import os
from flask import Flask, render_template, request, redirect, url_for
from PIL import Image
app = Flask(__name__)

@app.route('/')
def hello():
    return render_template("index.html")


if __name__ == '__main__':
    app.run(debug=True)