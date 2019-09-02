#############################################################
# IMPORTS
#############################################################
import base64
import os
import shutil
from urllib.parse import quote as urlquote

import dash
import dash_core_components as dcc
import dash_html_components as html
import dash_table
import numpy as np
from dash.dependencies import Input, Output
from expenses_code import parse_message_list
from flask import Flask, send_from_directory


#############################################################
# SETUP
#############################################################
UPLOAD_DIRECTORY = "./UPLOADS"

if not os.path.exists(UPLOAD_DIRECTORY):
    os.makedirs(UPLOAD_DIRECTORY)


# Normally, Dash creates its own Flask server internally. By creating our own,
# we can create a route for downloading files directly:
server = Flask(__name__)
app = dash.Dash(server=server)


@server.route("/download/<path:path>")
def download(path):
    """Serve a file from the upload directory."""
    return send_from_directory(UPLOAD_DIRECTORY, path, as_attachment=True)


#############################################################
# APP LAYOUT
#############################################################
app.layout = html.Div(
    [
        html.H1("Email Expenses Parser"),
        html.H2("Upload Expense Emails"),
        html.P("This expense parser currently accepts emails from:"),
        html.Li("uber.uae@uber.com (Uber taxis)"),
        html.Li("support@deliveroo.ae"),
        html.Li("mmc@bcdtravel.ae"),
        html.P(
            "Contact Stephen.Rawson@OliverWyman.com if you receive other expenses via email and would like to parse them too."
        ),
        dcc.Upload(
            id="upload-data",
            children=html.Div(
                [
                    "Drag and drop or click to select your files to upload (upload all at once)."
                ]
            ),
            style={
                "width": "100%",
                "height": "60px",
                "lineHeight": "60px",
                "borderWidth": "1px",
                "borderStyle": "dashed",
                "borderRadius": "5px",
                "textAlign": "center",
                "margin": "10px",
            },
            multiple=True,
        ),
        html.H2("Output and uploaded emails"),
        html.Ul(id="file-list"),
    ],
    style={"max-width": "80%", "margin": "50px"},
)


def save_file(name, content, minipath):
    """Decode and store a file uploaded with Plotly Dash."""
    print(f"Saving content for email: {name}")
    data = content.encode("utf8").split(b";base64,")[1]
    with open(os.path.join(UPLOAD_DIRECTORY, minipath, name), "wb") as fp:
        fp.write(base64.decodebytes(data))


def uploaded_files(minipath="."):
    """List the files in the upload directory."""
    files = []
    for filename in os.listdir(os.path.join(UPLOAD_DIRECTORY, minipath)):
        path = os.path.join(UPLOAD_DIRECTORY, minipath, filename)
        if os.path.isfile(path):
            files.append(path)
    return files


def create_output_excel(minipath):
    df, skipped = parse_message_list(uploaded_files(minipath))
    df.to_excel(os.path.join(UPLOAD_DIRECTORY, minipath, "Email_Expenses.xlsx"), index=False)
    
    return df


def file_download_link(filename, minipath = '', output=False):
    """Create a Plotly Dash 'A' element that downloads a file from the app."""
    if output:
        location = f"/download/{minipath}/{urlquote(filename)}"
        print(location)
        return html.A("DOWNLOAD EXPENSES DATA", href=location)
    else:
        location = f"/download/{minipath}/{urlquote(filename)}"
        print(location)
        return html.A(os.path.basename(filename), href=location)


@app.callback(
    Output("file-list", "children"),
    [Input("upload-data", "filename"), Input("upload-data", "contents")],
)
def update_output(uploaded_filenames, uploaded_file_contents):
    """Save uploaded files and regenerate the file list."""

    minipath = f"{str(np.random.rand()).replace('.', '')}"

    dirs = os.listdir(UPLOAD_DIRECTORY)
    for d in dirs:
        if d != minipath and "output" not in d.lower():
            try:
                shutil.rmtree(os.path.join(UPLOAD_DIRECTORY, d))
                print(f"Removed {d}")
            except Exception as e:
                print(e)
        if "output" in d.lower() and minipath not in d.lower():
            try:
                os.remove(os.path.join(UPLOAD_DIRECTORY, d))
                print(f"Removed {d}")
            except Exception as e:
                print(e)

    os.makedirs(os.path.join(UPLOAD_DIRECTORY, minipath))
    if uploaded_filenames is not None and uploaded_file_contents is not None:
        for name, data in zip(uploaded_filenames, uploaded_file_contents):
            save_file(name, data, minipath)
        df = create_output_excel(minipath)

    output_files = [f for f in uploaded_files(minipath) if 'Email_Expenses' in f]
    other_files = [f for f in uploaded_files(minipath) if 'Email_Expenses' not in f]
    print(output_files)
    print(other_files)
    
    if len(output_files) == 0:
        return [html.Li("Upload your emails then download the output here.")]
    else:
        return (
            [
                html.Li(
                    file_download_link(
                        os.path.basename(filename),
                        minipath,
                        output=True)
                )
                for filename in output_files
                if "Email_Expenses" in filename
            ]
            + [html.Hr(style={"margin_top": "25px", "margin_bottom": "25px"})]
            + 
            [dash_table.DataTable(
                id='table',
                columns=[{"name": i, "id": i} for i in df.columns],
                data=df.to_dict("rows"),
                style_table={'overflowX': 'scroll'},
            )
            ]
            + [html.Hr(style={"margin_top": "25px", "margin_bottom": "25px"})]
#             + [
#                 html.Li(
#                     file_download_link(
#                         os.path.basename(filename), minipath
#                     )
#                 )
#                 for filename in other_files
#             ]
        )


if __name__ == "__main__":
    app.run_server(debug=True)
