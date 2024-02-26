# Welcome to Cloud Functions for Firebase for Python!
# To get started, simply uncomment the below code or create your own.
# Deploy with `firebase deploy`

import time
import json

from firebase_functions import https_fn
from firebase_admin import initialize_app
from google.cloud import firestore

import excelToJson

# initialize_app()
#
#
# @https_fn.on_request()
# def on_request_example(req: https_fn.Request) -> https_fn.Response:
#     return https_fn.Response("Hello world!")

app = initialize_app()

@https_fn.on_request()
def convertSheet(request: https_fn.Request) -> https_fn.Response:
    t = time.process_time()

    """Take the text parameter passed to this HTTP endpoint and insert it into
    a new document in the messages collection."""
    # Grab the text parameter.
    # sheet = req.args["sheet"]
    print(request)

    body = request.get_json(silent=True)
    sheet = None
    
    try:
        sheet = body["sheet"]
    except:
        print("bad request")

    if sheet is None:
        return https_fn.Response("No sheet provided", status=400)
    
    file = excelToJson.readFileFromBase64(sheet)

    json = excelToJson.exportJson(file)

    json = json.replace("'", '"')

    # Push the new message into Cloud Firestore using the Firebase Admin SDK.
    elapsed_time = time.process_time() - t

    firestore_client: firestore.Client = firestore.Client()

    # Push the new character into Cloud Firestore using the Firebase Admin SDK.
    responseObject = '{"time": "all done at ' + str(elapsed_time) + ' seconds", "sheet": ' + json + '}'

    try:
        firestore_client.collection("characters").add({"time": str(elapsed_time), "sheet": json})
    except:
        print("error saving character")

    print('{"time": "all done at ' + str(elapsed_time) + ' seconds"')


    # Send back a message that we've successfully written the message
    return https_fn.Response(responseObject)
