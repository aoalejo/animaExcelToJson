# Welcome to Cloud Functions for Firebase for Python!
# To get started, simply uncomment the below code or create your own.
# Deploy with `firebase deploy`

import time
import hashlib

from firebase_functions import https_fn
from firebase_admin import initialize_app
from google.cloud import firestore

import excelToJson
app = initialize_app()

@https_fn.on_request()
def convertSheet(request: https_fn.Request) -> https_fn.Response:
    t = time.process_time()
    print(request.method)

    # Set CORS headers for the preflight request
    if request.method == "OPTIONS":
        # Allows POST requests from app origin with the Content-Type
        # header and caches preflight response for an 3600s
        headers = {
            "Access-Control-Allow-Origin": "https://amt-v3.web.app",
            "Access-Control-Allow-Methods": "POST",
            "Access-Control-Allow-Headers": "Content-Type, Accept",
            "Access-Control-Max-Age": "300",
        }

        return ("", 204, headers)
    
    """Take the text parameter passed to this HTTP endpoint and insert it into
    a new document in the messages collection."""
    print(request.headers)

    body = request.get_json(silent=True)

    sheet = None
    json = None
    file = None
    usedFile = False

    try:
        sheet = body["sheet"]
        print("sheet in body provided")
    except:
        print("No body provided")

    try:
        sheet = request.files["sheet"]
        print("sheet in files provided")
        usedFile = True

    except:
        print("No multipart file provided")

    if sheet is None:
        return https_fn.Response("No sheet provided", status=400)
    
    try:
        if (usedFile):
            print("readFile")
            file = excelToJson.readFile(sheet)

        else:
            print("readFileFromBase64")
            file = excelToJson.readFileFromBase64(sheet)

        json = excelToJson.exportJson(file)
        json = json.replace("'", '"')
    except:
        print("error converting sheet")
        return https_fn.Response("Error converting sheet", status=500)
    
    if json is None:
        return https_fn.Response("Error converting sheet", status=500)

    # Push the new message into Cloud Firestore using the Firebase Admin SDK.
    elapsed_time = time.process_time() - t

    m = hashlib.sha256(json.encode(), usedforsecurity=False)
    objectId = str(m.hexdigest())

    # Push the new character into Cloud Firestore using the Firebase Admin SDK.
    responseObject = '{"time": "all done at ' + str(elapsed_time) + ' seconds", "objectId": "' + objectId + '", "sheet": ' + json + '}'

    try:
        firestore_client: firestore.Client = firestore.Client()
        firestore_client.collection("characters").add({"time": str(elapsed_time), "sheet": json}, document_id=objectId)
    except:
        print("error saving character")

    print('{"time": "all done at ' + str(elapsed_time) + ' seconds"')

    headers = {"Access-Control-Allow-Origin": "*"}

    # Send back a message that we've successfully written the message
    return https_fn.Response(responseObject, 200, headers)