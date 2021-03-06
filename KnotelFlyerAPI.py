# flyer api

# start boilerplate
from __future__ import print_function
import uuid
import time
import io
import os.path
import webbrowser
from io import StringIO
from apiclient import discovery
from apiclient.http import MediaIoBaseDownload
from httplib2 import Http
from oauth2client import file, client, tools
from apiclient.http import MediaFileUpload

start = time.time()
firstsec = time.time()
tempID = ''
TMPLFILE = 'Flyer Template API'

SCOPES = {
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/presentations',
}
store = file.Storage('token.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
    creds = tools.run_flow(flow, store)
HTTP = creds.authorize(Http())
DRIVE = discovery.build('drive', 'v3', http=HTTP)
drive_service = discovery.build('drive', 'v3', http=creds.authorize(Http()))
SHEETS = discovery.build('sheets', 'v4', http=HTTP)
SLIDES = discovery.build('slides', 'v1', http=HTTP)


def gen_uuid(): return str(uuid.uuid4())

# end boilerplate


# start sheets section
print('Getting sheets text')
# published sheet ID to read from
sheetID = '__'
properties = SHEETS.spreadsheets().values().get(
    range='Sheet1', spreadsheetId=sheetID).execute().get('values')

# Google Drive ID to operate in
driveID = '__'
images = {}
Keys = []

propAttrs = {}
for i, val in enumerate(properties):
    if (i == 0):
        for r, keyset in enumerate(val):  # sets keys as the first row
            Keys.append(keyset)
    elif (i > 0):
        tempdict = {}
        for j, key in enumerate(Keys):
            if (key == 'Features'):
                # turn the features column into an array
                result = [x.strip() for x in properties[i][j].split(';')]
                tempdict[key] = result
            else:
                # sets key to be equal to the corresponding value
                tempdict[key] = properties[i][j]
        # adds to larger dictionary containing all the property
        propAttrs['prop' + str(i)] = tempdict

#### end of sheets section ####

for key in propAttrs:
    for col in propAttrs[key]:
        if(col == 'Features'):
            # add bullets to the features
            for i, feat in enumerate(propAttrs[key][col]):
                tempstr = '\u2022' + ' ' + feat
                feat = tempstr
                propAttrs[key][col][i] = feat

# calculate google sheet section run time
firstsecend = time.time()
print('sheets takes: ', firstsecend-firstsec)


def FindImageLoc(slide, obj):
    print('grabbing image placeholder')
    for obj in slide['pageElements']:
        if 'shape' in obj.keys():
            if 'shapeType' in obj['shape'].keys():
                if obj['shape']['shapeType'] == 'RECTANGLE':
                    return obj


def pullImg(image):
    imageURL = '%s&access_token=%s' % (
        DRIVE.files().get_media(fileId=image).uri, creds.access_token)
    return imageURL


def MakeFlyer(DECK_ID, DATA, val):
    flyermakes = time.time()
    obj = None
    slide1 = SLIDES.presentations().get(presentationId=DECK_ID,
                                        fields='slides').execute().get('slides', [])[0]  # first slide
    slide2 = SLIDES.presentations().get(presentationId=DECK_ID,
                                        fields='slides').execute().get('slides', [])[1]  # Second slide
    obj = FindImageLoc(slide1, obj)

    HERO_IMAGE = propAttrs[val]['heroImg']
    SECOND_IMAGE = propAttrs[val]['secImg']
    FLOORPLAN_IMAGE = propAttrs[val]['fpImg']

    print('pull hero image')
    HEROIMG_url = pullImg(HERO_IMAGE)

    print('pull second image')
    SECONDIMG_url = pullImg(SECOND_IMAGE)

    print('pull fp image')
    FPIMG_url = pullImg(FLOORPLAN_IMAGE)

    print('replacing placeholder texts and image')

    reqs = [
        # first page text replacements
        {'replaceAllText': {
            'containsText': {'text': 'num', 'matchCase': True},  # number
            'replaceText': propAttrs[val]['Number'],
        }
        },
        {'replaceAllText': {
            # address
            'containsText': {'text': '{{ADDRESS}}', 'matchCase': True},
            'replaceText': propAttrs[val]['Address'],
        }
        },
        {'replaceAllText': {
            'containsText': {'text': '{{FLOOR}}', 'matchCase': True},  # floor
            'replaceText': propAttrs[val]['Floor'],
        }
        },
        {'replaceAllText': {
            # cross streets
            'containsText': {'text': '{{CROSS_STREETS}}', 'matchCase': True},
            'replaceText': propAttrs[val]['Cross'],
        }
        },
        {'replaceAllText': {
            'containsText': {'text': '{{SUITE}}', 'matchCase': True},  # suite
            'replaceText': propAttrs[val]['Suite'],
        }
        },
        {'replaceAllText': {
            # square feet
            'containsText': {'text': '{{SF}}', 'matchCase': True},
            'replaceText': propAttrs[val]['SF'],
        }
        },
        {'replaceAllText': {
            'containsText': {'text': '{{TERM}}', 'matchCase': True},  # term
            'replaceText': propAttrs[val]['TERM'],
        }
        },
        {'replaceAllText': {
            # available date
            'containsText': {'text': '{{DATE}}', 'matchCase': True},
            'replaceText': propAttrs[val]['available'],
        }
        },

        {'replaceAllShapesWithImage': {  # hero image
            'imageUrl': HEROIMG_url,
            'imageReplaceMethod': 'CENTER_CROP',
            'pageObjectIds': slide1['objectId'],
            'containsText': {
                'text': 'heroImage',
                'matchCase': True,
            },
        }
        },
        {'replaceAllShapesWithImage': {  # second page image
            'imageUrl': SECONDIMG_url,
            'imageReplaceMethod': 'CENTER_CROP',
            'pageObjectIds': slide2['objectId'],
            'containsText': {
                'text': 'smallInterior',
                'matchCase': True,
            },
        }
        },
        {'replaceAllShapesWithImage': {  # floor plan image
            'imageUrl': FPIMG_url,
            'imageReplaceMethod': 'CENTER_CROP',
            'pageObjectIds': slide2['objectId'],
            'containsText': {
                'text': 'floorPlan',
                'matchCase': True,
            },
        }
        },

    ]

    length = len(propAttrs[val]['Features'])
    numbul = 6
    for x in range(0, numbul):
        if (x >= length):
            curbullet = 'bullet'+str(x+1)
            newreq = {
                'replaceAllText': {
                    # if fewer than 6 bullets remove text
                    'containsText': {'text': curbullet, 'matchCase': True},
                    'replaceText': '',
                },
            }
        else:
            curbullet = 'bullet'+str(x+1)
            newreq = {
                'replaceAllText': {
                    # bullet
                    'containsText': {'text': curbullet, 'matchCase': True},
                    'replaceText': propAttrs[val]['Features'][x],
                },
            }
        reqs.append(newreq)

    SLIDES.presentations().batchUpdate(body={'requests': reqs},
                                       presentationId=DECK_ID, fields='').execute()

    print('done')
    flyerend = time.time()
    print('replacement section takes: ', flyermakes-flyerend)


def Export(file_id, filename):
    MIMETYPE = 'application/pdf'
    data = DRIVE.files().export(fileId=DECK_ID, mimeType=MIMETYPE).execute()
    if data:
        fn = '%s.pdf' % os.path.splitext(filename)[0]
        with open(fn, 'wb') as fh:
            fh.write(data)
        print('downloaded "%s" (%s)' % (fn, MIMETYPE))


def Upload(file_id, filename, FOLDER_ID, driveId):
    filename = filename + '.pdf'
    metadata = {'name': filename, 'mimetype': 'application/pdf',
                'teamDriveId': driveId, 'parents': FOLDER_ID, 'supportsTeamDrives': True}
    res = DRIVE.files().create(body=metadata, media_body=filename).execute()
    if res:
        print('Uploaded "%s"' % filename)
        print('id', res['id'])
        pdfID = res['id']
        print('id to move into', pdfID)
        file = drive_service.files().update(fileId=pdfID, addParents=FOLDER_ID,
                                            fields='id, parents', supportsTeamDrives=True).execute()


# start drive section
for g, val in enumerate(propAttrs):  # go through each property in propAttrs

    propts = time.time()
    rsp = DRIVE.files().list(
        corpora='teamDrive',
        includeTeamDriveItems=True,
        q="name='Flyer Template API'",
        supportsTeamDrives=True,
        teamDriveId=driveID,
    ).execute()['files'][0]

    flyerName = propAttrs[val]['Number'] + ' ' + propAttrs[val]['Address'] + \
        ' ' + propAttrs[val]['Floor'] + ' ' + propAttrs[val]['Suite']
    DATA = {'name': flyerName}
    print('copying flyer template as %r' % DATA['name'])
    DECK_ID = DRIVE.files().copy(body=DATA, fileId=rsp['id'], supportsTeamDrives=True).execute()[
        'id']  # create new presentation to work out of

    FOLDER_ID = propAttrs[val]['FolderId']
    # Retrieve the existing parents to remove
    file = drive_service.files().get(fileId=DECK_ID, fields='parents',
                                     supportsTeamDrives=True).execute()
    previous_parents = ",".join(file.get('parents'))
    file = drive_service.files().update(fileId=DECK_ID, addParents=FOLDER_ID,
                                        removeParents=previous_parents, fields='id, parents', supportsTeamDrives=True).execute()
    print('deck-id is', DECK_ID)
    propte = time.time()
    print('copying and moving takes: ', propts - propte)
    MakeFlyer(DECK_ID, DATA, val)
    Export(DECK_ID, flyerName)
    Upload(DECK_ID, flyerName, FOLDER_ID, driveID)


end = time.time()
print(end-start)
