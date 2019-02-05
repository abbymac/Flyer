#flyer api

#start boilerplate
from __future__ import print_function
import uuid
import time
from apiclient import discovery 
from httplib2 import Http 
from oauth2client import file, client, tools

start = time.time()
firstsec = time.time()
tempID = '1OnKnCJv7FzE6JMTyWt_TvIXh7cSt-h98hOotI8zDX_g'
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
#end boilerplate

# start sheets section
print('Getting sheets text')
sheetID = '1KhJHmWToEP2SrTLrznx8bx83ki7swhnuiyQebZ48ueg' #published google sheet we use as reference: titled- 'Flyer API input' in knotel account
properties = SHEETS.spreadsheets().values().get(range='Sheet1',spreadsheetId=sheetID).execute().get('values')   

images = {}
Keys = []

propAttrs = {}
for i, val in enumerate(properties):
    if (i == 0):
        for r, keyset in enumerate(val):        #sets keys as the first row
            Keys.append(keyset)
    elif (i > 0):
        tempdict = {}
        for j, key in enumerate(Keys):
            if (key == 'Features'):
                result = [x.strip() for x in properties[i][j].split(';')]       #turn the features column into an array
                tempdict[key] = result
            else: 
                tempdict[key] = properties[i][j]       #sets key to be equal to the corresponding value
        propAttrs['prop' + str(i)] = tempdict         #adds to larger dictionary containing all the property
#### end of sheets section ####

for key in propAttrs:
    for col in propAttrs[key]:
        if(col == 'Features'):
            for i, feat in enumerate(propAttrs[key][col]):      #add bullets to the features
                tempstr = '\u2022'+ ' ' + feat
                feat = tempstr
                propAttrs[key][col][i] = feat
for some in propAttrs:   
    property = propAttrs[some]['Number'] + ' ' + propAttrs[some]['Address'] + ' ' + propAttrs[some]['Floor']
    images[some] = {
        'heroimagename' : property + ' ' + 'heroImage.jpg',            #make image names to find
        'secondimagename' : property + ' ' + 'secondImage.jpg',
        'fpimagename' : property + ' ' + 'FloorPlan.jpg',
    }
firstsecend = time.time()
print('sheets takes: ',firstsecend-firstsec)

def FindImageLoc(slide, obj):
    print('grabbing image placeholder')
    for obj in slide['pageElements']:
        if 'shape' in obj.keys():
            if 'shapeType' in obj['shape'].keys():
                if obj['shape']['shapeType'] == 'RECTANGLE':                    
                    return obj

def pullImg(image):
    rsp = DRIVE.files().list(
        corpora = 'teamDrive',
        includeTeamDriveItems = True,
        q="name='%s'" % image,
        spaces = 'drive',
        supportsTeamDrives= True,
        teamDriveId='0ALFQiBxFeGSCUk9PVA').execute()['files'][0]

    imageURL = '%s&access_token=%s' % (
            DRIVE.files().get_media(fileId=rsp['id']).uri, creds.access_token)
    
    return imageURL

def MakeFlyer(DECK_ID, DATA, val):
    flyermakes = time.time()
    obj = None
    slide1 = SLIDES.presentations().get(presentationId=DECK_ID,
            fields='slides').execute().get('slides', [])[0] #first slide
    slide2 = SLIDES.presentations().get(presentationId=DECK_ID,
            fields='slides').execute().get('slides', [])[1] #Second slide
    obj = FindImageLoc(slide1, obj)
    
    HERO_IMAGE = images[val]['heroimagename']
    SECOND_IMAGE = images[val]['secondimagename']
    FLOORPLAN_IMAGE = images[val]['fpimagename']

    print('pull hero image')
    HEROIMG_url = pullImg(HERO_IMAGE)

    print('pull second image')
    SECONDIMG_url = pullImg(SECOND_IMAGE)

    print('pull fp image')
    FPIMG_url = pullImg(FLOORPLAN_IMAGE)

    print('replacing placeholder texts and image')
    
    reqs = [
        #first page text replacements
        {'replaceAllText': {
            'containsText': {'text':'num', 'matchCase': True},                      #number
            'replaceText': propAttrs[val]['Number'],
            }
        },
        {'replaceAllText': {
            'containsText': {'text':'{{ADDRESS}}', 'matchCase': True},              #address
            'replaceText': propAttrs[val]['Address'],
            }
        },
        {'replaceAllText': {
            'containsText': {'text':'{{FLOOR}}', 'matchCase': True},                #floor       
            'replaceText': propAttrs[val]['Floor'],
            }
        },
        {'replaceAllText': {
            'containsText': {'text':'{{CROSS_STREETS}}', 'matchCase': True},        #cross streets       
            'replaceText': propAttrs[val]['Cross'],
            }
        },
        {'replaceAllText': {
            'containsText': {'text':'{{SUITE}}', 'matchCase': True},                #suite        
            'replaceText': propAttrs[val]['Suite'],
            }
        },
        {'replaceAllText': {
            'containsText': {'text':'{{SF}}', 'matchCase': True},                   #square feet      
            'replaceText': propAttrs[val]['SF'],
            }
        },
        {'replaceAllText': {    
            'containsText': {'text':'{{TERM}}', 'matchCase': True},                 #term      
            'replaceText': propAttrs[val]['TERM'],
            }
        },
        {'replaceAllText': {
            'containsText': {'text':'{{DATE}}', 'matchCase': True},                 #available date      
            'replaceText': propAttrs[val]['available'],
            }
        },

        { 'replaceAllShapesWithImage' : {                                           #hero image
            'imageUrl' : HEROIMG_url,
            'imageReplaceMethod' : 'CENTER_CROP',
            'pageObjectIds': slide1['objectId'],
            'containsText' : {
                'text': 'heroImage',
                'matchCase': True,
            },
            }
        },
        { 'replaceAllShapesWithImage' : {                                           #second page image
            'imageUrl' : SECONDIMG_url,
            'imageReplaceMethod' : 'CENTER_CROP',
            'pageObjectIds': slide2['objectId'],
            'containsText' : {
                'text': 'smallInterior',
                'matchCase': True,
            },
            }
        },      
        { 'replaceAllShapesWithImage' : {                                           #floor plan image
            'imageUrl' : FPIMG_url,
            'imageReplaceMethod' : 'CENTER_CROP',
            'pageObjectIds': slide2['objectId'],
            'containsText' : {
                'text': 'floorPlan',
                'matchCase': True,
            },
            }
        },
        
    ]

    length = len(propAttrs[val]['Features'])
    numbul = 6
    for x in range(0,numbul):
        if (x >= length):
            curbullet = 'bullet'+str(x+1)
            newreq = {
                'replaceAllText': {
                    'containsText': {'text':curbullet, 'matchCase': True},               #if fewer than 6 bullets remove text      
                    'replaceText': '',
                },
            }
        else:
            curbullet = 'bullet'+str(x+1)
            newreq = {
                'replaceAllText': {
                    'containsText': {'text':curbullet, 'matchCase': True},               #bullet    
                    'replaceText': propAttrs[val]['Features'][x],
                },
            }
        reqs.append(newreq)
    
    SLIDES.presentations().batchUpdate(body={'requests': reqs},
        presentationId=DECK_ID, fields='').execute()
    
    print('done')
    flyerend = time.time()
    print('replacement section takes: ',flyermakes-flyerend)

# start drive section 
for g, val in enumerate(propAttrs):  #go through each property in propAttrs

    propts = time.time()
    rsp = DRIVE.files().list(
        corpora = 'teamDrive',
        includeTeamDriveItems = True,
        q = "name='Flyer Template API'",
        supportsTeamDrives = True,
        teamDriveId = '0ALFQiBxFeGSCUk9PVA'
    ).execute()['files'][0]
    
    flyerName = propAttrs[val]['Number']+ ' ' + propAttrs[val]['Address']+ ' ' + propAttrs[val]['Floor']+ ' ' + propAttrs[val]['Suite']
    DATA = {'name': flyerName}  
    print('copying flyer template as %r' %DATA['name'])
    # DECK_ID = DRIVE.files().copy(body=DATA, fileId=tempID).execute()['id']
    DECK_ID = DRIVE.files().copy(body=DATA, fileId=rsp['id'], supportsTeamDrives = True).execute()['id'] #create new presentation to work out of 
    
    FOLDER_ID = propAttrs[val]['FolderId']
    # Retrieve the existing parents to remove
    file = drive_service.files().get(fileId=DECK_ID, fields='parents', supportsTeamDrives = True).execute()
    previous_parents = ",".join(file.get('parents'))
    file = drive_service.files().update(fileId=DECK_ID, addParents=FOLDER_ID, removeParents=previous_parents, fields='id, parents', supportsTeamDrives = True).execute()


    propte = time.time()
    print('copying and moving takes: ', propts - propte)
    MakeFlyer(DECK_ID, DATA, val) 





end = time.time()
print(end-start)
