#flyer api

#start boilerplate
from __future__ import print_function
import uuid
from apiclient import discovery 
from httplib2 import Http 
from oauth2client import file, client, tools

# HERO_IMAGE = 'heroImage.jpg'
# SECOND_IMAGE = 'secondImage.jpg'
# FLOORPLAN_IMAGE = 'FloorPlan.jpg'
tempID = '1WEPSrLZ4LiMFMOTsqzVat6Ko1qHe3fB_i-NcBK_w35w'
TMPLFILE = 'Test Flyer Template API'
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

#### start sheets section ####
print('Getting sheets text')
sheetID = '1R3kO2uRPwtzSTyHQJ1kgFKAN-obAmLOtziW5c9OCRrQ' #published google sheet we use as reference: titled- 'testslideapi' in personal google accout
properties = SHEETS.spreadsheets().values().get(range='Sheet1',spreadsheetId=sheetID).execute().get('values')   #get all values in sheet1
#properties ends up as: [all row 1 values],[all row 2 values],[all row 3 values], etc

#Question- should we have 1 sheet with multiple tabs for each market and store as different var?^
print('getting text info from sheets')
sheet = SHEETS.spreadsheets().get(spreadsheetId=sheetID, 
        ranges=['Sheet1']).execute().get('sheets')[0]

images = {}
Keys = []
#Question- if we have one google sheet to host all of these- do we re-make all the flyers each time the API is run?
    # - should we look for cells that have changed since last run and only do those?

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


def FindImageLoc(slide, obj):
    print('grabbing image placeholder')
    #find picture 
    for obj in slide['pageElements']:
        if 'shape' in obj.keys():
            if 'shapeType' in obj['shape'].keys():
                if obj['shape']['shapeType'] == 'RECTANGLE':                    
                    return obj

def MakeFlyer(DECK_ID, DATA, val):
    obj = None
    slide1 = SLIDES.presentations().get(presentationId=DECK_ID,
            fields='slides').execute().get('slides', [])[0] #first slide
    slide2 = SLIDES.presentations().get(presentationId=DECK_ID,
            fields='slides').execute().get('slides', [])[1] #Second slide
    obj = FindImageLoc(slide1, obj)
    
    HERO_IMAGE = images[val]['heroimagename']
    SECOND_IMAGE = images[val]['secondimagename']
    FLOORPLAN_IMAGE = images[val]['fpimagename']

    print('pull hero image to use')
    rsp = DRIVE.files().list(q="name='%s'" % HERO_IMAGE).execute()['files'][0]
    print('found the image %r' % rsp['name'])
    HEROIMG_url = '%s&access_token=%s' % (
            DRIVE.files().get_media(fileId=rsp['id']).uri, creds.access_token)

    print('pull secondary image to use')
    rsp = DRIVE.files().list(q="name='%s'" % SECOND_IMAGE).execute()['files'][0]
    print('found the image %r' % rsp['name'])
    SECONDIMG_url = '%s&access_token=%s' % (
            DRIVE.files().get_media(fileId=rsp['id']).uri, creds.access_token)

    print('pull floorplan to use')
    rsp = DRIVE.files().list(q="name='%s'" % FLOORPLAN_IMAGE).execute()['files'][0]
    print('found the image %r' % rsp['name'])
    FPIMG_url = '%s&access_token=%s' % (
            DRIVE.files().get_media(fileId=rsp['id']).uri, creds.access_token)


    print('replacing placeholder texts and image')
    #
    reqs = [
        #first page text replacements
        {'replaceAllText': {
            'containsText': {'text':'num', 'matchCase': True},                      #number
            'replaceText': propAttrs[val]['Number'],
            }
        },
        {'replaceAllText': {
            'containsText': {'text':'{{ADDRESS}}', 'matchCase': True},                     #address
            'replaceText': propAttrs[val]['Address'],
            }
        },
        {'replaceAllText': {
            'containsText': {'text':'{{FLOOR}}', 'matchCase': True},               #floor       
            'replaceText': propAttrs[val]['Floor'],
            }
        },
        {'replaceAllText': {
            'containsText': {'text':'{{CROSS_STREETS}}', 'matchCase': True},              #cross streets       
            'replaceText': propAttrs[val]['Cross'],
            }
        },
        {'replaceAllText': {
            'containsText': {'text':'{{SUITE}}', 'matchCase': True},              #suite        
            'replaceText': propAttrs[val]['Suite'],
            }
        },
        {'replaceAllText': {
            'containsText': {'text':'{{SF}}', 'matchCase': True},               #square feet      
            'replaceText': propAttrs[val]['SF'],
            }
        },
        {'replaceAllText': {    
            'containsText': {'text':'{{TERM}}', 'matchCase': True},               #term      
            'replaceText': propAttrs[val]['TERM'],
            }
        },
        {'replaceAllText': {
            'containsText': {'text':'{{DATE}}', 'matchCase': True},               #available date      
            'replaceText': propAttrs[val]['available'],
            }
        },

        { 'replaceAllShapesWithImage' : {
            'imageUrl' : HEROIMG_url,
            'imageReplaceMethod' : 'CENTER_CROP',
            'pageObjectIds': slide1['objectId'],
            'containsText' : {
                'text': 'heroImage',
                'matchCase': True,
            },
            }
        },
        { 'replaceAllShapesWithImage' : {
            'imageUrl' : SECONDIMG_url,
            'imageReplaceMethod' : 'CENTER_CROP',
            'pageObjectIds': slide2['objectId'],
            'containsText' : {
                'text': 'smallInterior',
                'matchCase': True,
            },
            }
        },
        { 'replaceAllShapesWithImage' : {
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
                    'containsText': {'text':curbullet, 'matchCase': True},               #available date      
                    'replaceText': '',
                },
            }
        else:
            curbullet = 'bullet'+str(x+1)
            newreq = {
                'replaceAllText': {
                    'containsText': {'text':curbullet, 'matchCase': True},               #available date      
                    'replaceText': propAttrs[val]['Features'][x],
                },
            }
        reqs.append(newreq)



    
    SLIDES.presentations().batchUpdate(body={'requests': reqs},
        presentationId=DECK_ID, fields='').execute()
    
    print('done')


#---- start drive section ----#
for g, val in enumerate(propAttrs):  #go through each property in propAttrs

    # rsp = DRIVE.files().list(q="name='%s'" % TMPLFILE).execute()['files'][0]    #q parameter is the search query


    #searches thru drive files and finds the template file 
    flyerName = propAttrs[val]['Number']+ ' ' + propAttrs[val]['Address']+ ' ' + propAttrs[val]['Floor']+ ' ' + propAttrs[val]['Suite']
    DATA = {'name': flyerName}   #Question- need to make this a dynamic title -> concatenate the address/floor/suite for unique id? Override if already exists-> update flyer
    print('copying flyer template as %r' %DATA['name']))
    DECK_ID = DRIVE.files().copy(body=DATA, fileId=tempID).execute()['id']
    # DECK_ID = DRIVE.files().copy(body=DATA, fileId=rsp['id']).execute()['id'] #create new presentation to work out of 
    
    MakeFlyer(DECK_ID, DATA, val) 

    # folderName = propAttrs[val]['Number']+ ' ' + propAttrs[val]['Address']
    # moveID = DRIVE.files().list(q="name='%s'" % folderName).execute()['id']
    # Retrieve the existing parents to remove
    file = drive_service.files().get(fileId=DECK_ID, fields='parents').execute()
    # previous_parents = ",".join(file.get('parents'))
    # Move the file to the new folder
    FOLDER_ID = propAttrs[val]['FolderId']

    file = drive_service.files().update(fileId=DECK_ID, addParents=FOLDER_ID, fields='id, parents').execute()


#raymond-- naming structure for presentation 
# automate to turn presentation into pdf 
# files: export  
####end drive section

