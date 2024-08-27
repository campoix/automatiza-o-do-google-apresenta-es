from google.oauth2 import service_account
from googleapiclient.discovery import build
from colorama import Fore
SCOPES = ['https://www.googleapis.com/auth/presentations']
SERVICE_ACCOUNT_FILE = 'code/googleapr/credentials.json'

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('slides', 'v1', credentials=credentials)

presentation_id = 'secreto'

slide_ids = {
    'g4dfce81f19_0_45': 0, 'index_slide': 1, 'g20e13fe16bb_10_22': 2, 'g2e549f90254_4_37': 3,
    'g2e549f90254_4_0': 4, 'g2e12e853079_6_31': 5, 'g202e19c39e0_0_18': 6, 'g2df4587c440_1_52': 7,
    'g2df4587c440_2_0': 8, 'g2df4587c440_0_25': 9, 'g2df4587c440_0_49': 10, 'g5f32cd8ae4b5beda_231': 11,
    'g2df4587c440_5_84': 12, 'g4d9475f5c1ac3ce2_0': 13, 'g2df4587c440_3_5': 14, 'g3a4af429ea36baec_63': 15,
    'g2df4587c440_0_71': 16, 'g2df4587c440_5_63': 17, 'g2df4587c440_1_3': 18, 'g2e12e853079_12_71': 19,
    'g2e12e853079_12_1': 20, 'g2df4587c440_3_28': 21, 'g5f32cd8ae4b5beda_0': 22, 'g5f32cd8ae4b5beda_168': 23,
    'g2df4587c440_0_94': 24, 'g2df4587c440_3_157': 25, 'g2df4587c440_1_102': 26, 'g2e12e853079_6_99': 27,
    'g3a4af429ea36baec_0': 28, 'g2df4587c440_3_50': 29, 'g2e12e853079_6_5': 30, 'g5f32cd8ae4b5beda_252': 31,
    'g2df4587c440_3_187': 32, 'g3a4af429ea36baec_90': 33, 'g2e12e853079_13_56': 34, 'g3a4af429ea36baec_21': 35,
    'g2df4587c440_2_142': 36, 'g2e12e853079_12_24': 37, 'g2e12e853079_13_103': 38, 'g2df4587c440_0_0': 39,
    'g2df4587c440_5_0': 40, 'g20e13fe16bb_9_88': 41, 'g2e549f90254_3_0': 42, 'g2df4587c440_1_28': 43,
    'g20e13fe16bb_9_163': 44, 'g20e13fe16bb_9_113': 45, 'g2df4587c440_0_141': 46, 'g2e12e853079_6_73': 47,
    'g2df4587c440_2_95': 48, 'g2df4587c440_3_131': 49, 'g2e12e853079_13_79': 50, 'g2e12e853079_13_13': 51,
    'g2e549f90254_3_23': 52, 'g2e12e853079_12_49': 53, 'g2e12e853079_6_123': 54, 'g2df4587c440_2_117': 55,
    'g2df4587c440_5_105': 56, 'g2df4587c440_2_48': 57, 'g2df4587c440_5_42': 58, 'g2e549f90254_3_52': 59,
    'g5f32cd8ae4b5beda_210': 60, 'g20e13fe16bb_9_138': 61
}
# Grupos de slides
grupos_slides = {
    'mf': [6, 7, ],
    'ifi': [4],
    'gof': [8, ],
    'cf': [],
    'tpf': [],
    'rrf': [3, 5, ],
    'tpe': [],
    'mtrl': [],
    "dc": [2,]
}

# Requests para atualizar a apresentação
requests = [
    {
        'createSlide': {
            'objectId': 'index_slide',
            'insertionIndex': 1,
            'slideLayoutReference': {
                'predefinedLayout': 'TITLE_AND_BODY'
            }
        }
    },
    {
        'createShape': {
            'objectId': 'index_title',
            'shapeType': 'TEXT_BOX',
            'elementProperties': {
                'pageObjectId': 'index_slide',
                'size': {
                    'height': {'magnitude': 50, 'unit': 'PT'},
                    'width': {'magnitude': 300, 'unit': 'PT'}
                },
                'transform': {
                    'scaleX': 1,
                    'scaleY': 1,
                    'translateX': 50,
                    'translateY': 50,
                    'unit': 'PT'
                }
            }
        }
    },
    {
        'insertText': {
            'objectId': 'index_title',
            'insertionIndex': 0,
            'text': 'Índice'
        }
    }
]

for i, (grupo, slides_grupo) in enumerate(grupos_slides.items()):
    box_id = f'{grupo}_box_{i}'
    requests.append({
        'createShape': {
            'objectId': box_id,
            'shapeType': 'TEXT_BOX',
            'elementProperties': {
                'pageObjectId': 'index_slide',
                'size': {
                    'height': {'magnitude': 50, 'unit': 'PT'},
                    'width': {'magnitude': 300, 'unit': 'PT'}
                },
                'transform': {
                    'scaleX': 1,
                    'scaleY': 1,
                    'translateX': 50,
                    'translateY': 100 + (i * 60),
                    'unit': 'PT'
                }
            }
        }
    })
    
    text = f'{grupo.upper()}: Slides {", ".join(map(str, slides_grupo))}'
    
    requests.append({
        'insertText': {
            'objectId': box_id,
            'insertionIndex': 0,
            'text': text
        }
    })
    
    for slide_number in slides_grupo:
        slide_id = slide_ids[slide_number]
        link_text = f'{slide_number}'
        link_start_index = text.find(link_text)
        link_end_index = link_start_index + len(link_text)
        
        requests.append({
            'updateTextStyle': {
                'objectId': box_id,
                'style': {
                    'link': {
                        'url': f'https://docs.google.com/presentation/d/{presentation_id}/edit#slide=id.{slide_id}'
                    }
                },
                'fields': 'link',
                'textRange': {
                    'type': 'FIXED_RANGE',
                    'startIndex': link_start_index,
                    'endIndex': link_end_index
                }
            }
        })

body = {
    'requests': requests
}

response = service.presentations().batchUpdate(
    presentationId=presentation_id,
    body=body
).execute()

print("Slide criado...")
