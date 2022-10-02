"This module contains authentication credentials and other QlikView settings."
import os

QLIKVIEW_PROTOCOL: str = "qvp://"
QLIK_VIEW_SERVER_URL: str = QLIKVIEW_PROTOCOL + os.getenv(
    'QLIK_VIEW_SERVER_URL', '127.0.0.1'
)
QLIKVIEW_SERVER_LOGIN: str = os.getenv('QLIKVIEW_SERVER_LOGIN', 'user')
QLIKVIEW_SERVER_PASSWORD: str = os.getenv(
    'QLIKVIEW_SERVER_PASSWORD', 'password'
)
QLIKVIEW_APP_PATH: str = os.getenv(
    'QLIKVIEW_APP_PATH', 'C:\Program Files\QlikView\Qv.exe'
)
QLIKVIEW_USER: str = os.getenv('QLIKVIEW_USER', 'user')
QLIKVIEW_USER_PASSWORD: str = os.getenv(
    'QLIKVIEW_USER_PASSWORD', 'password'
)
