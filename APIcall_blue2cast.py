# -*- coding: utf-8 -*-
"""
Created on Mon Apr 29 17:13:25 2024

@author: LENLUI
"""

# runfile('C:/python/MECADA/mecadaAPI.py', wdir='C:/python/MECADA')

# Note that the default secure protocol TLS v1.2 (SSL) must be enabled for Web API clients so this information will be sent encrypted.
# The request header thus becomes:
# Authorization Basic ZjhiMDEyODAzMTBiNGY4OThlNjI2NzkwYTVkOWYxMGU6WVdSa09EQTRNbUZtTkdZMk5HUXpaV0ZqTURKbFltSmpPVFF6WlRZME9XWT0=


# from Corcel:
import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.poolmanager import PoolManager
from requests.auth import HTTPBasicAuth
import ssl

# Create an SSLContext object with TLS v1.2
ssl_context = ssl.create_default_context()
ssl_context.options |= ssl.OP_NO_TLSv1
ssl_context.options |= ssl.OP_NO_TLSv1_1
ssl_context.set_ciphers('DEFAULT@SECLEVEL=1')

# Create a custom adapter that uses the SSLContext
class TLSAdapter(HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        kwargs['ssl_context'] = ssl_context
        return super(TLSAdapter, self).init_poolmanager(*args, **kwargs)

# Create a session and mount the adapter
session = requests.Session()
session.mount('https://', TLSAdapter())

# Now you can make your post request using the session
# url = 'https://yourapi.com/endpoint'
# data = {'key': 'value'}
# headers = {'Content-Type': 'application/json'}
# response = session.post(url, json=data, headers=headers)
sAuthPublicKey = 'f8b01280310b4f898e626790a5d9f10e'
sAuthSecret = 'YWRkODA4MmFmNGY2NGQzZWFjMDJlYmJjOTQzZTY0OWY='
sAuth = HTTPBasicAuth(sAuthPublicKey, sAuthSecret)

    
sURL = 'https://blue2cast.com/auth/token'
sAuthHeaders = {
    'accept': '*/*',
    'Content-Type': 'application/x-www-form-urlencoded',
    # 'Authorization': 'Basic ZjhiMDEyODAzMTBiNGY4OThlNjI2NzkwYTVkOWYxMGU6WVdSa09EQTRNbUZtTkdZMk5HUXpaV0ZqTURKbFltSmpPVFF6WlRZME9XWT0=',
#Below here maybe not needed    
    # 'Host': 'localhost:5001',
    # 'Accept-Encoding': 'gzip, deflate, br',
    # 'Connection': 'keep-alive',
    # 'Content-Length': '44'
}
response = session.post(url=sURL, auth=sAuth, headers = sAuthHeaders)
# Check the response
print('--Authentication request--')
print(f'Response status: {response.status_code}')
print(f'Response text  : {response.text}\n')

# Now an authenticated session, construct new header for authorized requests.
sHeaders = {
    'accept': '*/*',
    # 'accept': 'application/json',
    # 'Content-Type': 'application/x-www-form-urlencoded',
    # 'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + response.json()['access_token'],
}

sDataURL = 'https://blue2cast.com/oauthapi/MeasurementsBatch?hoursAgo=1'
r = session.post(url=sDataURL, headers=sHeaders)
print('--Data request--')
print(f'Response status: {r.status_code}')
print(f'Response text  : {r.text}\n')

sParamURL = 'https://blue2cast.com/oauthapi/ParametersInfo'
r = session.post(url=sDataURL, headers=sHeaders)
print('--Parameters request--')
print(f'Response status: {r.status_code}')
print(f'Response text  : {r.text}\n')


