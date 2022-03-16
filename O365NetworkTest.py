#!/usr/bin/python

import json
import sys
import urllib.request

tool_name = 'Office 365 for Mac Network Tester'
tool_version = '1.0'

## Copyright (c) 2019 Microsoft Corp. All rights reserved.
## Scripts are not supported under any Microsoft standard support program or service. The scripts are provided AS IS without warranty of any kind.
## Microsoft disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a
## particular purpose. The entire risk arising out of the use or performance of the scripts and documentation remains with you. In no event shall
## Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever
## (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary
## loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility
## of such damages.
## Feedback: pbowden@microsoft.com

## CONSTANTS
json_endpoint_map = "https://macadmins.software/json/network_endpoints.json"
text_red = '\033[0;31m'
text_yellow = '\033[0;33m'
text_green = '\033[0;32m'
text_blue = '\033[0;34m'
text_normal = '\033[0m'

## FUNCTIONS
def read_rules():
	try:
		response = urllib.request.urlopen(json_endpoint_map)
		json_data = json.loads(response.read())
	except:
		print("Fatal: Unable to load list of endpoints from network")
		sys.exit(1)
	try:
		for key, value in json_data['network'].items():
			if key == 'version':
				sys.stdout.write('\rRules version: {}'.format(value))
				sys.stdout.flush()
			if key == 'last-updated':
				sys.stdout.write('\tLast updated: {}\n'.format(value))
	except:
		print("Fatal: Unable to determine rule version")
		sys.exit(1)
	return json_data

def print_response(code):
	if code == 200 or code == 201:
		sys.stdout.write(text_green + 'OK' + text_normal + '\n')
	else:
	    sys.stdout.write(text_red + code + text_normal + '\n')

def test_url(url, accept, reqtype, content, type, clientid, soapaction):
    if reqtype == 'POST':
        data = content
        data = data.encode('ascii') # data should be bytes
        req = urllib.request.Request(url, data=data)
    elif reqtype == 'OPTIONS':
        data = content
        data = data.encode('ascii') # data should be bytes
        req = urllib.request.Request(url, data=data)
        req.get_method = lambda: 'OPTIONS'
    else:
        req = urllib.request.Request(url)
    if accept:
        req.add_header('Accept', accept)
    if type:
    	req.add_header('Content-Type', type)
    if clientid:
        req.add_header('Client-Id', clientid)
        req.add_header('Content-Encoding', 'gzip')
    if soapaction:
        req.add_header('SOAPAction', soapaction)
    req.add_header('User-Agent', 'Microsoft Office')
    try:
        response = urllib.request.urlopen(req)
        response.getcode()
    except urllib.request.HTTPError as e:
        return str(e.code) + str(' ') + str(e.reason)
    except urllib.request.URLError as e:
        return str(e.code) + str(' ') + str(e.reason)
    else:
        return response.code

## MAIN
print(text_blue + tool_name + ' - ' + tool_version + text_normal)

json_map = read_rules()
for key in json_map['url-entry']:
	sys.stdout.write('\r{0:25s}'.format(key['url-purpose']))
	sys.stdout.write('\t{0:60s}'.format(key['url-friendly']))
	sys.stdout.flush()
	http_resp = test_url(key['url-actual'], key['url-accept'], key['url-request'], key['url-post-content'], key['url-content-type'], key['url-client-id'], key['url-soap-action'])
	print_response(http_resp)

print('')
sys.exit(0)
