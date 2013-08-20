#!/usr/bin/python
#

import urllib, urllib2

parm_key = 'Auth='

data = {}
data['accountType'] = 'GOOGLE'

def getGoogleAuthorizationKey(service, email, password):

    data['Email'] = email
    data['Passwd'] = password
    data['service'] = service

    request = urllib2.Request(
              "https://www.google.com/accounts/ClientLogin"
            , urllib.urlencode(data)
        )
        
    for line in urllib2.urlopen(request).readlines():
        if line.startswith(parm_key):
            return line[len(parm_key):]


if __name__ == "__main__":
    import sys
    print 'Your authorization key is :\n {}'.format(getGoogleAuthorizationKey(sys.argv[1], sys.argv[2], sys.argv[3]))
    

