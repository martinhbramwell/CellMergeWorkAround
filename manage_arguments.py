#!/usr/bin/python
# -*- coding: utf-8 -*-
#
import argparse

def getArgs() :

    default_sheet_name = 'meta_patches'

#    desc = 'A tool for writing a CSV file into a new/existing sheet in a spreadsheet.\nAn OAuth access token is required.'
    
    desc = 'A tool for getting the row span, column span & / or border style data of one or more sheets in a Google Drive Spreadsheet.\n (Note that for command line access, Google requires client authentication *even if* a document is "Public to anyone on the web").\n\nIf OAuth2 arguments are configured and "-i/--immediate_load" is indicated, then the CSV file will be pushed into the same spreadsheet from which the data was first read.'
    
    
    usage = "usage: %prog [options] arg"
    parser = argparse.ArgumentParser(description=desc)

    # Optional
    parser.add_argument("-il", "--immediate_load", dest="immediate_load"
                      , help="The name for the sheet in the target workbook to be created/replaced with the generated CSV. (Default : Null)")
                      
    # Required
    parser.add_argument("-ssk", "--spreadsheet_key", dest="workbook_key", required=True
                      , help="The key parameter taken from URL of the spreadsheet. (Required.)")
                      
    parser.add_argument("-si", "--sheet_ids", dest="sheet_ids", required=True
                      , help="A comma separated list of sheet identifiers of the single sheet to be accessed, eg 0,1,8,4,3.")
                      
    parser.add_argument("-ue", "--user_email", dest="user_email", required=True, help="This is the GMail account of the spreadsheet owner.")

    parser.add_argument("-ci", "--client_id", dest="client_project", required=True, help="This is the Client ID you get from Google's API console.")
    
    parser.add_argument("-ce", "--client_email", dest="client_email", required=False, help="This is the GMail account that will be used for SMTP communication with spreadsheet owners. (Can't be used with 'service_authentication'")
    
    '''
    group1 = parser.add_mutually_exclusive_group(required=True)
    group1.add_argument("-sa", "--service_authentication", dest="gauth"
                      , help="The 'Auth' parameter returned by Google Client Login. (Can't be used with 'user_id')")
    
#    group1.add_argument("-ui", "--user_id", dest="google_UID", help="Thi is the Google user's email address. )")
                      

    parser.add_argument("-up", "--user_passwd", dest="google_PWD", help="The Google user's password")
    '''

    # Required but default afforded
    parser.add_argument("-st", "--service_type", dest="google_service", default="wise"
                      , help="If you intend to work with something other than a spreadsheet you'll need to change this. (Default : 'wise')")

    parser.add_argument("-sn", "--sheet_name", dest="sheet_name", default=default_sheet_name
                      , help="The name of the sheet to create/replace. (Required. Default : {})".format(default_sheet_name))
    
    # Required, but only the first time                  
    parser.add_argument("-sat", "--smtp_access_token", dest="smtp_access_token", required=False, help="The token you get from Google that allows you to use your own GMail account as an SMTP server, without needing to use your username and password.")
                      
    parser.add_argument("-cs", "--client_secret", dest="client_secret", required=False, help="The Client Secret you get from Google's API console.")

    parser.add_argument("-am", "--auth_method", dest="auth_method", required=False, help="Can be either 'ForDevices', 'InstalledApp' or 'ToBeDropped' (to edit storage to drop the connection between project and user)")
                      
    parser.add_argument("-srt", "--smtp_refresh_token", dest="smtp_refresh_token", required=False, help="This refreshes the access token that refreshes your authorization to use your GMail account as an SMTP server without having to ask you again.")

    parser.add_argument("-ru", "--args.redirect_uri", dest="redirect_uri", required=False, help="This is the Redirect URI you get from Google's API console.")



#    parser.add_argument("CSV_file", help="The CSV format data you want to push out to the spreadsheet.(Required. Default : None.)")

    args = parser.parse_args()
    
       
    now =       NameSpace({
                      'debug': False
                    , 'user_id' : args.user_email
                    , 'client_id' : args.client_project
                })
    user =      {
                   args.user_email: 
                   {
                       args.client_project: 
                       NameSpace({
                           'auth_method': args.auth_method
                       })
                   }
                }
                
    client =    {
                    args.client_project:
                    NameSpace({
                          'client_secret': args.client_secret
                        , 'redirect_uri': args.redirect_uri
                        , 'client_email': args.client_email
                        , 'smtp_access_token': args.smtp_access_token
                        , 'smtp_refresh_token': args.smtp_refresh_token
                    })
                }


    credentials = NameSpace({'gclient': client, 'user' : user, 'now' : now})

    return NameSpace({'args':args, 'credentials': credentials})

class NameSpace(object):
  def __init__(self, adict):
    self.__dict__.update(adict)

