import requests
import json
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.client_request import ClientRequest
from office365.runtime.utilities.request_options import RequestOptions
from office365.sharepoint.client_context import ClientContext


def get_excel():
    # Authentication
    ctx_auth = AuthenticationContext(r"https://Domain.sharepoint.com/")

    # Get Access Token
    ctx_auth.acquire_token_for_user("Username@domain.com", "Password")
    ctx = ClientContext(r'Exact URL of the file you want to download,ctx_auth)

    # Initiate Client Request Using Authentication
    request = ClientRequest(ctx_auth)

    # Create Options and create Headers
    options = RequestOptions(ctx.web.resource_url)
    options.set_header('Accept', 'application/json')
    options.set_header('Content-Type', 'application/json')

    # Start Request
    data = request.execute_request_direct(options)

    # get Result content in Json String Format
    myjsondump = json.dumps(data.content.decode('utf-8'))

    myjsonload = json.loads(myjsondump)
    curr_str = ""
    for load in myjsonload:
        curr_str = curr_str + load
    # extract "File Get URL" from json string
    start_index = curr_str.find(r'"FileGetUrl":"') + len(r'"FileGetUrl":"')
    url_text_dump = curr_str[start_index:]
    url_end_Index = url_text_dump.find(r'",')

    # File URL
    my_url = url_text_dump[:url_end_Index]
    my_url = my_url.strip(' \n\t')
    print(my_url)

    # get replace encoded characters
    qurl = my_url.replace(r"\u0026", "&")

    # get request
    resp = requests.get(url=qurl)
    url_data = requests.get(url=qurl)

    # Open an Write in Excel file
    VIP_Excel_File = open("Filename which you want to download.xlsx", mode="wb")
    VIP_Excel_File.write(url_data.content)

    print("Excel Sheet from Sharepoint Extracted Successfully")

get_excel()

