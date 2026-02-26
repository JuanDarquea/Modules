#!/usr/bin/env python3 

from urllib.error import HTTPError, URLError
from urllib.request import urlopen, Request

def make_request(url, headers=None, data=None):

    request = Request(url, headers=headers, data=data)
    try:
        with urlopen(request, timeout=10) as response:
            data = response.read()
            print(f'Page reading status: {response.status}')
            print('Page body:')
            print(data.decode('utf-8'))
            print()
            if response.headers == '' or response.headers is None:
                print('No headers to show.')
            else:
                print(f"Page response: {response}")
                print('Page headers:')
                print(response.headers)

            #return data, response
    except HTTPError as error:
        print(error.status, error.reason)
    except URLError as error:
        print(error.reason)
    except TimeoutError:
        print("Request timed out")
