#!/usr/bin/python3
# -*- coding: utf-8 -*-

"""
File: wellnessexport.py
Original author: Tristan (https://github.com/steckbrief)
Date: January 2019

Description:    Use this script to export your fitness data from Garmin Connect.
                See README.md for more information.

Activity & event types:
    https://connect.garmin.com/modern/main/js/properties/event_types/event_types.properties
    https://connect.garmin.com/modern/main/js/properties/activity_types/activity_types.properties
"""

from datetime import datetime, timedelta, date
from getpass import getpass
from os import mkdir, remove, stat
from os.path import isdir, isfile
from subprocess import call
from sys import argv
from xml.dom.minidom import parseString

import argparse
import http.cookiejar
import json
import re
import urllib.error
import urllib.parse
import urllib.request
import zipfile

SCRIPT_VERSION = "2.0.0"
CURRENT_DATE = datetime.now().strftime("%Y-%m-%d")
ACTIVITIES_DIRECTORY = "./" + CURRENT_DATE + "_wellness_export"

PARSER = argparse.ArgumentParser()

# TODO: Implement verbose and/or quiet options.
# PARSER.add_argument('-v', '--verbose', help="increase output verbosity", action="store_true")
PARSER.add_argument("--version", help="print version and exit", action="store_true")
PARSER.add_argument(
    "--username",
    help="your Garmin Connect username (otherwise, you will be prompted)",
    nargs="?",
)
PARSER.add_argument(
    "--password",
    help="your Garmin Connect password (otherwise, you will be prompted)",
    nargs="?",
)
PARSER.add_argument(
    "-d",
    "--directory",
    nargs="?",
    default=ACTIVITIES_DIRECTORY,
    help="the directory to export to (default: './YYYY-MM-DD_garmin_connect_export')",
)
PARSER.add_argument(
    "-u",
    "--unzip",
    help="if downloading ZIP files (format: 'original'), unzip the file and removes the ZIP file",
    action="store_true",
)
PARSER.add_argument(
    "-s",
    "--start",
    help="The date to start to download from",
    nargs="?"
)
PARSER.add_argument(
    "-e",
    "--end",
    help="The end date of the period to download the files",
    nargs="?"
)
PARSER.add_argument(
    "-y",
    "--yesterday",
    help="Will download only the file from yesterday",
    action="store_true"
)

ARGS = PARSER.parse_args()

if ARGS.version:
    print(argv[0] + ", version " + SCRIPT_VERSION)
    exit(0)

COOKIE_JAR = http.cookiejar.CookieJar()
OPENER = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(COOKIE_JAR))
# print(COOKIE_JAR)

def write_to_file(filename, content, mode):
    """Helper function that persists content to file."""
    write_file = open(filename, mode)
    write_file.write(content)
    write_file.close()

# url is a string, post is a dictionary of POST parameters, headers is a dictionary of headers.
def http_req(url, post=None, headers=None):
    """Helper function that makes the HTTP requests."""
    request = urllib.request.Request(url)
    # Tell Garmin we're some supported browser.
    request.add_header(
        "User-Agent",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, \
        like Gecko) Chrome/54.0.2816.0 Safari/537.36",
    )
    if headers:
        for header_key, header_value in headers.items():
            request.add_header(header_key, header_value)
    if post:
        post = urllib.parse.urlencode(post)
        post = post.encode("utf-8")  # Convert dictionary to POST parameter string.
    # print("request.headers: " + str(request.headers) + " COOKIE_JAR: " + str(COOKIE_JAR))
    # print("post: " + str(post) + "request: " + str(request))
    response = OPENER.open((request), data=post)

    if response.getcode() == 204:
        # For activities without GPS coordinates, there is no GPX download (204 = no content).
        # Write an empty file to prevent redownloading it.
        print("Writing empty file since there was no GPX activity data...")
        return ""
    elif response.getcode() != 200:
        raise Exception("Bad return code (" + str(response.getcode()) + ") for: " + url)
    # print(response.getcode())

    return response.read()


print("Welcome to Garmin Connect Exporter!")

# Create directory for data files.
if isdir(ARGS.directory):
    print(
        "Warning: Output directory already exists. Will skip already-downloaded files and \
append to the CSV file."
    )

USERNAME = ARGS.username if ARGS.username else input("Username: ")
PASSWORD = ARGS.password if ARGS.password else getpass()
if not ARGS.yesterday:
    FROM = ARGS.start if ARGS.start else input("Getting data starting with date (YYYY-mm-dd): ")
    TO = ARGS.end if ARGS.end else input("Getting data ending with date (YYYY-mm-dd): ")
else:
    FROM = (date.today() - timedelta(1)).strftime("%Y-%m-%d")
    TO = CURRENT_DATE

WEBHOST = "https://connect.garmin.com"
REDIRECT = "https://connect.garmin.com/post-auth/login"
BASE_URL = "http://connect.garmin.com/en-US/signin"
SSO = "https://sso.garmin.com/sso"
CSS = "https://static.garmincdn.com/com.garmin.connect/ui/css/gauth-custom-v1.2-min.css"

DATA = {
    "service": REDIRECT,
    "webhost": WEBHOST,
    "source": BASE_URL,
    "redirectAfterAccountLoginUrl": REDIRECT,
    "redirectAfterAccountCreationUrl": REDIRECT,
    "gauthHost": SSO,
    "locale": "en_US",
    "id": "gauth-widget",
    "cssUrl": CSS,
    "clientId": "GarminConnect",
    "rememberMeShown": "true",
    "rememberMeChecked": "false",
    "createAccountShown": "true",
    "openCreateAccount": "false",
    "usernameShown": "false",
    "displayNameShown": "false",
    "consumeServiceTicket": "false",
    "initialFocus": "true",
    "embedWidget": "false",
    "generateExtraServiceTicket": "false",
}

print(urllib.parse.urlencode(DATA))

# URLs for various services.
URL_GC_LOGIN = "https://sso.garmin.com/sso/login?" + urllib.parse.urlencode(DATA)
URL_GC_POST_AUTH = "https://connect.garmin.com/modern/activities?"

# Initially, we need to get a valid session cookie, so we pull the login page.
print("Request login page")
http_req(URL_GC_LOGIN)
print("Finish login page")

# Now we'll actually login.
# Fields that are passed in a typical Garmin login.
POST_DATA = {
    "username": USERNAME,
    "password": PASSWORD,
    "embed": "true",
    "lt": "e1s1",
    "_eventId": "submit",
    "displayNameRequired": "false",
}

print("Post login data")
LOGIN_RESPONSE = http_req(URL_GC_LOGIN, POST_DATA).decode()
print("Finish login post")

# extract the ticket from the login response
PATTERN = re.compile(r".*\?ticket=([-\w]+)\";.*", re.MULTILINE | re.DOTALL)
MATCH = PATTERN.match(LOGIN_RESPONSE)
if not MATCH:
    raise Exception(
        "Did not get a ticket in the login response. Cannot log in. Did \
you enter the correct username and password?"
    )
LOGIN_TICKET = MATCH.group(1)
print("Login ticket=" + LOGIN_TICKET)

print("Request authentication URL: " + URL_GC_POST_AUTH + "ticket=" + LOGIN_TICKET)
http_req(URL_GC_POST_AUTH + "ticket=" + LOGIN_TICKET)
print("Finished authentication")

# We should be logged in now.
if not isdir(ARGS.directory):
    mkdir(ARGS.directory)

# This while loop will download data from the server in multiple chunks, if necessary.


def daterange(start_date, end_date):
    for n in range(int ((end_date - start_date).days)):
        yield start_date + timedelta(n)

#start_date = date(2018, 7, 19)
#end_date = date(2019, 1, 8)
start_date = datetime.strptime(FROM, '%Y-%m-%d')
end_date = datetime.strptime(TO, '%Y-%m-%d')
for single_date in daterange(start_date, end_date):
    dtret = single_date.strftime("%Y-%m-%d")
    data_filename = (
        ARGS.directory + "/" + dtret + "_wellness.zip"
    )
    download_url = "https://connect.garmin.com/modern/proxy/download-service/files/wellness/" + dtret
    print(download_url)
    file_mode = "wb"
    print("\tDownloading file...", end=" ")
    try:
        data = http_req(download_url)
    except urllib.error.HTTPError as errs:
        if errs.code == 404:
            # For manual activities (i.e., entered in online without a file upload), there is
            # no original file. # Write an empty file to prevent redownloading it.
            print(
                "Writing empty file since there was no original activity data...",
                end=" ",
            )
            data = ""
        else:
            raise Exception(
                "Failed. Got an unexpected HTTP error ("
                + str(errs.code)
                + download_url
                + ")."
            )


    # Persist file
    write_to_file(data_filename, data, file_mode)
    if ARGS.unzip and data_filename[-3:].lower() == "zip":
        print("Unzipping wellness files...", end=" ")
        print("Filesize is: " + str(stat(data_filename).st_size))
        if stat(data_filename).st_size > 0:
            zip_file = open(data_filename, "rb")
            z = zipfile.ZipFile(zip_file)
            for name in z.namelist():
                z.extract(name, ARGS.directory)
            zip_file.close()
        else:
            print("Skipping 0Kb zip file.")
    else:
        # TODO: Consider validating other formats.
        print("Done.")

print("Done!")
