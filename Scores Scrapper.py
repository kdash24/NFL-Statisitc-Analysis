import sys
import os, ssl
import urllib.request
from bs4 import BeautifulSoup, SoupStrainer
import re
import requests
from pprint import pprint
from selenium import webdriver
import xlsxwriter
from openpyxl import load_workbook
import time
import urllib.request
from html_table_parser.parser import HTMLTableParser
import pandas as pd
from datetime import date

ssl._create_default_https_context = ssl._create_unverified_context
response = urllib.request.urlopen('https://www.python.org')
print(response.read().decode('utf-8'))

def url_get_contents(url):
    # Opens a website and read its
    # binary contents (HTTP Response Body)

    #making request to the website
    req = urllib.request.Request(url=url)
    f = urllib.request.urlopen(req)

    #reading contents of the website
    return f.read()

Path=r"F:\Projects\NFL Stats\ "

writer = pd.ExcelWriter(Path + "NFL Stats.xlsx", engine = 'xlsxwriter')


#workbook = xlsxwriter.Workbook(Path + 'NFL Stats Test.xlsx')


#This section creates a dataframe of the offensive PASSING stats for all NFL Teams
NFL_Scoring_url = "https://www.footballdb.com/games/index.html"
xhtml = url_get_contents(NFL_Scoring_url).decode('utf-8')
p = HTMLTableParser()
p.feed(xhtml)

print(p.tables[0])
