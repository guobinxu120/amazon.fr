# -*- coding: utf-8 -*-
from scrapy import Spider, Request

import sys
import re, os, requests, urllib
from scrapy.utils.response import open_in_browser
from collections import OrderedDict
import time
from xlrd import open_workbook
from shutil import copyfile
import json, re, csv
from scrapy.http import FormRequest
from scrapy.http import TextResponse

def readExcel(path):
    wb = open_workbook(path)
    result = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        herders = []
        for row in range(0, number_of_rows):
            values = OrderedDict()
            for col in range(number_of_columns):
                value = (sheet.cell(row,col).value)
                if row == 0:
                    herders.append(value)
                else:

                    values[herders[col]] = value
            if len(values.values()) > 0:
                result.append(values)
        break
    return result


class AngelSpider(Spider):
    name = "amazon_fr"
    start_urls = 'https://www.amazon.fr/'
    count = 0
    use_selenium = False
    ean_codes = readExcel("sample_EAN.xlsx")
    models = []
    headers = ['EAN', 'shot description']
    for code in ean_codes:
        item = OrderedDict()
        try:
            item['EAN'] = str(int(code['EAN']))
        except:
            item['EAN'] = str(code['EAN'])
 
        models.append(item)

    def start_requests(self):
        for i, val in enumerate(self.models):
            ern_code = val['EAN']
            url ='https://www.amazon.fr/s/ref=nb_sb_noss?__mk_fr_FR=%C3%85M%C3%85%C5%BD%C3%95%C3%91&url=search-alias%3Daps&field-keywords='
            yield Request(url + ern_code, callback=self.parse1, meta={'order_num':i})

    def parse1(self, respond):
        url = respond.xpath('//div[@id="atfResults"]//a[@class="a-link-normal a-text-normal"]/@href').extract_first()
        if url:
            yield Request(url, callback=self.parse, meta={'order_num':respond.meta['order_num']})

    def parse(self, response):
        order_num = response.meta['order_num']
        item = self.models[order_num]

        # item = OrderedDict()
        # item['EAN'] = ear_code
        # item['shot description'] = ''
        # item['Marque'] = ''
        # item['Numéro de modèle'] = ''
        # item['Couleur'] = ''
        # item['Poids de l\'article'] = ''
        # item['Dimensions du produit (L x l x h)'] = ''
        # item['Puissance'] = ''
        # item['Voltage'] = ''
        # item['Fonction arrêt automatique'] = ''
        # item['Classe d\'énergie'] = ''
        list = response.xpath('//ul[@class="a-unordered-list a-vertical a-spacing-none"]/li/span/text()').extract()
        dest = ''
        for str in list:
            dest = dest + '\n' + str.strip()
        item['shot description'] = dest.encode("utf-8", "ignore").strip()
        if item['shot description']:
            item['shot description'] = item['shot description'].strip()
        
        keys = response.xpath('//div[@class="column col1 "]//td[@class="label"]//text()').extract()
        values = response.xpath('//div[@class="column col1 "]//td[@class="value"]/text()').extract()
        for i, key in enumerate(keys):
            if not key.lower().strip() in self.headers:
                self.headers.append(key.lower().strip())
            item[key.lower().strip()] = values[i].encode("utf-8", "ignore")
            

        self.models[order_num] = item
        yield item
        # item['Marque'] = response.xpath('//a[@id="brand"]/text()').extract_first()
