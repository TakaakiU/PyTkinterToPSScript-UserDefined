import os
# import xml.etree.ElementTree as et
from lxml import etree

from classes.control import ctrlCommon


class ctrlConfig():
    # XMLデータの取得
    # def get_xmldata(filename):
    def get_xmldata(xml_path):
        xmldata = None

        # ファイルパスの取得 ※ Pythonソースに設定ファイルを内包する場合
        # xml_path = ctrlCommon.get_path('classes/config', filename)

        # ファイル存在チェック
        if not (os.path.isfile(xml_path)):
            return None
        # 拡張子チェック
        if not (xml_path.lower().endswith("xml")):
            return None
        # 読み込み
        try:
            parser_enc = etree.XMLParser(encoding='UTF-8', recover=True)
            xmldata = (etree.parse(xml_path, parser=parser_enc)).getroot()
        except OSError:
            print('例外エラー：XML読み込み')
        return xmldata

    # XMLファイルの読み込み
    def read_xmlfile(settings, xmldata):
        if not (xmldata is None):
            settings['hash_algorithm'] = xmldata.find(
                './basicsettings/hash_algorithm').text
            settings['package_maxsize'] = xmldata.find(
                './basicsettings/package_maxsize').text
            settings['package_maxfiles'] = xmldata.find(
                './basicsettings/package_maxfiles').text
            settings['check_maxsize'] = xmldata.find(
                './basicsettings/check_maxsize').text
            settings['check_maxfiles'] = xmldata.find(
                './basicsettings/check_maxfiles').text
            settings['password'] = xmldata.find(
                './basicsettings/password').text
        return settings
