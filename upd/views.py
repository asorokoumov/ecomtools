# -*- coding: utf-8 -*-

from django.shortcuts import render
from lxml import etree
from xlrd import open_workbook
from upd.tools.udp import *
import os
from django.conf import settings
from django.http import *
import datetime
from django.core.files.storage import FileSystemStorage


# Create your views here.


def index(request):
    return 1


def download_template(request):
    return 1


def get_region_code(request):
    return request.POST.get('city_kladr_id', '')[:2]


def get_area(request):
    area = request.POST.get('area', '')
    if area != '':
        return area
    else:
        return False


def get_city(request):
    city = request.POST.get('city', '')
    settlement = request.POST.get('settlement', '')
    if city != '':
        return u'Город', city
    elif settlement != '':
        return u"НаселПункт", settlement
    else:
        return False


def get_street(request):
    street = request.POST.get('street', '')
    if street != '':
        return street
    else:
        return False


def get_house(request):
    house = request.POST.get('house', '')
    if house != '':
        return house
    else:
        return False


def get_building(request):
    buildind = request.POST.get('block', '')
    if buildind != '':
        return buildind
    else:
        return False


def get_address(request):
    address = {}

    address[u'Индекс'] = request.POST.get('postal_code', '')
    address[u'КодРегион'] = get_region_code(request)
    if get_area(request):
        address[u'Район'] = get_area(request)
    if get_city(request):
        city_type, city = get_city(request)
        address[city_type] = city
    if get_street(request):
        address[u'Улица'] = get_street(request)
    if get_house(request):
        address[u'Дом'] = get_house(request)
    if get_building(request):
        address[u'Корпус'] = get_building(request)
    return  address


def get_parse_rules (request):
    parse_rules_name = request.POST.get('account', '')
    print ('--------------------'+(parse_rules_name))
    if parse_rules_name == u'ТТН':
        return {u'Начало страницы': {'fix': 0, 'value': "Код продукции"},
                   u'НаимТов': {'fix': 0, 'column': 'AA'},
                   u'ОКЕИ_Тов': {'fix': 1, 'value': 796},
                   u'КолТов': {'fix': 0, 'column': 'V'},
                   u'ЦенаТов': {'fix': 0, 'column': 'X'},
                   u'СтТовБезНДС': {'fix': 0, 'column': 'BO'},
                   u'СтТовУчНал': {'fix': 0, 'column': 'BO'},
                   u'ПрТовРаб': {'fix': 1, 'value': 1},
                   u'КодТов': {'fix': 0, 'column': 'A'},
                   u'НаимЕдИзм': {'fix': 1, 'value': u'шт'},
                   u'КодПроисх': {'fix': 1, 'value': 643},
                   u'НомерТД': {'fix': 1, 'value': '-'},
                   u'КрНаимСтрПр': {'fix': 1, 'value': u'Россия'},
                   u'НДС': 0,
                   }
    else:
        return {u'Начало страницы': {'fix': 1, 'value': 1},
                   u'НаимТов': {'fix': 0, 'column': 'B'},
                   u'ОКЕИ_Тов': {'fix': 1, 'value': 796},
                   u'КолТов': {'fix': 0, 'column': 'C'},
                   u'ЦенаТов': {'fix': 0, 'column': 'D'},
                   u'СтТовБезНДС': {'fix': 2, 'column': '----'},
                   u'СтТовУчНал': {'fix': 2, 'column': '----'},
                   u'ПрТовРаб': {'fix': 1, 'value': 1},
                   u'КодТов': {'fix': 0, 'column': 'A'},
                   u'НаимЕдИзм': {'fix': 1, 'value': u'шт'},
                   u'КодПроисх': {'fix': 1, 'value': 643},
                   u'НомерТД': {'fix': 1, 'value': '-'},
                   u'КрНаимСтрПр': {'fix': 1, 'value': u'Россия'},
                   u'НДС': 0,
                   }


def get_data_from_form(request):
    print (get_address(request))
    seller = {u"НаимОрг": request.POST.get('party', ''),
              u"ИННЮЛ": request.POST.get('inn', ''),
              u"КПП": request.POST.get('kpp', ''),
              "Тлф": request.POST.get('phone', ''),
              u'НомерСчета': request.POST.get('schet', ''),
              u'Банк': {u'НаимБанк': request.POST.get('bank', ''),
                        u'БИК': request.POST.get('bic', ''),
                        u'КорСчет': request.POST.get('correspondent_account', '')},
              u'Адрес': get_address(request),
              u'ОснованиеПередачи': {u'НаимОсн': request.POST.get('osnovanie', ''),
                                     u'ДатаОсн': request.POST.get('data_osnovaniya', '')}}

    wb = {u'Покупатель': {u'НаимОрг': u"ООО Вайлдберриз", u'ИННЮЛ': "7721546864", u'КПП': "997750001",
                          u'Адрес': {u'Индекс': '142715', u'КодРегион': '50',
                                     u'Район': u"Ленинский, с/п Развилковское",
                                     u'НаселПункт': u"Мильково", u'Дом': "владение 1"}},
          u'Грузополучатель_Подольск': {u'НаимОрг': u"ООО Вайлдберриз, Обособленное подразделение Подольск",
                                        u'ИННЮЛ': "7721546864", u'КПП': "503645001", "Тлф": "8(495)775-55-05",
                                        u'Адрес': {u'Индекс': '142103', u'КодРегион': '50',
                                                   u'Город': u"Подольск", u'Улица': u"Поливановская", u'Дом': "9"}}
          }

    file = request.FILES['docfile']

    now = datetime.datetime.now()
    filename_split = file.name.split('.')
    file_name = filename_split[0]
    file_ext = filename_split[1]
    filename = file_name + '.' + file_ext

    sf_number = request.POST.get('sf_number', '')
    parse_rules = get_parse_rules(request)

    return seller, wb, filename, file, sf_number, parse_rules


def save_file_to_input_folder(input_folder, file, filename):
    folder = os.path.join(settings.THIS_FOLDER, input_folder)
    fs = FileSystemStorage(location=folder)
    fs.save(filename, file)


def index(request):
    if request.method == "POST":

        seller, wb, filename, file, sf_number, parse_rules = get_data_from_form(request)

        save_file_to_input_folder(input_folder='tech/input/', file=file, filename=filename)

        create_upd(seller=seller, wb=wb, filename=filename, sf_number=sf_number, parse_rules=parse_rules)

        now = datetime.datetime.now()
        output_filename = filename+'_result_' + now.strftime("%d%m%y%H%M%S")

        file_path = os.path.join(settings.THIS_FOLDER, 'tech/output/') + output_filename + '.xml'
        with open(file_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type="application/force-download")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response

    else:
        return render(request, 'upd/wizard.html', {})


'''
def download_template(request):
    file_path = os.path.join(settings.THIS_FOLDER, 'tech/templates/xls/template.xlsx')
    with open(file_path, 'rb') as fh:
        response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
        response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
        return response
'''


def create_upd(seller, wb, filename, sf_number, parse_rules):
    upd = Upd(
        seller=seller, wb=wb,
        sf_number=sf_number, filename=filename, parse_rules=parse_rules)
    upd.save_file_to_output_folder()

'''
def test(request):
    # ООО
    seller = {u"НаимОрг": u'ООО Форси', u"ИННЮЛ": "5904650273", u"КПП": "590401001", "Тлф": "+79024786088",
              u'НомерСчета': '40702810401280002173',
              u'Банк': {u'НаимБанк': u"Филиал ОАО Уралсиб в г.Уфа", u'БИК': "048073770",
                        u'КорСчет': "30101810600000000770"},
              u'Адрес': {u'Индекс': '614000', u'КодРегион': '59', u'Город': u"Пермь", u'Улица': u"Чкалова", u'Дом': "9",
                         u'Корпус': u"Д"},
              u'Основание': {u'НаимОсн': 'Агентский договор №49', u'ДатаОсн': '15.03.2018'}}

    wb = {u'Покупатель': {u'НаимОрг': u"ООО Вайлдберриз", u'ИННЮЛ': "7721546864", u'КПП': "997750001",
                          u'Адрес': {u'Индекс': '142715', u'КодРегион': '50',
                                     u'Район': u"Ленинский, с/п Развилковское",
                                     u'НаселПункт': u"Мильково", u'Дом': "владение 1"}},
          u'Грузополучатель_Подольск': {u'НаимОрг': u"ООО Вайлдберриз, Обособленное подразделение Подольск",
                                        u'ИННЮЛ': "7721546864", u'КПП': "503645001", "Тлф": "8(495)775-55-05",
                                        u'Адрес': {u'Индекс': '142103', u'КодРегион': '50',
                                                   u'Город': u"Подольск", u'Улица': u"Поливановская", u'Дом': "9"}}
          }
    filename = u'ТТН 358.xls'
    parse_rules = {u'Начало страницы': {'fix': 0, 'value': "Код продукции"},
                   u'НаимТов': {'fix': 0, 'column': 'AA'},
                   u'ОКЕИ_Тов': {'fix': 1, 'value': 796},
                   u'КолТов': {'fix': 0, 'column': 'V'},
                   u'ЦенаТов': {'fix': 0, 'column': 'X'},
                   u'СтТовБезНДС': {'fix': 0, 'column': 'BO'},
                   u'СтТовУчНал': {'fix': 0, 'column': 'BO'},
                   u'ПрТовРаб': {'fix': 1, 'value': 1},
                   u'КодТов': {'fix': 0, 'column': 'A'},
                   u'НаимЕдИзм': {'fix': 1, 'value': u'шт'},
                   u'КодПроисх': {'fix': 1, 'value': 643},
                   u'НомерТД': {'fix': 1, 'value': '-'},
                   u'КрНаимСтрПр': {'fix': 1, 'value': u'Россия'},
                   u'НДС': 0,
                   }
    upd = Upd(
        seller=seller, wb=wb,
        sf_number='11', filename=filename, parse_rules=parse_rules)
    upd.save_file_to_output_folder()
    return render(request, 'upd/index.html', {})
'''
'''
def test2(request):
    # ООО
    seller = {u"НаимОрг": u'ООО КОСТЬЕРА ФЕШН', u"ИННЮЛ": "7725482499", u"КПП": "772501001", "Тлф": "+79857840821",
              u'НомерСчета': '40702810810000317057',
              u'Банк': {u'НаимБанк': u"Банк АО 'ТИНЬКОФФ БАНК'", u'БИК': "044525974",
                        u'КорСчет': "30101810145250000974"},
              u'Адрес': {u'Индекс': '614000', u'КодРегион': '59', u'Город': u"Москва", u'Улица': u"Шаболовка",
                         u'Дом': "34",
                         u'Корпус': u"5"},
              u'ОснованиеПередачи': {u'НаимОсн': 'Агентский договор №123', u'ДатаОсн': '11.11.2018'}}

    wb = {u'Покупатель': {u'НаимОрг': u"ООО Вайлдберриз", u'ИННЮЛ': "7721546864", u'КПП': "997750001",
                          u'Адрес': {u'Индекс': '142715', u'КодРегион': '50',
                                     u'Район': u"Ленинский, с/п Развилковское",
                                     u'НаселПункт': u"Мильково", u'Дом': "владение 1"}},
          u'Грузополучатель_Подольск': {u'НаимОрг': u"ООО Вайлдберриз, Обособленное подразделение Подольск",
                                        u'ИННЮЛ': "7721546864", u'КПП': "503645001", "Тлф": "8(495)775-55-05",
                                        u'Адрес': {u'Индекс': '142103', u'КодРегион': '50',
                                                   u'Город': u"Подольск", u'Улица': u"Поливановская", u'Дом': "9"}}
          }
    filename = u'Ecom template.xlsx'
    parse_rules = {u'Начало страницы': {'fix': 1, 'value': 1},
                   u'НаимТов': {'fix': 0, 'column': 'B'},
                   u'ОКЕИ_Тов': {'fix': 1, 'value': 796},
                   u'КолТов': {'fix': 0, 'column': 'C'},
                   u'ЦенаТов': {'fix': 0, 'column': 'D'},
                   u'СтТовБезНДС': {'fix': 2, 'column': '----'},
                   u'СтТовУчНал': {'fix': 2, 'column': '----'},
                   u'ПрТовРаб': {'fix': 1, 'value': 1},
                   u'КодТов': {'fix': 0, 'column': 'A'},
                   u'НаимЕдИзм': {'fix': 1, 'value': u'шт'},
                   u'КодПроисх': {'fix': 1, 'value': 643},
                   u'НомерТД': {'fix': 1, 'value': '-'},
                   u'КрНаимСтрПр': {'fix': 1, 'value': u'Россия'},
                   u'НДС': 0,
                   }
    upd = Upd(
        seller=seller, wb=wb,
        sf_number='11', filename=filename, parse_rules=parse_rules)
    upd.save_file_to_output_folder()
    return render(request, 'upd/index.html', {})
'''