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

# TODO <СвТД КодПроисх="643" НомерТД="-" />
# TODO КрНаимСтрПр="Россия"


def index(request):
    if request.method == "POST":
        file = request.FILES['docfile']
        folder = os.path.join(settings.THIS_FOLDER, 'tech/input/')

        now = datetime.datetime.now()
        fs = FileSystemStorage(location=folder)
        filename_split = file.name.split('.')
        file_name = filename_split[0]+'-'+now.strftime("%d%m%y%H%M%S")
        file_ext = filename_split[1]
        filename = file_name+'.'+file_ext

        fs.save(filename, file)

        create_upd(
            function=request.POST.get('function', ''),
            doc_num=request.POST.get('doc_num', ''),
            ul=request.POST.get('ul', ''),
            inn=request.POST.get('inn', ''),
            kpp=request.POST.get('kpp', ''),
            index=request.POST.get('index', ''),
            kodreg=request.POST.get('kodreg', ''),
            city=request.POST.get('city', ''),
            street=request.POST.get('street', ''),
            house=request.POST.get('house', ''),
            korp=request.POST.get('korp', ''),
            phone=request.POST.get('phone', ''),
            schet=request.POST.get('schet', ''),
            bank=request.POST.get('bank', ''),
            bik=request.POST.get('bik', ''),
            korschet=request.POST.get('korschet', ''),
            file_name=str(file_name),
            file_ext=str(file_ext),

        )
        output_filename = 'result'+ now.strftime("%d%m%y%H%M%S")

        file_path = os.path.join(settings.THIS_FOLDER, 'tech/output/') + output_filename + '.xml'
        with open(file_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type="application/force-download")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response

    else:
        return render(request, 'upd/index.html', {})




def download_template(request):

    file_path = os.path.join(settings.THIS_FOLDER, 'tech/templates/xls/template.xlsx')
    with open(file_path, 'rb') as fh:
        response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
        response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
        return response


def create_upd(**data):
    now = datetime.datetime.now()

    upd = Upd(os.path.join(settings.THIS_FOLDER, 'tech/templates/xml/synerdocs/template_ooo.xml'))
    upd.set_file('ON_SCHFDOPPR___20190204_6de6a7f1-9816-4d9b-8973-8dfa68ccfc8c')
    upd.set_document(data['function'], now.strftime("%d.%m.%y"), now.strftime("%H:%M:%S"), data['ul'])
    upd.set_sv_sch_fact(data['doc_num'], now.strftime("%d.%m.%y"))

    upd.set_sv_prod_ul(data['ul'], data['inn'], data['kpp'])
    upd.set_sv_prod_adr(data['index'], data['kodreg'], data['city'], data['street'], data['house'], data['korp'])
    upd.set_sv_prod_phone(data['phone'])
    upd.set_sv_prod_bank(data['schet'], data['bank'], data['bik'], data['korschet'])

    upd.set_gruz_ot_ul(data['ul'], data['inn'], data['kpp'])
    upd.set_gruz_ot_adr(data['index'], data['kodreg'], data['city'], data['street'], data['house'], data['korp'])
    upd.set_gruz_ot_phone(data['phone'])
    upd.set_gruz_ot_bank(data['schet'], data['bank'], data['bik'], data['korschet'])

    upd.set_gruz_pol_ul('ООО «Вайлдберриз», Обособленное подразделение «Подольск»', '7721546864', '503645001')
    upd.set_gruz_pol_adr('142103', '50', 'Подольск', 'Поливановская', '9', '')

    #   не обязательно, согласно гайдам WB
    #    upd.set_gruz_pol_phone('+74957555505')
    #    upd.set_gruz_pol_bank('40702810500110000939', 'ПАО ВТБ', '044525187', '30101810700000000187')

    #   Новосиб
    #    upd.set_gruz_pol_ul('ООО «Вайлдберриз», Обособленное подразделение «Новосибирск-6»', '7721546864', '540345001')
    #    upd.set_gruz_pol_adr('630088', '54', 'Новосибирск', 'Петухова', '71', '')

    #   TODO: Хабаровск
    #    upd.set_gruz_pol_ul('ООО «Вайлдберриз», Обособленное подразделение «Новосибирск-6»', '7721546864', '540345001')
    #    upd.set_gruz_pol_adr('630088', '54', 'Новосибирск', 'Петухова', '71', '')

    #   TODO: Екат
    #    upd.set_gruz_pol_ul('ООО «Вайлдберриз», Обособленное подразделение «Новосибирск-6»', '7721546864', '540345001')
    #    upd.set_gruz_pol_adr('630088', '54', 'Новосибирск', 'Петухова', '71', '')

    #   TODO: Краснодар
    #    upd.set_gruz_pol_ul('ООО «Вайлдберриз», Обособленное подразделение «Новосибирск-6»', '7721546864', '540345001')
    #    upd.set_gruz_pol_adr('630088', '54', 'Новосибирск', 'Петухова', '71', '')

    #   TODO: Питер
    #    upd.set_gruz_pol_ul('ООО «Вайлдберриз», Обособленное подразделение «Новосибирск-6»', '7721546864', '540345001')
    #    upd.set_gruz_pol_adr('630088', '54', 'Новосибирск', 'Петухова', '71', '')

    upd.set_sv_pokup_ul('ООО «Вайлдберриз», Обособленное подразделение «Подольск»"', '7721546864', '503645001')
    upd.set_sv_pokup_adr('142715', '50', '"Ленинский, с/п Развилковское', 'Мильково', 'владение 1')

    upd.set_dop_sv_fhzh1('Российский рубль')
    upd.set_table(data['file_name']+'.'+data['file_ext'])
    upd.output()


def create_upd_sveta():
    tree = etree.parse('upd/templates/synerdocs/template.xml')
    print(tree)

    prefix = 'ON_SCHFDOPPR'
    receiver = 'OperatorServiceCode'
    sender = 'OperatorServiceCode'
    year = '2019'
    month = '02'
    day = '04'
    guid = '2'
    output_file_name = prefix + '_' + receiver + '_' + sender + '_' + year + month + day + '_' + guid

    date_creation = '09.08.2018'
    time_creation = '13.22.35'
    creator = u'ИП Мирзоева Нина Ивановна'
    sf_number = '102'
    sf_date = '04.02.2019'

    svProd_svIP_INNFL = '290120982471'
    svProd_svIP_FIO_surname = u'Мирзоева'
    svProd_svIP_FIO_name = u'Нина'
    svProd_svIP_FIO_patronymic = u'Ивановна'

    # генерация файла

    attr = tree.xpath(u'//Файл')
    attr[0].attrib[u'ИдФайл'] = output_file_name

    attr = tree.xpath(u'//Документ')
    attr[0].attrib[u'ДатаИнфПр'] = date_creation
    attr[0].attrib[u'ВремИнфПр'] = time_creation

    attr = tree.xpath(u'//Документ/СвСчФакт')
    attr[0].attrib[u'НомерСчФ'] = sf_number
    attr[0].attrib[u'ДатаСчФ'] = sf_date

    # ищем начало таблицы с товарами
    book = open_workbook('upd/input/Новая таблица.xlsx')
    sheet = book.sheet_by_index(0)

    start_row = 3

    # парсим таблицу с товарами
    total_row = 0
    total_sum_incl_vat = 0
    total_sum_excl_vat = 0
    total_qty = 0

    parse_table = True
    for rownum in range(start_row, sheet.nrows):
        row = sheet.row_values(rownum)
        if parse_table:
            if row[0] != '':
                # товары
                parent = tree.xpath(u'//Документ/ТаблСчФакт')
                child = etree.SubElement(parent[0], u'СведТов')
                child.set(u'НомСтр', str(int(row[0])))
                child.set(u'НаимТов', row[2])
                child.set(u'ОКЕИ_Тов', u'796')
                child.set(u'КолТов', str(int(row[9])))
                child.set(u'ЦенаТов', str(row[12]))
                child.set(u'СтТовБезНДС', str(row[13]))
                total_sum_excl_vat = total_sum_excl_vat + float(row[13])

                child.set(u'НалСт', u'без НДС')
                child.set(u'СтТовУчНал', str(row[22]))

                print(str(total_sum_incl_vat) + ' + ' + str(float(row[22])))

                total_sum_incl_vat = total_sum_incl_vat + float(row[22])

                child.tail = "\n  "
                child.text = "\n  "

                subchild = etree.SubElement(child, u'Акциз')
                subsubchild = etree.SubElement(subchild, u'БезАкциз')
                subsubchild.text = u'без акциза'
                subchild.tail = "\n  "

                subchild = etree.SubElement(child, u'СумНал')
                subsubchild = etree.SubElement(subchild, u'БезНДС')
                subsubchild.text = u'без НДС'
                subchild.tail = "\n  "

                subchild = etree.SubElement(child, u'ДопСведТов')
                subchild.set(u'ПрТовРаб', u'1')
                subchild.set(u'КодТов', str(int(row[1])))
                subchild.set(u'НаимЕдИзм', u'шт')
                subchild.tail = "\n  "
            else:
                parse_table = False
                total_row = rownum

    # Итого
    row = sheet.row_values(total_row)

    parent = tree.xpath(u'//Документ/ТаблСчФакт')
    child = etree.SubElement(parent[0], u'ВсегоОпл')
    child.set(u'СтТовБезНДСВсего', str(total_sum_excl_vat))
    child.set(u'СтТовУчНалВсего', str(total_sum_incl_vat))
    child.tail = "\n  "

    subchild = etree.SubElement(child, u'СумНалВсего')
    subsubchild = etree.SubElement(subchild, u'СумНал')
    subsubchild.text = '0.0'
    subchild.tail = "\n  "

    tree.write('upd/tech/output/' + output_file_name + '.xml', xml_declaration=True, encoding='windows-1251',
               pretty_print=True)
