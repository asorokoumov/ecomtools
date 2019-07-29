# -*- coding: utf-8 -*-

from lxml import etree
from xlrd import open_workbook
from django.conf import settings
import os
import datetime
from decimal import Decimal


class Upd:
    """docstring"""

    def __init__(self, seller, sf_number, wb, filename, parse_rules):
        """Constructor"""
        self.seller = seller
        self.wb = wb
        self.sf_number = sf_number
        self.table = {}
        self.filename = filename
        self.parse_rooles = parse_rules

        page = etree.Element(u'Файл')
        self.tree = etree.ElementTree(page)
        self.set_header()
        self.parse_file()
        self.set_table()
        self.set_footer()


        pass

    def set_header(self):
        self.set_file()
        self.set_sv_uch_doc_obor()
        self.set_document()

        self.set_sv_sch_fact()
        self.set_sv_prod_ul()
        self.set_sv_prod_adr()
        self.set_sv_prod_phone()
        self.set_sv_prod_bank()

        self.set_gruz_ot_ul()
        self.set_gruz_ot_adr()
        self.set_gruz_ot_phone()
        self.set_gruz_ot_bank()

        self.set_gruz_pol_ul()
        self.set_gruz_pol_adr()
        self.set_gruz_pol_phone()

        self.set_sv_pokup_ul()
        self.set_sv_pokup_adr()
    def set_footer(self):
        self.set_sv_prod_per()

    # Файл
    def set_file(self):
        # <Файл xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ИдФайл="ON_SCHFDOPPR___20190204_6de6a7f1-9816-4d9b-8973-8dfa68ccfc8c" ВерсФорм="5.01" ВерсПрог="1.0">
        self.tree.xpath(u'//Файл')[0].attrib[u'ИдФайл'] = "WB_UPD"
        self.tree.xpath(u'//Файл')[0].attrib[u'ВерсФорм'] = "5.01"

    # СвУчДокОбор
    def set_sv_uch_doc_obor(self):
        # <СвУчДокОбор> <СвОЭДОтпр НаимОрг="synerdocs" ИННЮЛ="7728075928" ИдЭДО="2TS" /> </СвУчДокОбор>
        parent = self.tree.xpath(u'//Файл')
        child = etree.SubElement(parent[0], u'СвУчДокОбор')
        child.set(u'ИдОтпр', '00000')
        child.set(u'ИдПол', '00000')

        sub_child = etree.SubElement(child, u'СвОЭДОтпр')
        sub_child.set(u'НаимОрг', 'synerdocs')
        sub_child.set(u'ИННЮЛ', '7728075928')
        sub_child.set(u'ИдЭДО', '2TS')


    # Документ
    def set_document(self):
        # <Документ КНД="1115125" Функция="ДОП" ПоФактХЖ="Документ об отгрузке товаров (выполнении работ), передаче имущественных прав (документ об оказании услуг)" НаимДокОпр="Документ об отгрузке товаров (выполнении работ), передаче имущественных прав (Документ об оказании услуг)" ДатаИнфПр="09.08.2018" ВремИнфПр="13.22.35" НаимЭконСубСост="ООО &quot;Форси&quot;">
        parent = self.tree.xpath(u'//Файл')
        child = etree.SubElement(parent[0], u'Документ')
        child.set(u'КНД', u'1115125')
        child.set(u'Функция', u'ДОП')
        now = datetime.datetime.now()
        child.set(u'ПоФактХЖ', u'.')
        child.set(u'НаимДокОпр', u'Документ об отгрузке товаров')
        child.set(u'ДатаИнфПр', now.strftime("%d.%m.%Y"))
        child.set(u'ВремИнфПр', now.strftime("%H.%M.%S"))
        child.set(u'НаимЭконСубСост', self.seller.get(u"НаимОрг", ''))

        # необязательные поля
        # child.set(u'ПоФактХЖ', u'Документ об отгрузке товаров (выполнении работ), передаче имущественных прав (документ об оказании услуг)')
        # child.set(u'НаимДокОпр', u'Документ об отгрузке товаров (выполнении работ), передаче имущественных прав (Документ об оказании услуг)')
        # child.set(u'ОснДоверОргСост', u'????')

    # СвСчФакт
    def set_sv_sch_fact(self):
        # <СвСчФакт НомерСчФ="48" ДатаСчФ="07.02.2019" КодОКВ="643">

        parent = self.tree.xpath(u'//Документ')
        child = etree.SubElement(parent[0], u'СвСчФакт')
        child.set(u'НомерСчФ', self.sf_number)
        now = datetime.datetime.now()
        child.set(u'ДатаСчФ', now.strftime("%d.%m.%Y"))
        child.set(u'КодОКВ', '643')

    # СвЮЛУч
    def set_sv_prod_ul(self):
        # <СвЮЛУч НаимОрг="ООО &quot;ФОРСИ&quot;" ИННЮЛ="5904650273" КПП="590401001" />
        parent = self.tree.xpath(u'//СвСчФакт')
        child = etree.SubElement(parent[0], u'СвПрод')
        sub_child1 = etree.SubElement(child, u'ИдСв')
        sub_child2 = etree.SubElement(sub_child1, u'СвЮЛУч')

        sub_child2.set(u'НаимОрг', self.seller.get("НаимОрг", ''))
        sub_child2.set(u'ИННЮЛ', self.seller.get("ИННЮЛ", ''))
        sub_child2.set(u'КПП', self.seller.get("КПП", ''))

    # СвПрод/АдрРФ
    def set_sv_prod_adr(self):
        # <АдрРФ Индекс="614064" КодРегион="59" Город="Пермь" Улица="Чкалова" Дом="9" Корпус="Д" />
        parent = self.tree.xpath(u'//СвПрод')
        child = etree.SubElement(parent[0], u'Адрес')
        sub_child = etree.SubElement(child, u'АдрРФ')

        address = self.seller.get('Адрес', {})

        sub_child.set(u'Индекс', address.get(u'Индекс', ''))
        sub_child.set(u'КодРегион', address.get(u'КодРегион', ''))
        sub_child.set(u'Город', address.get(u'Город', ''))
        sub_child.set(u'Улица', address.get(u'Улица', ''))
        sub_child.set(u'Дом', address.get(u'Дом', ''))
        sub_child.set(u'Корпус', address.get(u'Корпус', ''))

    # Контакт/Тлф
    def set_sv_prod_phone(self):
        # <Контакт Тлф="+79024786088" />
        parent = self.tree.xpath(u'//СвПрод')
        child = etree.SubElement(parent[0], u'Контакт')
        child.set(u'Тлф', self.seller.get(u'Тлф', ''))

    # БанкРекв/СвБанк
    def set_sv_prod_bank(self):
        # <СвБанк НаимБанк="Филиал ОАО &quot;Уралсиб&quot; в г.Уфа" БИК="048073770" КорСчет="30101810600000000770" />
        parent = self.tree.xpath(u'//СвПрод')
        child = etree.SubElement(parent[0], u'БанкРекв')
        sub_child = etree.SubElement(child, u'СвБанк')

        child.set(u'НомерСчета', self.seller.get(u'НомерСчета', ''))

        bank = self.seller.get('Банк', {})
        sub_child.set(u'НаимБанк', bank.get(u'НаимБанк', ''))
        sub_child.set(u'БИК', bank.get(u'БИК', ''))
        sub_child.set(u'КорСчет', bank.get(u'КорСчет', ''))

    # /СвСчФакт/ГрузОт/ГрузОтпр/ИдСв/СвЮЛУч
    def set_gruz_ot_ul(self):
        # <СвЮЛУч НаимОрг="ООО &quot;ФОРСИ&quot;" ИННЮЛ="5904650273" КПП="590401001" />
        parent = self.tree.xpath(u'//СвСчФакт')
        child = etree.SubElement(parent[0], u'ГрузОт')
        sub_child = etree.SubElement(child, u'ГрузОтпр')
        sub_sub_child = etree.SubElement(sub_child, u'ИдСв')
        sub_sub_sub_child = etree.SubElement(sub_sub_child, u'СвЮЛУч')

        sub_sub_sub_child.set(u'НаимОрг', self.seller.get(u"НаимОрг", ''))
        sub_sub_sub_child.set(u'ИННЮЛ', self.seller.get(u"ИННЮЛ", ''))
        sub_sub_sub_child.set(u'КПП', self.seller.get(u"КПП", ''))

    # СвСчФакт/ГрузОт/ГрузОтпр/Адрес/АдрРФ
    def set_gruz_ot_adr(self):
        # <АдрРФ Индекс="614064" КодРегион="59" Город="Пермь" Улица="Чкалова" Дом="9" Корпус="Д" />

        parent = self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр')
        child = etree.SubElement(parent[0], u'Адрес')
        sub_child = etree.SubElement(child, u'АдрРФ')

        address = self.seller.get('Адрес', {})

        sub_child.set(u'Индекс', address.get(u'Индекс', ''))
        sub_child.set(u'КодРегион', address.get(u'КодРегион', ''))
        sub_child.set(u'Город', address.get(u'Город', ''))
        sub_child.set(u'Улица', address.get(u'Улица', ''))
        sub_child.set(u'Дом', address.get(u'Дом', ''))
        sub_child.set(u'Корпус', address.get(u'Корпус', ''))

    # /СвСчФакт/ГрузОт/ГрузОтпр/Контакт
    def set_gruz_ot_phone(self):
        # <Контакт Тлф="+79024786088" />
        parent = self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр')
        child = etree.SubElement(parent[0], u'Контакт')
        child.set(u'Тлф', self.seller.get(u'Тлф', ''))

    # /СвСчФакт/ГрузОт/ГрузОтпр/БанкРекв/СвБанк
    def set_gruz_ot_bank(self):
        # <СвБанк НаимБанк="Филиал ОАО &quot;Уралсиб&quot; в г.Уфа" БИК="048073770" КорСчет="30101810600000000770" />

        parent = self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр')
        child = etree.SubElement(parent[0], u'БанкРекв')
        sub_child = etree.SubElement(child, u'СвБанк')

        child.set(u'НомерСчета', self.seller.get('НомерСчета', ''))

        bank = self.seller.get(u'Банк', {})
        sub_child.set(u'НаимБанк', bank.get(u'НаимБанк', ''))
        sub_child.set(u'БИК', bank.get(u'БИК', ''))
        sub_child.set(u'КорСчет', bank.get(u'КорСчет', ''))

    # /СвСчФакт/ГрузПолуч/ИдСв/СвЮЛУч
    def set_gruz_pol_ul(self):
        # <СвЮЛУч НаимОрг="ООО &quot;ФОРСИ&quot;" ИННЮЛ="5904650273" КПП="590401001" />

        parent = self.tree.xpath(u'//СвСчФакт')
        child = etree.SubElement(parent[0], u'ГрузПолуч')
        sub_child1 = etree.SubElement(child, u'ИдСв')
        sub_child2 = etree.SubElement(sub_child1, u'СвЮЛУч')

        receiver = self.wb.get(u'Грузополучатель_Подольск', {})

        sub_child2.set(u'НаимОрг', receiver.get(u"НаимОрг", ''))
        sub_child2.set(u'ИННЮЛ', receiver.get(u"ИННЮЛ", ''))
        sub_child2.set(u'КПП', receiver.get(u"КПП", ''))

    # /СвСчФакт/ГрузПолуч/ИдСв/СвЮЛУч/Адрес/АдрРФ
    def set_gruz_pol_adr(self):
        # <АдрРФ Индекс="614064" КодРегион="59" Город="Пермь" Улица="Чкалова" Дом="9" Корпус="Д" />

        parent = self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч')
        child = etree.SubElement(parent[0], u'Адрес')
        sub_child = etree.SubElement(child, u'АдрРФ')

        receiver = self.wb.get(u'Грузополучатель_Подольск', {})
        address = receiver.get(u'Адрес', {})

        sub_child.set(u'Индекс', address.get(u'Индекс', ''))
        sub_child.set(u'КодРегион', address.get(u'КодРегион', ''))
        sub_child.set(u'Город', address.get(u'Город', ''))
        sub_child.set(u'Улица', address.get(u'Улица', ''))
        sub_child.set(u'Дом', address.get(u'Дом', ''))

    # //Документ/СвСчФакт/ГрузПолуч/Контакт
    def set_gruz_pol_phone(self):
        # <Контакт Тлф="+79024786088" />

        parent = self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч')
        child = etree.SubElement(parent[0], u'Контакт')

        receiver = self.wb.get('Грузополучатель_Подольск', {})
        child.set(u'Тлф', receiver.get(u'Тлф', ''))

    # //Документ/СвСчФакт/СвПокуп/ИдСв/СвЮЛУч
    def set_sv_pokup_ul(self):
        # <СвЮЛУч НаимОрг="ООО «Вайлдберриз»" ИННЮЛ="7721546864" КПП="997750001" />

        parent = self.tree.xpath(u'//СвСчФакт')
        child = etree.SubElement(parent[0], u'СвПокуп')
        sub_child1 = etree.SubElement(child, u'ИдСв')
        sub_child2 = etree.SubElement(sub_child1, u'СвЮЛУч')

        buyer = self.wb.get(u'Покупатель', {})

        sub_child2.set(u'НаимОрг', buyer.get(u"НаимОрг", ''))
        sub_child2.set(u'ИННЮЛ', buyer.get(u"ИННЮЛ", ''))
        sub_child2.set(u'КПП', buyer.get(u"КПП", ''))

    # //Документ/СвСчФакт/СвПокуп/Адрес/АдрРФ
    def set_sv_pokup_adr(self):
        # <АдрРФ Индекс="142715" КодРегион="50" Район="Ленинский, с/п Развилковское" НаселПункт="Мильково" Дом="владение 1" />

        parent = self.tree.xpath(u'//Документ/СвСчФакт/СвПокуп')
        child = etree.SubElement(parent[0], u'Адрес')
        sub_child = etree.SubElement(child, u'АдрРФ')

        buyer = self.wb.get('Покупатель', {})
        address = buyer.get('Адрес', {})

        sub_child.set(u'Индекс', address.get(u'Индекс', ''))
        sub_child.set(u'КодРегион', address.get(u'КодРегион', ''))
        sub_child.set(u'Район', address.get(u'Район', ''))
        sub_child.set(u'НаселПункт', address.get(u'НаселПункт', ''))
        sub_child.set(u'Дом', address.get(u'Дом', ''))

    #  def set_dop_sv_fhzh1(self, naim_okv):
    # <ДопСвФХЖ1 НаимОКВ="Российский рубль" />
    #     self.tree.xpath(u'//Документ/СвСчФакт/ДопСвФХЖ1')[0].attrib[u'НаимОКВ'] = naim_okv

    def set_sv_prod_per(self):
        parent = self.tree.xpath(u'//Документ')
        child = etree.SubElement(parent[0], u'СвПродПер')
        sub_child = etree.SubElement(child, u'СвПер')
        sub_child.set(u'СодОпер', u'Отгрузка товара')

        #TODO дата, номер вынести в настройки
        sub_sub_child = etree.SubElement(sub_child, u'ОснПер')
        osnovanie_peredachi = self.seller.get(u'ОснованиеПередачи', {})
        sub_sub_child.set(u'НаимОсн', osnovanie_peredachi.get(u'НаимОсн', ''))
        sub_sub_child.set(u'ДатаОсн', osnovanie_peredachi.get(u'ДатаОсн', ''))

    def set_item_value (self, name, row):
        if self.parse_rooles.get(name, '').get(u'fix', '') == 0:
            column_num = self.column_num(self.parse_rooles.get(name, '').get(u'column', ''))
            return str(row[column_num])
        else:
            return self.parse_rooles.get(name, '').get(u'value', '')

    def parse_file(self):
        book = open_workbook(os.path.join(settings.THIS_FOLDER, 'tech/input/' + self.filename))
        sheet = book.sheet_by_index(0)
        items = []
        num_str = 1
        total_sum_incl_vat = 0
        total_sum_excl_vat = 0
        page_started = False
        page_started_string = ''

        if self.parse_rooles.get(u'Начало страницы', '').get(u'fix', '') == 0:
            started_line = False
            first_row = 0
            page_started_string = self.parse_rooles.get(u'Начало страницы', '').get(u'value', '')
        else:
            started_line = True
            first_row = int(self.parse_rooles.get(u'Начало страницы', '').get(u'value', ''))

        for rownum in range(first_row, sheet.nrows):
            row = sheet.row_values(rownum)

            if row[0] in (None, ""):
                    page_started = False

            if started_line:
                if not page_started:
                    if row[0] not in (None, ""):
                        page_started = True



            print ('.'+str(row[0]) + '.   ' + str(rownum) + '  ' + str(page_started))
            if page_started:
                item = {}

                item[u'НомСтр'] = str(num_str)
                item[u'НаимТов'] = str(self.set_item_value(name=u'НаимТов', row=row))
                item[u'ОКЕИ_Тов'] = str(self.set_item_value(name=u'ОКЕИ_Тов', row=row))
                item[u'КолТов'] = str(self.set_item_value(name=u'КолТов', row=row))
                item[u'ЦенаТов'] = str(self.set_item_value(name=u'ЦенаТов', row=row).replace(u'\xa0', '').
                                       replace(u' ', '').replace(u',', '.'))
                if self.parse_rooles.get('СтТовБезНДС', '').get(u'fix', '') == 2:
                    item[u'СтТовБезНДС'] = str(float(item[u'КолТов'])*float(item[u'ЦенаТов']))
                else:
                    item[u'СтТовБезНДС'] = str(self.set_item_value(name=u'СтТовБезНДС', row=row).replace(u'\xa0', '').
                                               replace(u' ', '').replace(u',', '.'))
                #todo Без НДС
                item[u'НалСт'] = u'без НДС'
                if self.parse_rooles.get('СтТовУчНал', '').get(u'fix', '') == 2:
                    item[u'СтТовУчНал'] = str(float(item[u'КолТов'])*float(item[u'ЦенаТов']))

                else:
                    item[u'СтТовУчНал'] = str(self.set_item_value(name=u'СтТовУчНал', row=row).replace(u'\xa0', '').
                                              replace(u' ', '').replace(u',', '.'))

                item[u'ПрТовРаб'] = str(self.set_item_value(name=u'ПрТовРаб', row=row))
                item[u'КодТов'] = str(self.set_item_value(name=u'КодТов', row=row))
                item[u'НаимЕдИзм'] = str(self.set_item_value(name=u'НаимЕдИзм', row=row))
                item[u'КодПроисх'] = str(self.set_item_value(name=u'КодПроисх', row=row))
                item[u'НомерТД'] = str(self.set_item_value(name=u'НомерТД', row=row))
                item[u'КрНаимСтрПр'] = str(self.set_item_value(name=u'КрНаимСтрПр', row=row))
                items.append(item)
                num_str = num_str + 1
                total_sum_excl_vat = total_sum_excl_vat + Decimal(item[u'СтТовБезНДС'])
                total_sum_incl_vat = total_sum_incl_vat + Decimal(item[u'СтТовУчНал'])

            if not started_line:
                if not page_started:
                    if page_started_string in row[0]:
                        page_started = True


        total = {}
        total[u'СтТовБезНДСВсего'] = str(total_sum_excl_vat)
        total[u'СтТовУчНалВсего'] = str(total_sum_incl_vat)

        self.table[u'Товары'] = items
        self.table[u'Итог'] = total

    def set_table(self):
        items = self.table.get(u'Товары', {})
        total = self.table.get(u'Итог', {})
        parent = self.tree.xpath(u'//Документ')
        child0 = etree.SubElement(parent[0], u'ТаблСчФакт')

        for item in items:

            child = etree.SubElement(child0, u'СведТов')
            child.set(u'НомСтр', item.get(u'НомСтр', ''))
            child.set(u'НаимТов', item.get(u'НаимТов', ''))
            child.set(u'ОКЕИ_Тов', item.get(u'ОКЕИ_Тов', ''))
            child.set(u'КолТов', item.get(u'КолТов', ''))
            child.set(u'ЦенаТов', item.get(u'ЦенаТов', ''))
            child.set(u'СтТовБезНДС', item.get(u'СтТовБезНДС', ''))
            child.set(u'СтТовУчНал', item.get(u'СтТовУчНал', ''))
            child.set(u'НалСт', item.get(u'НалСт', ''))

            sub_child1 = etree.SubElement(child, u'Акциз')
            sub_sub_child1 = etree.SubElement(sub_child1, u'БезАкциз')
            sub_sub_child1.text = u'без акциза'

            sub_child2 = etree.SubElement(child, u'СумНал')
            sub_sub_child2 = etree.SubElement(sub_child2, u'БезНДС')
            sub_sub_child2.text = u'без НДС'

            sub_child4 = etree.SubElement(child, u'СвТД')
            sub_child4.set(u'КодПроисх', item.get('КодПроисх', ''))
            sub_child4.set(u'НомерТД', item.get('НомерТД', ''))

            sub_child3 = etree.SubElement(child, u'ДопСведТов')
            sub_child3.set(u'ПрТовРаб', item.get('ПрТовРаб', ''))
            sub_child3.set(u'КодТов', item.get('КодТов', ''))
            sub_child3.set(u'НаимЕдИзм', item.get('НаимЕдИзм', ''))
            sub_child3.set(u'КрНаимСтрПр', item.get('КрНаимСтрПр', ''))


        child2 = etree.SubElement(child0, u'ВсегоОпл')
        child2.set(u'СтТовБезНДСВсего', total.get('СтТовБезНДСВсего', ''))
        child2.set(u'СтТовУчНалВсего', total.get('СтТовУчНалВсего', ''))

        sub_child5 = etree.SubElement(child2, u'СумНалВсего')
        sub_sub_child5 = etree.SubElement(sub_child5, u'СумНал')
        sub_sub_child5.text = u'0.0'



    #TODO страна производитель


    def set_table_old(self, filename):
        book = open_workbook(os.path.join(settings.THIS_FOLDER, 'tech/input/' + filename))
        sheet = book.sheet_by_index(0)

        # парсим таблицу с товарами
        start_row = 2
        total_row = 0
        num_str = 1
        total_sum_incl_vat = 0
        total_sum_excl_vat = 0

        parse_table = True
        for rownum in range(start_row, sheet.nrows):

            row = sheet.row_values(rownum)
            if parse_table:
                if row[0] != '':
                    # товары
                    parent = self.tree.xpath(u'//Документ/ТаблСчФакт')
                    child = etree.SubElement(parent[0], u'СведТов')
                    child.set(u'НомСтр', str(num_str))
                    child.set(u'НаимТов', str(row[1]))
                    child.set(u'КолТов', str(int(row[4])))
                    child.set(u'ЦенаТов', str(row[5]).replace(",", "."))
                    sum_excl_vat = "".join(str(row[6]).replace(",", ".").split())
                    child.set(u'СтТовБезНДС', sum_excl_vat)
                    total_sum_excl_vat = total_sum_excl_vat + Decimal(sum_excl_vat)

                    sum_incl_vat = "".join(str(row[10]).replace(",", ".").split())
                    child.set(u'СтТовУчНал', sum_incl_vat)

                    total_sum_incl_vat = total_sum_incl_vat + Decimal(sum_incl_vat)

                    child.tail = "\n  "
                    child.text = "\n  "

                    # TODO акциз
                    subchild1 = etree.SubElement(child, u'Акциз')
                    subsubchild = etree.SubElement(subchild1, u'БезАкциз')
                    subsubchild.text = u'без акциза'
                    subchild1.tail = "\n  "

                    # TODO НДС
                    child.set(u'НалСт', str(row[8]))
                    subchild2 = etree.SubElement(child, u'СумНал')
                    subsubchild = etree.SubElement(subchild2, u'БезНДС')
                    subsubchild.text = u'без НДС'
                    subchild2.tail = "\n  "

                    subchild3 = etree.SubElement(child, u'ДопСведТов')
                    subchild3.set(u'ПрТовРаб', u'1')
                    subchild3.set(u'КодТов', str(int(row[0])))

                    # TODO единицы измерения
                    child.set(u'ОКЕИ_Тов', str(int(row[2])))
                    subchild3.set(u'НаимЕдИзм', u'шт')

                    # TODO Страна происхождения (если не Россия)
                    # subchild4 = etree.SubElement(child, u'СвТД')
                    # subchild4.set(u'КодПроисх', u'392')
                    # subchild4.set(u'НомерТД', u'10216120/200317/0015847/1')))
                    # subchild3.set(u'ДопСведТов', u'Япония')

                    subchild3.tail = "\n  "
                    num_str = num_str + 1
                else:
                    parse_table = False
                    total_row = rownum

        # Итого
        row = sheet.row_values(total_row)

        parent = self.tree.xpath(u'//Документ/ТаблСчФакт')
        child = etree.SubElement(parent[0], u'ВсегоОпл')
        child.set(u'СтТовБезНДСВсего', str(total_sum_excl_vat))
        child.set(u'СтТовУчНалВсего', str(total_sum_incl_vat))

        subchild = etree.SubElement(child, u'СумНалВсего')
        subsubchild = etree.SubElement(subchild, u'СумНал')
        subsubchild.text = '0.0'

    def save_file_to_output_folder(self):
        now = datetime.datetime.now()
        output_filename = self.filename+'_result_' + now.strftime("%d%m%y%H%M%S")
        print(settings.THIS_FOLDER, 'tech/output/' + output_filename + '.xml')
        self.tree.write(os.path.join(settings.THIS_FOLDER, 'tech/output/' + output_filename + '.xml'),
                        xml_declaration=True, encoding='windows-1251', pretty_print=True)

    def column_num(self, char):
        chars_upper = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S',
                       'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
                       'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP',
                       'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
                       'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP',
                       'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ']
        return (chars_upper.index(char))

