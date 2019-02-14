from lxml import etree
from xlrd import open_workbook
from django.conf import settings
import os




class Upd:
    """docstring"""

    def __init__(self, template):
        """Constructor"""
        self.template = template
        self.tree = etree.parse(self.template)
        pass

    # Файл
    def set_file(self, id_file):
        # <Файл xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ИдФайл="ON_SCHFDOPPR___20190204_6de6a7f1-9816-4d9b-8973-8dfa68ccfc8c" ВерсФорм="5.01" ВерсПрог="1.0">
        self.tree.xpath(u'//Файл')[0].attrib[u'ИдФайл'] = id_file

    # Документ
    def set_document(self, function, date_inf_pr, time_inv_pr, naim_econ_sub_sost):
        # <Документ КНД="1115125" Функция="ДОП" ПоФактХЖ="Документ об отгрузке товаров (выполнении работ), передаче имущественных прав (документ об оказании услуг)" НаимДокОпр="Документ об отгрузке товаров (выполнении работ), передаче имущественных прав (Документ об оказании услуг)" ДатаИнфПр="09.08.2018" ВремИнфПр="13.22.35" НаимЭконСубСост="ООО &quot;Форси&quot;">
        self.tree.xpath(u'//Документ')[0].attrib[u'Функция'] = function
        self.tree.xpath(u'//Документ')[0].attrib[u'ДатаИнфПр'] = date_inf_pr
        self.tree.xpath(u'//Документ')[0].attrib[u'ВремИнфПр'] = time_inv_pr
        self.tree.xpath(u'//Документ')[0].attrib[u'НаимЭконСубСост'] = naim_econ_sub_sost

    # СвСчФакт
    def set_sv_sch_fact(self, sf_number, sf_date):
        # <СвСчФакт НомерСчФ="48" ДатаСчФ="07.02.2019" КодОКВ="643">
        self.tree.xpath(u'//Документ/СвСчФакт')[0].attrib[u'НомерСчФ'] = sf_number
        self.tree.xpath(u'//Документ/СвСчФакт')[0].attrib[u'ДатаСчФ'] = sf_date

    # СвЮЛУч
    def set_sv_prod_ul(self, org, innul, kpp):
        # <СвЮЛУч НаимОрг="ООО &quot;ФОРСИ&quot;" ИННЮЛ="5904650273" КПП="590401001" />
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/ИдСв/СвЮЛУч')[0].attrib[u'НаимОрг'] = org
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/ИдСв/СвЮЛУч')[0].attrib[u'ИННЮЛ'] = innul
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/ИдСв/СвЮЛУч')[0].attrib[u'КПП'] = kpp

    #
    def set_sv_prod_adr(self, index, kod_reg, city, street, house, korp):
        # <АдрРФ Индекс="614064" КодРегион="59" Город="Пермь" Улица="Чкалова" Дом="9" Корпус="Д" />
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/Адрес/АдрРФ')[0].attrib[u'Индекс'] = index
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/Адрес/АдрРФ')[0].attrib[u'КодРегион'] = kod_reg
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/Адрес/АдрРФ')[0].attrib[u'Город'] = city
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/Адрес/АдрРФ')[0].attrib[u'Улица'] = street
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/Адрес/АдрРФ')[0].attrib[u'Дом'] = house
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/Адрес/АдрРФ')[0].attrib[u'Корпус'] = korp

    # СвЮЛУч
    def set_sv_prod_phone(self, phone):
        # <Контакт Тлф="+79024786088" />
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/Контакт')[0].attrib[u'Тлф'] = phone

    # СвБанк
    def set_sv_prod_bank(self, schet, bank, bik, korrschet):
        # <СвБанк НаимБанк="Филиал ОАО &quot;Уралсиб&quot; в г.Уфа" БИК="048073770" КорСчет="30101810600000000770" />
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/БанкРекв')[0].attrib[u'НомерСчета'] = schet
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/БанкРекв/СвБанк')[0].attrib[u'НаимБанк'] = bank
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/БанкРекв/СвБанк')[0].attrib[u'БИК'] = bik
        self.tree.xpath(u'//Документ/СвСчФакт/СвПрод/БанкРекв/СвБанк')[0].attrib[u'КорСчет'] = korrschet



    # СвЮЛУч
    def set_gruz_ot_ul(self, org, innul, kpp):
        # <СвЮЛУч НаимОрг="ООО &quot;ФОРСИ&quot;" ИННЮЛ="5904650273" КПП="590401001" />
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/ИдСв/СвЮЛУч')[0].attrib[u'НаимОрг'] = org
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/ИдСв/СвЮЛУч')[0].attrib[u'ИННЮЛ'] = innul
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/ИдСв/СвЮЛУч')[0].attrib[u'КПП'] = kpp

    #
    def set_gruz_ot_adr(self, index, kod_reg, city, street, house, korp):
        # <АдрРФ Индекс="614064" КодРегион="59" Город="Пермь" Улица="Чкалова" Дом="9" Корпус="Д" />
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/Адрес/АдрРФ')[0].attrib[u'Индекс'] = index
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/Адрес/АдрРФ')[0].attrib[u'КодРегион'] = kod_reg
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/Адрес/АдрРФ')[0].attrib[u'Город'] = city
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/Адрес/АдрРФ')[0].attrib[u'Улица'] = street
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/Адрес/АдрРФ')[0].attrib[u'Дом'] = house
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/Адрес/АдрРФ')[0].attrib[u'Корпус'] = korp

    # СвЮЛУч
    def set_gruz_ot_phone(self, phone):
        # <Контакт Тлф="+79024786088" />
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/Контакт')[0].attrib[u'Тлф'] = phone

    # СвБанк
    def set_gruz_ot_bank(self, schet, bank, bik, korrschet):
        # <СвБанк НаимБанк="Филиал ОАО &quot;Уралсиб&quot; в г.Уфа" БИК="048073770" КорСчет="30101810600000000770" />
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/БанкРекв')[0].attrib[u'НомерСчета'] = schet
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/БанкРекв/СвБанк')[0].attrib[u'НаимБанк'] = bank
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/БанкРекв/СвБанк')[0].attrib[u'БИК'] = bik
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/БанкРекв/СвБанк')[0].attrib[u'КорСчет'] = korrschet



    # СвЮЛУч
    def set_gruz_pol_ul(self, org, innul, kpp):
        # <СвЮЛУч НаимОрг="ООО &quot;ФОРСИ&quot;" ИННЮЛ="5904650273" КПП="590401001" />
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/ИдСв/СвЮЛУч')[0].attrib[u'НаимОрг'] = org
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/ИдСв/СвЮЛУч')[0].attrib[u'ИННЮЛ'] = innul
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/ИдСв/СвЮЛУч')[0].attrib[u'КПП'] = kpp

    #
    def set_gruz_pol_adr(self, index, kod_reg, city, street, house, korp):
        # <АдрРФ Индекс="614064" КодРегион="59" Город="Пермь" Улица="Чкалова" Дом="9" Корпус="Д" />
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/Адрес/АдрРФ')[0].attrib[u'Индекс'] = index
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/Адрес/АдрРФ')[0].attrib[u'КодРегион'] = kod_reg
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/Адрес/АдрРФ')[0].attrib[u'Город'] = city
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/Адрес/АдрРФ')[0].attrib[u'Улица'] = street
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/Адрес/АдрРФ')[0].attrib[u'Дом'] = house
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/Адрес/АдрРФ')[0].attrib[u'Корпус'] = korp

    # СвЮЛУч
    def set_gruz_pol_phone(self, phone):
        # <Контакт Тлф="+79024786088" />
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузОт/ГрузОтпр/Контакт')[0].attrib[u'Тлф'] = phone

    # СвБанк
    def set_gruz_pol_bank(self, schet, bank, bik, korrschet):
        # <СвБанк НаимБанк="Филиал ОАО &quot;Уралсиб&quot; в г.Уфа" БИК="048073770" КорСчет="30101810600000000770" />
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/БанкРекв')[0].attrib[u'НомерСчета'] = schet
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/БанкРекв/СвБанк')[0].attrib[u'НаимБанк'] = bank
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/БанкРекв/СвБанк')[0].attrib[u'БИК'] = bik
        self.tree.xpath(u'//Документ/СвСчФакт/ГрузПолуч/БанкРекв/СвБанк')[0].attrib[u'КорСчет'] = korrschet



    # СвЮЛУч
    def set_sv_pokup_ul(self, org, innul, kpp):
        # <СвЮЛУч НаимОрг="ООО &quot;ФОРСИ&quot;" ИННЮЛ="5904650273" КПП="590401001" />
        self.tree.xpath(u'//Документ/СвСчФакт/СвПокуп/ИдСв/СвЮЛУч')[0].attrib[u'НаимОрг'] = org
        self.tree.xpath(u'//Документ/СвСчФакт/СвПокуп/ИдСв/СвЮЛУч')[0].attrib[u'ИННЮЛ'] = innul
        self.tree.xpath(u'//Документ/СвСчФакт/СвПокуп/ИдСв/СвЮЛУч')[0].attrib[u'КПП'] = kpp


    #TODO переделать населенный пункт и район
    def set_sv_pokup_adr(self, index, kod_reg, rayon, nas_punkt, house):
        # <АдрРФ Индекс="614064" КодРегион="59" Город="Пермь" Улица="Чкалова" Дом="9" Корпус="Д" />
        self.tree.xpath(u'//Документ/СвСчФакт/СвПокуп/Адрес/АдрРФ')[0].attrib[u'Индекс'] = index
        self.tree.xpath(u'//Документ/СвСчФакт/СвПокуп/Адрес/АдрРФ')[0].attrib[u'КодРегион'] = kod_reg
        self.tree.xpath(u'//Документ/СвСчФакт/СвПокуп/Адрес/АдрРФ')[0].attrib[u'Район'] = rayon
        self.tree.xpath(u'//Документ/СвСчФакт/СвПокуп/Адрес/АдрРФ')[0].attrib[u'НаселПункт'] = nas_punkt
        self.tree.xpath(u'//Документ/СвСчФакт/СвПокуп/Адрес/АдрРФ')[0].attrib[u'Дом'] = house


    def set_dop_sv_fhzh1(self, naim_okv):
        # <ДопСвФХЖ1 НаимОКВ="Российский рубль" />
        self.tree.xpath(u'//Документ/СвСчФакт/ДопСвФХЖ1')[0].attrib[u'НаимОКВ'] = naim_okv

    def set_table(self, filename):
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
                    child.set(u'ЦенаТов', str(row[5]))
                    sum_excl_vat = "".join(str(row[6]).replace(",", ".").split())
                    child.set(u'СтТовБезНДС', sum_excl_vat)
                    total_sum_excl_vat = total_sum_excl_vat + float(sum_excl_vat)

                    sum_incl_vat = "".join(str(row[10]).replace(",", ".").split())
                    child.set(u'СтТовУчНал', sum_incl_vat)

                    total_sum_incl_vat = total_sum_incl_vat + float(sum_incl_vat)

                    child.tail = "\n  "
                    child.text = "\n  "

                    #TODO акциз
                    subchild1 = etree.SubElement(child, u'Акциз')
                    subsubchild = etree.SubElement(subchild1, u'БезАкциз')
                    subsubchild.text = u'без акциза'
                    subchild1.tail = "\n  "

                    #TODO НДС
                    child.set(u'НалСт', str(row[8]))
                    subchild2 = etree.SubElement(child, u'СумНал')
                    subsubchild = etree.SubElement(subchild2, u'БезНДС')
                    subsubchild.text = u'без НДС'
                    subchild2.tail = "\n  "

                    subchild3 = etree.SubElement(child, u'ДопСведТов')
                    subchild3.set(u'ПрТовРаб', u'1')
                    subchild3.set(u'КодТов', str(int(row[0])))

                    #TODO единицы измерения
                    child.set(u'ОКЕИ_Тов', str(int(row[2])))
                    subchild3.set(u'НаимЕдИзм', u'шт')

                    #TODO Страна происхождения (если не Россия)
                    #subchild4 = etree.SubElement(child, u'СвТД')
                    #subchild4.set(u'КодПроисх', u'392')
                    #subchild4.set(u'НомерТД', u'10216120/200317/0015847/1')))
                    #subchild3.set(u'ДопСведТов', u'Япония')

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

    def output(self, file_name):

        self.tree.write(os.path.join(settings.THIS_FOLDER, 'tech/output/' + file_name + '.xml'),
                        xml_declaration=True, encoding='windows-1251', pretty_print=True)
