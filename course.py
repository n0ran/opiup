# coding=utf-8
__author__ = 'n0ran'


import xlwt

wbk = xlwt.Workbook()
sheet = wbk.add_sheet('sheet 1')
########################################################################################################
def sheetWriteStrWithDoc( row, func ):
    strdoc = str(func.__doc__).decode("utf-8")
    strval = func()
    if type(strval) == str:
        strval = strval.decode("utf-8")
    else:
        strval = str(func())#.decode("utf-8")
    print strdoc, strval
    sheet.write( row, 0, strdoc  )
    sheet.write( row, 1, strval  )
########################################################################################################
class Variant():
    def __init__(self):
        """"""
        """Время на каждую операцию"""
        self.TimeForOperations = [ 5.6, 9.0, 8.4, 4.0, 7.0, 6.0, 6.8, 7.2 ]
        """Наименование операций"""
        self.OperationsDescription = [
                                    "1. Фрезерная",
                                    "2. Шлифовальная",
                                    "3. Слесарная",
                                    "4. Токарная",
                                    "5. Фрезерная",
                                    "6. Слесарная",
                                    "7. Сверлильная",
                                    "8. Токарная"
                                    ]
        """Количество операций"""
        self.NumOfOperations = len(self.OperationsDescription)
        """Количество изделий за год"""
        self.Nz = 2403
        """Количество рабочих дней в месяце"""
        self.d = 20
        """Длительность смены, ч"""
        self.t = 8
        """Количество смен"""
        self.s = 2
        """Используется для расчета К по"""
        self.alpha = 3
    def getTimeForOperation(self, index):
        if index >= len( self.TimeForOperations ):
            return -1
        return  self.TimeForOperations[index]
    def getOperationDescription(self, index):
        if index >= len( self.TimeForOperations ):
            return -1
        return  self.OperationsDescription[index]
    def getOperationsNumber(self):
        """aka m"""
        return self.NumOfOperations
    def Kpo(self):
        """К по - Коэффициент, учитывающий время простоя оборудования в плановом ремонте"""
        return 1 - self.alpha*1.0/100
    def Fn(self):
        """F н. - номинальный фонд времени работы оборудования"""
        return self.d*self.t*self.s
    def Fe(self):
        """F э. - годовой (месячный) эффективный фонд времени работы оборудования"""
        return self.Fn() * self.Kpo()
    def Rnp(self):
        """r н.п. - такт выпуска изделий мин/шт"""
        return 60.0 * self.Fe() / self.Nz
    def Cpr(self):
        """C пр. - количество рабочих мест( единиц оборудования), необходимых для выполнения данного техн. процесса """
        cpr = 0
        for i in range(0, self.getOperationsNumber()):
            cpr += self.getTimeForOperation( i )
        cpr /= self.Rnp()
        return cpr
    def Ksp(self):
        """К сп. - Коэффициент специализации"""
        return 1.0*self.getOperationsNumber() / self.Cpr()
    def Km(self):
        """К м. - Коэффициент массовости"""
        km = 0
        for i in range(0, self.getOperationsNumber()):
            km += self.getTimeForOperation( i )
        km /= 1.0 * self.getOperationsNumber() * self.Rnp()
        return km
########################################################################################################
#variant data init

variant = Variant()
########################################################################################################
#Chapter 2.2.2
#Выбор и обоснование типа производства и вида поточной линии (участка)

def Formula1():
    """Формула 1 Коэффициент специализации Ксп
    """
    global variant
    return variant.Ksp()

def conclusion_Formula1( Ksp = Formula1()):
    """тип производства в зависимости от коэффициента специализации Ксп (формула 1)
    """
    threaded = False
    ret = "Тип производства (Ксп) "
    if Ksp <= 1:
        ret += "массовый"
        threaded = True
    elif Ksp > 1 and Ksp <= 10:
        ret += "крупносерийный"
        threaded = True
    elif Ksp > 10 and Ksp <= 20:
        ret += "среднесерийный"
        threaded = True
    elif Ksp > 20 and Ksp <= 40:
        ret += "мелкосерийный"
    else: #Ksp > 40
        ret += "единичное производство"
    ret += ". Для данного типа производства целесообразна организация "
    if threaded:
        ret += "поточного производства"
    else:
        ret += "предметно-замкнутого участка изготовления деталей или участка серийной сборки изделия"
    return ret
def Formula2( ):
    """Формула 2 Коэффициент массовости Км

    где  – норма штучного времени на i-й операции с учётом коэффициента выполнения норм времени (мин)
    m – количество операций по данному технологическому процессу;
    Rnp – такт выпуска изделий, определяется по формуле
    """
    global variant
    return variant.Km()

def conclusion_Formula2( Km = Formula2() ):
    """тип производства в зависимости от коэффициента массовости Км (формула 2)
    """
    ret = "Тип производства (Км) "
    if Km > 1:
        ret += "массовый"
    else: #Ksp < 1
        ret += "серийный"
    return ret
def Formula3():
    """Формула 3 Такт выпуска изделий r н.п. мин/шт

    Nz  – годовая (месячная) программа запускаемого изделия, шт.
    Fe– годовой (месячный) эффективный фонд времени работы оборудования
    """
    global variant
    return variant.Rnp()

def Formula4():
    """Формула 4 - годовой(месячный) эффективный фонд времени работы оборудования.
    """
    global variant
    return variant.Fe()

def Chapter1():
    """1. Выбор и обоснование типа производства и вида поточной линии (участка)
    """
    global sheet
    global variant
    sheet.write(0, 0, Chapter1.__doc__.decode("utf-8") )
    sheetWriteStrWithDoc( 1, variant.Ksp )
    sheetWriteStrWithDoc( 2, variant.Cpr )
    sheetWriteStrWithDoc( 3, variant.Rnp )
    sheetWriteStrWithDoc( 4, variant.Km  )
    sheetWriteStrWithDoc( 5, variant.Fn  )
    sheetWriteStrWithDoc( 6, variant.Kpo )
    sheetWriteStrWithDoc( 7, variant.Fe  )

    sheetWriteStrWithDoc( 9, Formula1  )
    sheetWriteStrWithDoc( 10, conclusion_Formula1)
    sheetWriteStrWithDoc( 11, Formula2  )
    sheetWriteStrWithDoc( 12, conclusion_Formula2)
    sheetWriteStrWithDoc( 13, Formula3  )
    sheetWriteStrWithDoc( 14, Formula4  )
########################################################################################################
# indexing is zero based, row then column
#sheet.write(0,1,'test text')
Chapter1()
wbk.save(r'w:\test2.xls')