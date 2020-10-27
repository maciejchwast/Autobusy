#Program do odczytywania godziny odjazdu nastepnego autobusu
#Napisany 27.10.2020 przez Maciej Chwast

from datetime import datetime                                   #Import modulu odpowiedzialnego za czas
def CzyJestSobota():                                            #Funkcja sprawdza czy jest sobota
    if(datetime.today().weekday()== 5):                         #datetime.today().weekday() zwraca wartosc liczbowa odpowiadajaca dniu tygodnia
        return True
    else:
        return False

def CzyJestNiedziela():                                         #Funkcja sprawdza czy jest niediezla
    if(datetime.today().weekday()==6):                          #Dzialanie wytlumaczone powyzej
        return True
    else:
        return False

def KtoraGodzina():                                             #Funkcja zwraca obecną datę
    teraz = datetime.now()
    return teraz

import xlrd                                                     #Import modulu odpowiedzialnego za operacje na excelu
def KiedyOdjedzie(dzien):                                       #Funkcja sprawdzajaca za ile odjedzie autobus. parametr dzien sluzy rozpoznawaniu dni roboczych sobot i niedziel
    teraz = KtoraGodzina()
    godzina = int(teraz.strftime("%H"))                         #zmienna godzina jest liczba wyciagnieta ze zmiennej teraz
    minuty = int(teraz.strftime("%M"))                          #j.w.
    loc = (r"D:\Programy\PythonApplication3\PythonApplication2\173_azory.xlsx")         #adres pliku z rozkladem autobusu
    wb = xlrd.open_workbook(loc)                                #otwarcie pliku exclela do pracy z nim
    sheet = wb.sheet_by_index(0)                                #ustalenie arkusza na ktorym pracujemy
    if(godzina <=23 and godzina >=4):                           #sprawdzenie czy w ogole jezdza autobusy, mozna to przeniesc poza ta funkcje
        mozliwosci = sheet.cell_value(godzina-2, 1+dzien)       #wyciagniecie wszystkich odjazdow o danej godzinie
        odjazdy = mozliwosci.split(" ")                         #Podzielenie wszystkich mozliwosci na tabele
        for odjazd in odjazdy:                                  #zmienna odjazd przyjmuje po kolei wszystkie wartosci tabeli odjazdy
            if(minuty < int(odjazd)):                           #jezeli znajdzie sie kurs o tej godzinie ktory jeszcze nie odjechal
                return int(odjazd) - minuty                     #program zwraca roznice obecnego czasu a godzina odjazdu
                exit(self)                                      #oraz wychodzi z funkcji
        mozliwosci = sheet.cell_value(godzina-1,1+dzien)        #jezeli nie ma juz o danej godzinie odjazdow program przechodzi do kolejnej godziny
        odjazdy = mozliwosci.split()
        return int(odjazdy[0]) + (60 - minuty)                  #i zwraca pierwsza mozliwość

    else:
        return "O tej godzinie nigdzie nie pojedziesz :("       #jezeli jest przed 4ta albo po 23 to nie ma autobusow
    

if(CzyJestSobota() == False and CzyJestNiedziela() == False):
   print("Masz autobus za " + str(KiedyOdjedzie(0) )+ " minut")

elif(CzyJestSobota() == True):
    print("Masz autobus za " + str(KiedyOdjedzie(1) )+ " minut")

elif(CzyJestNiedziela() == True):
    print("Masz autobus za " + str(KiedyOdjedzie(2) )+ " minut")