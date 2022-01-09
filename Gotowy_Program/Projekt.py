import re
import xlwt
import datetime
import sys
import os

Lista_Sformatowana = []
Lista_Gotowa = []

def Konwertowanie_Listy(Link_Do_Danych):
    assert os.path.exists(Link_Do_Danych), "I did not find the file at, "+str(Link_Do_Danych)
    Plik = open(Link_Do_Danych, 'r')

    Lista = Plik.readlines()

    X = 0

    for Linia in Lista:
        X += 0
        print("Linia{}: {}".format(X, Linia.strip()))

        if Linia.find('K:SALA ') != -1 or Linia.find('manipulatora ') != -1 or Linia.find('S:Magazyn   ') != -1 or Linia.find('zabroniony        ') != -1 or Linia.find('blokada ') != -1 or Linia.find('błędne ') != -1:
            print('True')
        else:
            Sformatowana_Linia = ""
            Sformatowana_Linia = re.sub('  ', ' ', Linia)
            Sformatowana_Linia = Sformatowana_Linia.replace(' Dostęp użytkownika               K:WEJŚCIE (20h)   U:', ' ')
            Sformatowana_Linia = Sformatowana_Linia.replace(' Dostęp użytkownika               LCD:WEJSCIE PRODUKCJ U:', ' ')
            Sformatowana_Linia = Sformatowana_Linia.replace('  ', ' ')
            Lista_Sformatowana.append(Sformatowana_Linia)
    
    for Tabela in Lista_Sformatowana:
        Wynik = Tabela.split(" ")
        Lista_Gotowa.append(Wynik)

def Zapis_do_TXT():
    Plik_Sformatowany = open('Sformatowane_dane.txt', 'a')
    for Linia in Lista_Sformatowana:
            print(Linia)
            Plik_Sformatowany.write(Linia)
    Print("Plik TXT został stworzony pomyślnie w folerze programu")

def Zapis_do_Excela():
    Linie = 0
    XML_Plik = xlwt.Workbook()
    Arkusz = XML_Plik.add_sheet('Odbicia Kart')
    for P in Lista_Sformatowana:

           
            Linie = Linie+1
            Kolumny = 0
            for Linia in P.split(" "):
                Arkusz.write(Linie,Kolumny,Linia)
                Kolumny = Kolumny+1 
                    
            XML_Plik.save('stdoutput.xls')
    Print("Plik XML został stworzony pomyślnie w folerze programu")

def Godziny_Pracownika(Pracownik):
    Godziny_Pracy = 0

    Pierwszy_znaleziony = 0
    for Wiersz in Lista_Gotowa:
        Pracownik_Z_Listy = Wiersz[5] + " " + Wiersz[6]
        if Pracownik_Z_Listy == Pracownik:
            if Pierwszy_znaleziony == 0:
                Godzina_Konca = datetime.datetime.strptime(Wiersz[3], '%H:%M')
                Pierwszy_znaleziony = 1
            else:
                Godzina_Startu = datetime.datetime.strptime(Wiersz[3], '%H:%M')
                Pierwszy_znaleziony = 0
                X = Godzina_Konca - datetime.datetime(1900, 1, 1)
                Y = Godzina_Startu - datetime.datetime(1900, 1, 1)
                Czas_Przebywania = X.total_seconds() - Y.total_seconds()
                print(X.total_seconds() ,'-',Y.total_seconds())
                Godziny_Pracy = Godziny_Pracy + Czas_Przebywania 
    Wynik = Godziny_Pracy / 60
    Wynik = Wynik / 60
    print(Pracownik , " " , Wynik, "H pracy")

def Menu():
    print("Opcje Menu:")
    print("1. Ładowanie Danych")
    print("2. Zapis danych do pliku TXT")
    print("3. Zapis danych do pliku XLM")
    print("4. Sprawdzanie godzin pracownika")
    print("5. Zamknięcie programu")
    Opcja = input("Wprowadz opcje: ")

    if Opcja == "1":
        Konwertowanie_Listy(input("Podaj lokalizacje pliku z DLOADX: "))
    if Opcja == "2":
        if Lista_Sformatowana != []:
            Zapis_do_TXT()
            Main()
        else:
            print("Zaiportuj najpierw dane za pomocą opci 1 w menu")
    if Opcja == "3":
        if Lista_Sformatowana != []:
            Zapis_do_Excela()
            Main()
        else:
            print("Zaiportuj najpierw dane za pomocą opci 1 w menu")
    if Opcja == "4":
        if Lista_Sformatowana != []:
            Godziny_Pracownika(input("Podaj pracownika"))
            input("Kliknij by przejść do menu")
            Main()
        else:
            print("Zaiportuj najpierw dane za pomocą opci 1 w menu")
    if Opcja == "5":
        os._exit(0)
    else:
        print("Zła opcja")
        Main()

def Main():
    try:
        Menu()
    except Exception as blad:
        print("Nastąpił błąd programu. Kod błędu: ", str(blad))
        Menu()
    

Main()
input("Czekaj")